from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import sys
import uuid
import cv2
import numpy as np
import pandas as pd
from pdf2image import convert_from_path
from sklearn.cluster import KMeans
from openpyxl import load_workbook
from openpyxl.styles import Font,Alignment,Border,Side
from openpyxl.drawing.image import Image as Image_xl
from PIL import Image
import itertools
from functools import reduce
import shutil
from sklearn.externals import joblib
from numba import cuda
from datetime import datetime
import tempfile
import docx
from xml.etree import ElementTree
import xml.etree.cElementTree as ET
from io import StringIO
import sqlite3
import imagehash


app = Flask(__name__)
APP_ROOT = os.path.dirname(os.path.abspath(__file__))
CORS(app)

dbPath = sys.path[0]+'/database.db'
cx = sqlite3.connect(dbPath, check_same_thread=False)

@app.route("/login", methods=['POST'])
def login():
    response = {"success": False}
    if request.method == 'POST' and request.form.get('username') and request.form.get('password'):
        data = request.form.to_dict()
        username = data.get("username")
        password = data.get("password")
        cu = cx.cursor()
        try:
            cu.execute("select password from 'User' where username=\'" + username + "\'")
            pwd = cu.fetchone()[0]
        except:
            pwd = None
        cu.close()
        if not pwd:
            return "-1"
        elif password != pwd:
            return "0"
        elif username == "admin" and password == pwd:
            return "admin"
        else:
            return "1"
    else:
        return jsonify(response)

@app.route("/users", methods=["GET" , "POST"])
def users():
    response_object = {'status': 'success'}
    if request.method == 'GET':
        cu = cx.cursor()
        cu.execute("select username, password from 'User'")
        data = cu.fetchall()
        cu.close()
        userList = []
        for d in data:
            userList.append({"username": d[0], "password": d[1]})
        response_object['users'] = userList
    if request.method == 'POST':
        post_data = request.get_json()
        username = post_data.get('username')
        password = post_data.get('password')
        cu = cx.cursor()
        cu.execute("insert into User values(null, '" + username + "', '" + password + "')")
        cu.close()
        cx.commit()
        response_object['message'] = 'User added!'

    return jsonify(response_object)

@app.route('/user/<user_id>', methods=['PUT', 'DELETE'])
def user(user_id):
    response_object = {'status': 'success'}
    if request.method == 'PUT':
        post_data = request.get_json()
        username = post_data.get('username')
        password = post_data.get('password')
        cu = cx.cursor()
        cu.execute("UPDATE User SET password = '" + password + "' WHERE username = '" + username + "'" )
        cu.close()
        cx.commit()
        response_object['message'] = 'User updated!'
    if request.method == 'DELETE':
        cu = cx.cursor()
        cu.execute("DELETE FROM User WHERE username = '" + user_id + "'")
        cu.close()
        cx.commit()
        response_object['message'] = 'User removed!'
    return jsonify(response_object)


@app.route('/api/upload', methods=['POST'])
def upload():
    response = {"success": False}
    try:
        if request.method == 'POST':
            files=request.files
            if len(files) != 0:
                autoGenFolderName = uuid.uuid1()
                folderPath = os.path.join(APP_ROOT, str(autoGenFolderName))
                os.mkdir(folderPath)
            for f in files:
                file = request.files[f]
                filename = secure_filename(file.filename)
                file.save(os.path.join(folderPath, filename))
            response = {
                "success": True,
                "folderPath": folderPath,
            }
    except Exception as e:
        print(e)

    return jsonify(response)

@app.route('/api/delete', methods=['POST'])
def delete():
    print(request.get_json())
    dcit_request = request.get_json()
    path = os.path.join(dcit_request['folderPath'],dcit_request['file'])
    os.remove(path)
    fileList = os.listdir(dcit_request['folderPath'])
    if len(fileList)==0:
        shutil.rmtree(dcit_request['folderPath'])
    response = {"success": True}
    return jsonify(response)

@app.route('/api/getlist', methods=['POST'])
def uploadList():
    print(request.get_json())
    dcit_request = request.get_json()
    folderPath = dcit_request['folderPath']
    if os.path.isdir(folderPath):
        fileList = os.listdir(folderPath)
    else: 
        fileList= []
    response = {
        "fileList": fileList,
    }
    return jsonify(response)

@app.route('/download', methods=['GET'])
def download():
    f = open('./比對結果.xlsx', 'rb') 
    return f.read()


pdfsPath='./pdfs'
wordsPath='./words'
excelsPath='./excels'
rawImgsPath='./raw_imgs'
resizeImgsPath = './resize_imgs'

@app.route('/api/compare', methods=['POST'])
def compare():

    dcit_request = request.get_json()
    folderPath = dcit_request['folderPath']
    files=os.listdir(folderPath)
    imgsPath=folderPath+'_imgs'
    if not os.path.isdir(imgsPath):
        os.mkdir(imgsPath)

    # 擷取相片
    filesName = []
    for file in files:
        if str(file).split('.')[-1].lower() == 'docx':
            getWordImgs(folderPath,file,imgsPath)
            shutil.move(os.path.join(folderPath,file),os.path.join(wordsPath,file))
        elif str(file).split('.')[-1].lower() == 'xlsx':
            getExcelImgs(folderPath,file,imgsPath)
            shutil.move(os.path.join(folderPath,file),os.path.join(excelsPath,file)) 
        fileName = str(file).split('_')[0]
        if fileName not in filesName:
            filesName.append(fileName)
    shutil.rmtree(folderPath)

    
    # 上傳相片倆倆比對
    imgsName = os.listdir(imgsPath)
    result1 = []
    duplicateImgs = []    
    for i,(imgName1,imgName2) in enumerate(itertools.combinations(imgsName, 2)):
        imghash1 = imagehash.phash(Image.open(os.path.join(imgsPath,imgName1)))
        imghash2 = imagehash.phash(Image.open(os.path.join(imgsPath,imgName2)))
        similary = 1 - (imghash1 - imghash2)/len(imghash1.hash)**2
        if similary > 0.9:
            if imgName1 not in duplicateImgs:
                result1.append({'imgName1': imgName1, 'imgName2': imgName2, 'similary': similary})
                duplicateImgs.append(imgName2)



    for imgName in imgsName:
        shutil.move(os.path.join(imgsPath,imgName),os.path.join(rawImgsPath,imgName))
    shutil.rmtree(imgsPath)

    # 上傳相片與資料庫相片比對
    result2 = []
    cu = cx.cursor()
    cu.execute("select imageName from 'Image'")
    data = cu.fetchall()
    cu.close()
    dbImgsName = []
    for d in data:
        dbImgsName.append(d[0])
    for imgName1 in imgsName:
        imghash1 = imagehash.phash(Image.open(os.path.join(rawImgsPath,imgName1)))
        for imgName2 in dbImgsName:
            imghash2 = imagehash.phash(Image.open(os.path.join(rawImgsPath,imgName2)))
            similary = 1 - (imghash1 - imghash2)/len(imghash1.hash)**2
            if similary > 0.9:
                result2.append({'imgName1': imgName1, 'imgName2': imgName2, 'similary': similary})
                if imgName1 not in duplicateImgs:
                    duplicateImgs.append(imgName1)
    

    wb = load_workbook(filename= './比對結果.xlsx')
    sht = wb['工作表1']
    wb.remove(sht)
    wb.create_sheet("工作表1", 0)
    sht = wb['工作表1']
    sht.merge_cells('A1:D1')
    sht['A1'] = '工程案件相片重複性辨識'
    sht['A1'].font = Font(size=16, b=True, underline='single')
    sht['A1'].alignment = Alignment(horizontal='center', vertical='center')
    proj_num = ",".join(filesName)
    sht['A2'] = '工程編號：'+ proj_num
    sht['C2'] = '日期：' + datetime.now().strftime("%Y/%m/%d")
    sht['A2'].font = Font(size=14, b=True)
    sht['C2'].font = Font(size=14, b=True)
    sht.column_dimensions["A"].width = 50
    sht.column_dimensions["B"].width = 20
    sht.column_dimensions["C"].width = 50
    sht.column_dimensions["D"].width = 20
    
    i = 3
    if len(result1) !=0:
        sht['A'+str(i)] = '上傳的相片'
        sht['A'+str(i)].font = Font(b=True)
        sht['A'+str(i)].alignment = Alignment(horizontal='center', vertical='center')
        sht['C'+str(i)] = '上傳的相片'
        sht['C'+str(i)].font = Font(b=True)
        sht['C'+str(i)].alignment = Alignment(horizontal='center', vertical='center')
        i = i+1 
        for item in result1:
            sht.row_dimensions[i].height = 80
            with Image.open(os.path.join(rawImgsPath,item['imgName1'])) as img:
                img = img.resize((100,100),Image.ANTIALIAS)
                img.save(os.path.join(resizeImgsPath,item['imgName1']))
            with Image.open(os.path.join(rawImgsPath,item['imgName2'])) as img:
                img = img.resize((100,100),Image.ANTIALIAS)
                img.save(os.path.join(resizeImgsPath,item['imgName2']))
            sht['A'+str(i)] = item['imgName1']
            sht.add_image(Image_xl(os.path.join(resizeImgsPath,item['imgName1'])),"B"+str(i))
            sht['C'+str(i)] = item['imgName2']
            sht.add_image(Image_xl(os.path.join(resizeImgsPath,item['imgName2'])),"D"+str(i))
            i = i + 1 

    if len(result2) !=0:
        sht['A'+str(i)] = '上傳的相片'
        sht['A'+str(i)].font = Font(b=True)
        sht['A'+str(i)].alignment = Alignment(horizontal='center', vertical='center')
        sht['C'+str(i)] = '資料庫的相片'
        sht['C'+str(i)].font = Font(b=True)
        sht['C'+str(i)].alignment = Alignment(horizontal='center', vertical='center')
        i = i+1 
        for item in result2:
            sht.row_dimensions[i].height = 80
            with Image.open(os.path.join(rawImgsPath,item['imgName1'])) as img:
                img = img.resize((100,100),Image.ANTIALIAS)
                img.save(os.path.join(resizeImgsPath,item['imgName1']))
            with Image.open(os.path.join(rawImgsPath,item['imgName2'])) as img:
                img = img.resize((100,100),Image.ANTIALIAS)
                img.save(os.path.join(resizeImgsPath,item['imgName2']))
            sht['A'+str(i)] = item['imgName1']
            sht.add_image(Image_xl(os.path.join(resizeImgsPath,item['imgName1'])),"B"+str(i))
            sht['C'+str(i)] = item['imgName2']
            sht.add_image(Image_xl(os.path.join(resizeImgsPath,item['imgName2'])),"D"+str(i))
            i = i + 1 

    if len(duplicateImgs)==0:
        message = [
            '系統比對相片數量： ' + str(len(imgsName)) + ' 張，比對結果無重複相片',
            '是否將 ' + str(len(imgsName)) + ' 張相片寫入資料庫？(寫入後才能出表)'
        ]
    elif len(imgsName) == len(duplicateImgs):
        message = [
            '系統比對相片數量： ' + str(len(imgsName)) + ' 張，全部相片重複',
            ''
        ]
    else:
        message = [
            '系統比對相片數量： ' + str(len(imgsName)) + ' 張，重複相片： ' + str(len(duplicateImgs)) + ' 張',
            '是否將 ' + str(len(imgsName) - len(duplicateImgs)) + ' 張相片寫入資料庫？'
        ]

    wb.save('./比對結果.xlsx')
    wb.close()

    s1 = set(imgsName)
    s2 = set(duplicateImgs)
    nonDuplicateImgs= list(s1.symmetric_difference(s2))

        

    response = {"success": True, "result1": result1, "result2": result2, "message": message, "nonDuplicateImgs": nonDuplicateImgs}
    return jsonify(response)

def getWordImgs(folderPath, file, imgsPath):
    doc = docx.Document(os.path.join(folderPath,file))
    tables = doc.tables
    for page,table in enumerate(tables):
        xml_str=table._element.xml
        root= ET.fromstring(xml_str)
        namespaces = dict([node for _, node in ElementTree.iterparse(StringIO(xml_str), events=['start-ns'])])
        i=1
        for blip_elem in root.findall('.//a:blip', namespaces):
            embed_attr = blip_elem.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            img=doc.part.related_parts[embed_attr]
            page_num = (str(page+1)).zfill(3)
            dst_filenm = '_'.join([file,'Page'+page_num,str(i)])+'.jpg'
            with open(os.path.join(imgsPath,dst_filenm),'wb') as f:
                f.write(img.blob)
            i+=1
        for imagedata_elem in root.findall('.//v:imagedata', namespaces):
            id_attr = imagedata_elem.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            img=doc.part.related_parts[id_attr]
            page_num = (str(page+1)).zfill(3)
            dst_filenm = '_'.join([file,'Page'+page_num,str(i)])+'.jpg'
            with open(os.path.join(imgsPath,dst_filenm),'wb') as f:
                f.write(img.blob)
            i+=1

def getExcelImgs(folderPath, file, imgsPath):
    wb = load_workbook(os.path.join(folderPath,file))
    sheets = wb.get_sheet_names()
    for sheet in sheets:
        ws = wb[sheet]
        for i, image in enumerate(ws._images):
            img = Image.open(image.ref)
            imgFileName = '_'.join([file, sheet, str(i+1)])+'.jpg'
            img.save(os.path.join(imgsPath, imgFileName))

app.run()