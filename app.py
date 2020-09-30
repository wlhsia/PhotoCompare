from flask import Flask, jsonify, request
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import uuid
import cv2
import numpy as np
import pandas as pd
from pdf2image import convert_from_path
import json
from sklearn.cluster import KMeans
from openpyxl import load_workbook
from openpyxl.styles import Font,Alignment,Border,Side
from openpyxl.drawing.image import Image as Image_xl
from PIL import Image
import itertools
from functools import reduce
import shutil
from sklearn.externals import joblib
import keras
from numba import cuda
from datetime import datetime
import tempfile
import docx
from xml.etree import ElementTree
import xml.etree.cElementTree as ET
from io import StringIO
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
APP_ROOT = os.path.dirname(os.path.abspath(__file__))
CORS(app)


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

pdfsPath='./pdfs'
wordsPath='./words'
allImgsPath='./imgs'
resizeImgsPath = './resize_imgs'

@app.route('/api/compare', methods=['POST'])
def compare():

    dcit_request = request.get_json()
    folderPath = dcit_request['folderPath']
    files=os.listdir(folderPath)
    imgsPath=folderPath+'_imgs'
    if not os.path.isdir(imgsPath):
        os.mkdir(imgsPath)

    filesName = []
    for file in files:
        if str(file).split('.')[-1].lower() == 'pdf':
            cropImgs(folderPath,file,imgsPath)
            shutil.move(os.path.join(folderPath,file),os.path.join(pdfsPath,file))
        elif str(file).split('.')[-1].lower() == 'docx':
            getWordImgs(folderPath,file,imgsPath)
            shutil.move(os.path.join(folderPath,file),os.path.join(wordsPath,file))
        fileName = str(file).split('_')[0]
        if fileName not in filesName:
            filesName.append(fileName)
    shutil.rmtree(folderPath)

    # 將照片轉為特徵
    imgsName = os.listdir(imgsPath)
    dictPHash = {}
    for imgName in imgsName:
        img=Image.open(os.path.join(imgsPath,imgName))
        dictPHash[imgName] = phash(img)
        shutil.move(os.path.join(imgsPath,imgName),os.path.join(allImgsPath,imgName))
    shutil.rmtree(imgsPath)

    # 上傳相片倆倆比對
    result1 = []
    tempImgName = []        
    for i,(imgName1,imgName2) in enumerate(itertools.combinations(imgsName, 2)):
        distance = bin(dictPHash[imgName1] ^ dictPHash[imgName2]).count('1')
        similaryPHash = 1 - distance / 1024
        if similaryPHash > 0.9:
            if imgName1 not in tempImgName:
                result1.append({'imgName1': imgName1, 'imgName2': imgName2, 'imgPHash2':  dictPHash[imgName2], 'similaryPHash': similaryPHash})
                tempImgName.append(imgName2)


    wb = load_workbook(filename= './比對結果.xlsx')
    sht = wb['工作表1']
    wb.remove(sht)
    wb.create_sheet("工作表1", 0)
    sht = wb['工作表1']
    sht.merge_cells('A1:E1')
    sht['A1'] = '工程案件相片重複性辨識'
    sht['A1'].font = Font(size=16, b=True, underline='single')
    sht['A1'].alignment = Alignment(horizontal='center', vertical='center')
    proj_num = ",".join(filesName)
    sht['A2'] = '工程編號：'+ proj_num
    sht['E2'] = '日期：' + datetime.now().strftime("%Y/%m/%d")
    sht['A2'].font = Font(size=14, b=True)
    sht['E2'].font = Font(size=14, b=True)
    sht.column_dimensions["A"].width = 40
    sht.column_dimensions["B"].width = 40
    sht.column_dimensions["D"].width = 25
    sht.column_dimensions["E"].width = 25

    for i,item in enumerate(result1):
        sht.row_dimensions[i+4].height = 70
        with Image.open(os.path.join(allImgsPath,item['imgName1'])) as img:
            img = img.resize((160,90),Image.ANTIALIAS)
            img.save(os.path.join(resizeImgsPath,item['imgName1']))
        with Image.open(os.path.join(allImgsPath,item['imgName2'])) as img:
            img = img.resize((160,90),Image.ANTIALIAS)
            img.save(os.path.join(resizeImgsPath,item['imgName2']))
        sht['A'+str(i+4)] = item['imgName1']
        sht['B'+str(i+4)] = item['imgName2']
        sht.add_image(Image_xl(os.path.join(resizeImgsPath,item['imgName1'])),"D"+str(i+4))
        sht.add_image(Image_xl(os.path.join(resizeImgsPath,item['imgName2'])),"E"+str(i+4))    

    wb.save('./比對結果.xlsx')
    wb.close()

    response = {"success": True, "result":result1}
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

def cropImgs(folderPath,pdf,imgsPath):
    graph = tf.get_default_graph()
    model = unet()
    model.load_weights(os.path.join('./models','unet.hdf5'))
    # model = unet()
    # model.load_weights(os.path.join(model_path,'unet.hdf5'))
    with tempfile.TemporaryDirectory(dir= 'D:/temp') as path:
        page_imgs = convert_from_path(os.path.join(folderPath,pdf), output_folder=path, dpi=600)
        for page_number,page_img in enumerate(page_imgs):
            if page_img.size[0] < page_img.size[1]:
                page_img = page_img.rotate(90,Image.NEAREST,expand =True)
            page_img = np.array(page_img)
            rgb = cv2.cvtColor(page_img,cv2.COLOR_BGR2RGB)
            hsv = cv2.cvtColor(page_img,cv2.COLOR_BGR2HSV)
            gray = cv2.cvtColor(page_img,cv2.COLOR_BGR2GRAY)
            #unet
            img = cv2.resize(gray, (256, 256))
            img = np.reshape(img,img.shape+(1,)) if (not False) else img
            img = np.reshape(img,(1,)+img.shape)
            with graph.as_default():
                result=model.predict(img)
            # result = model._make_predict_function(img)
            result=result[0]
            img = labelVisualize(2,COLOR_DICT,result) if False else result[:,:,0]
            img = cv2.resize(img, (page_img.shape[1],page_img.shape[0]))
            img = (img*255).astype(np.uint8)
            (thresh, im_bw) = cv2.threshold(img, 0.05, 255, 0)
            contours, hierarchy = cv2.findContours(im_bw, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
            unet_conts = [contour for contour in contours if cv2.contourArea(contour)> 500000]
            #sat
            # 取出飽和度
            saturation = hsv[:,:,1]
            _, threshold = cv2.threshold(saturation, 1, 255.0, cv2.THRESH_BINARY)
            # 2值化圖去除雜訊
            kernel_radius = int(threshold.shape[1]/300)
            kernel = np.ones((kernel_radius, kernel_radius), np.uint8)
            threshold = cv2.morphologyEx(threshold,cv2.MORPH_OPEN,kernel)
            # 產生等高線
            contours, hierarchy = cv2.findContours(threshold, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
            sat_conts = [contour for contour in contours if cv2.contourArea(contour)> 500000]

            if len(unet_conts) == 6:
                conts=unet_conts
            elif len(sat_conts) == 6:
                conts=sat_conts
            elif len(sat_conts) > len(unet_conts):
                conts=sat_conts
            else:
                conts=unet_conts

            sortY_conts = sorted([cont for cont in conts],key = lambda x:x[0][0][1],reverse=False)
            up_conts = sortY_conts[:3]
            up_conts = sorted([cont for cont in up_conts],key = lambda x:x[0][0][0],reverse=False)
            down_conts = sortY_conts[3:]
            down_conts = sorted([cont for cont in down_conts],key = lambda x:x[0][0][0],reverse=False)
            merge_conts = up_conts+down_conts

            for i,c in enumerate(merge_conts):
                # 嘗試在各種角度，以最小的方框包住面積最大的等高線區域，以紅色線條標示
                rect = cv2.minAreaRect(c)
                box = cv2.boxPoints(rect)
                box = np.int0(box) 
                angle = rect[2]
                if angle < -45:
                    angle = 90 + angle
                # 以影像中心為旋轉軸心
                (h, w) = page_img.shape[:2]
                center = (w // 2, h // 2)
                # 計算旋轉矩陣
                M = cv2.getRotationMatrix2D(center, angle, 1.0)
                # 旋轉圖片
                rotated = cv2.warpAffine(rgb, M, (w, h),flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_CONSTANT)
                # 旋轉紅色方框座標
                pts = np.int0(cv2.transform(np.array([box]), M))[0]
                #  計算旋轉後的紅色方框範圍
                y_min = min(pts[0][0], pts[1][0], pts[2][0], pts[3][0])
                y_max = max(pts[0][0], pts[1][0], pts[2][0], pts[3][0])
                x_min = min(pts[0][1], pts[1][1], pts[2][1], pts[3][1])
                x_max = max(pts[0][1], pts[1][1], pts[2][1], pts[3][1])
                # 裁切影像
                img_crop = rotated[x_min:x_max, y_min:y_max]
                page_num = (str(page_number+1)).zfill(3)
                dst_filenm = '_'.join([pdf,'Page'+page_num,str(i+1)])+'.jpg'
                cv2.imwrite(os.path.join(imgsPath,dst_filenm),img_crop)


def unet(pretrained_weights = None,input_size = (256,256,1)):
    inputs = Input(input_size)
    conv1 = Conv2D(64, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(inputs)
    conv1 = Conv2D(64, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv1)
    pool1 = MaxPooling2D(pool_size=(2, 2))(conv1)
    conv2 = Conv2D(128, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(pool1)
    conv2 = Conv2D(128, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv2)
    pool2 = MaxPooling2D(pool_size=(2, 2))(conv2)
    conv3 = Conv2D(256, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(pool2)
    conv3 = Conv2D(256, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv3)
    pool3 = MaxPooling2D(pool_size=(2, 2))(conv3)
    conv4 = Conv2D(512, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(pool3)
    conv4 = Conv2D(512, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv4)
    drop4 = Dropout(0.5)(conv4)
    pool4 = MaxPooling2D(pool_size=(2, 2))(drop4)

    conv5 = Conv2D(1024, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(pool4)
    conv5 = Conv2D(1024, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv5)
    drop5 = Dropout(0.5)(conv5)

    up6 = Conv2D(512, 2, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(UpSampling2D(size = (2,2))(drop5))
    merge6 = concatenate([drop4,up6],axis=3) 
    conv6 = Conv2D(512, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(merge6)
    conv6 = Conv2D(512, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv6)

    up7 = Conv2D(256, 2, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(UpSampling2D(size = (2,2))(conv6))
    merge7 = concatenate([conv3,up7],axis=3) 
    conv7 = Conv2D(256, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(merge7)
    conv7 = Conv2D(256, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv7)

    up8 = Conv2D(128, 2, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(UpSampling2D(size = (2,2))(conv7))
    merge8 = concatenate([conv2,up8],axis=3) 
    conv8 = Conv2D(128, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(merge8)
    conv8 = Conv2D(128, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv8)

    up9 = Conv2D(64, 2, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(UpSampling2D(size = (2,2))(conv8))
    merge9 = concatenate([conv1,up9],axis=3) 
    conv9 = Conv2D(64, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(merge9)
    conv9 = Conv2D(64, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv9)
    conv9 = Conv2D(2, 3, activation = 'relu', padding = 'same', kernel_initializer = 'he_normal')(conv9)
    conv10 = Conv2D(1, 1, activation = 'sigmoid')(conv9)

    model = Model(input = inputs, output = conv10)

    model.compile(optimizer = Adam(lr = 1e-4), loss = 'binary_crossentropy', metrics = ['accuracy'])
    
    if(pretrained_weights):
    	model.load_weights(pretrained_weights)

    return model

def phash(img):
    img = img.resize((32,32), Image.ANTIALIAS).convert('L')
    avg = reduce(lambda x, y: x + y, img.getdata()) / 1024.
    hash_value=reduce(lambda x, y: x | (y[1] << y[0]), enumerate(map(lambda i: 0 if i < avg else 1, img.getdata())), 0)
    return hash_value

# if __name__ == '__main__':
#     app.run(debug=True)
app.run()