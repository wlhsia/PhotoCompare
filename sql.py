import sqlite3

class User(object):
    def __init__(self,dbname):
        self.cx = sqlite3.connect(dbname,check_same_thread=False)
        self.cu = self.cx.cursor()

    def getUserPassword(self,usernx):
        sql = "select password from 'User' where username=\'"+usernx+"\'"
        self.cu.execute(sql)
        ress = self.cu.fetchone()
        if ress:
            return ress
        else:
            return ()

    def getAllUser(self):
        sql = "select username, password from 'User'"
        self.cu.execute(sql)
        ress = self.cu.fetchall()
        lists = []
        for res in ress:
            lists.append({"username": res[0], "password": res[1]})
        return lists