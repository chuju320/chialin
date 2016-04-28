import pymysql

class dbClass:
    def __init__(self,host,user,passwd,db,port=3306,charset='utf8'):
        self.host = host
        self.user = user
        self.passwd = passwd
        self.db = db
        self.port =port
        self.charset = charset
    def update(self,sql):
        pass
    def insert(self):
        pass
    def select(self):
        pass
    def delete(self,sql):
        conn = pymysql.connect(host = self.host , user = self.user, passwd = self.passwd , db = self.db , port = self.port, charset =self.charset)
        cur = conn.cursor()
        sta = cur.execute(sql)
        conn.commit()
        cur.close()
        conn.close()
        return sta

if __name__ == "__main__":
    b= 'ab'
    a = "'%s'"%b
    c = "woshi"+ b + '='+ "'%s'"%b +'asd'
    print(c)
    db = dbClass('10.10.20.108','qatest','qatest123qwe','jyallqa')
    a = '123'
    sql = "delete from m_assign_golden where user_id='%s'"%a
    sta = db.delete(sql)     
