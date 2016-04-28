#-*-coding:utf-8-*-
import os,urllib,httplib2,profile,datetime,time,filecmp 
import win32com.client
import xml.etree.ElementTree as Excel
from urllib import urlencode
import json
from nt import times
#color???
OK_COLOR = 0XFFFFE0
NG_COLOR = 0xff
NT_COLOR = 0XC0C0C0

TESTIME =[1,14]
TESTRESULT = [2,14]

class create_excel:
    def __init__(self,sfile,dtitleindex = 3,dcasebegin = 4,dargbegin = 3 ,dargcount = 8):
        self.xlApp = win32com.client.Dispatch("Excel.Application")
        try:
            self.book = self.xlApp.Workbooks.Open(sfile)
        except:
            print ("打开失败")
            exit()
        self.file = sfile
        self.titleindex = dtitleindex
        self.casebegin = dcasebegin
        self.argbegin = dargbegin
        self.argcount = dargcount
        self.allresult = []
        
        self.retCol = self.argbegin+self.argcount  #11
        self.exp_data_col = self.retCol+1  #12
        self.Key_Col = self.retCol+2   #13
        self.resultCol = self.Key_Col+1    #14
        self.resTimeCol = self.resultCol + 1  #15
        self.timeCol = self.resTimeCol+1    #16
        
    def close(self):
        self.book.Save()
        self.book.Close()
        del self.xlApp
    #��ȡ���
    def read_data(self, isheet , iRow ,iCol):
        '''读取excel数据，三个必传参数：sheet名，行，列'''
        try:
            sht = self.book.Worksheets(isheet)
            sValue = str(sht.Cells(iRow,iCol).value.encode('utf8'))  #取单元格的值
            
        except:
            sValue = str(sht.Cells(iRow,iCol).value)
        if sValue[-2:]=='.0':
            sValue = sValue[0:-2]
        return sValue
        
    def write_data(self, isheet , iRow ,iCol ,sData , color = OK_COLOR):
        '''写数据，sheet名，行，列，数据'''
        try:
            sht = self.book.Worksheets(isheet)
            sht.Cells(iRow,iCol).value = sData #.decode("utf-8")
            sht.Cells(iRow,iCol).interior.Color = color
            self.book.Save()
        except:
            self.close()
            print('写失败')
            exit()
            
    def get_nrows(self,isheet):
        '''获取sheet行数'''
        try:
            sht = self.book.Worksheets(isheet)
            return sht.UsedRange.Rows.Count
        except:
            self.close()
            print('获取行失败')
            exit()
            
    def get_ncols(self,isheet):
        '''获取sheet列数'''
        try:
            sht = self.book.Worksheets(isheet)
            return sht.UsedRange.Columns.Count    
        except:
            self.close()
            print('列失败')
            exit()
    #��ȡ�������
    def get_ncase(self,isheet):
        '''获取sheet用例数'''
        try:
            sht = self.book.Worksheets(isheet)
            return self.get_nrows(isheet) - self.casebegin+1  #获取所有测试数据数量
        except:
            self.close()
            print('获取用例数失败')
            exit()
    
    def del_testrecord(self,suiteid):
        '''清除数据'''
        try:
            nrows = self.get_nrows(suiteid)+1  #获取所有行
            
            ncols = self.get_ncols(suiteid)
            resbegincol = self.argbegin +self.argcount   #11
            sht = self.book.Worksheets(suiteid)
            for row in range(self.casebegin,nrows):   #从第4行（即第一行数据）到最后的所有行
                for col in range(resbegincol,17):     #从第11列到第16列数据
                    str=self.read_data(suiteid, row, col)
                    #print 'str:',str,
                    startpos = str.find('[')
                    if startpos>0:
                        str = str[0:startpos].strip()
                        self.write_data(suiteid, row ,col ,str , OK_COLOR)
                    else:
                        sht.Cells(row,col).Interior.Color = OK_COLOR
    #��TestResult��
                #self.write_data(suiteid,row,self.Key_Col,' ',OK_COLOR)
                self.write_data(suiteid ,row , self.resultCol,'NT',NT_COLOR)  #self.resultCol=14
        except:
            self.close()
            print("清除失败")
            exit()
    #get����
def HTTP_Get(url):
    h = httplib2.Http() 
    start = time.time()
    response,content = h.request(url,'POST')
    print 'response:',response
    #content = json.JSONDecoder().decode(content.decode('utf-8'))
    end = time.time()
    times = end -start
        #data = content.read()
    return response,content,times   #返回响应内容和时间

def HTTP_Post(url,data,contentType = 'application/x-www-form-urlencoded'):
    h = httplib2.Http()
    start = time.time()
    header = {'Content-Type':'application/'+contentType,'Accept':'application/json'}
    jdata = json.dumps(data)
    response,content = h.request(url, 'POST',body = jdata,headers = header)
    end = time.time()
    times = end -start
    return (response,content,times)
    
def get_caseinfo(Data,suiteid):
    try:
        caseinfolist = []
        sInterface = Data.read_data(suiteid , 1 ,2)   #地址
        argcount = Data.read_data(suiteid , 2 , 2)   #2
        
        argnamelist = []
        for i in range(0,int(argcount)):
             #保存第3行，第4列数据=goldenid
            argnamelist.append(Data.read_data(suiteid,Data.titleindex,Data.argbegin+i))
        caseinfolist.append(sInterface)   #添加地址
        caseinfolist.append(int(argcount))   #2
        caseinfolist.append(argnamelist)    #goldenid
        return caseinfolist    #返回列表，包含地址、2和goldenid
    except:
        print('获取case失败')
        exit()
        
def get_input(Data, suiteid ,caseid , caseinfolist,method):
    sArg='?'
    data = {}
    if method == 'GET':
        for j in range(0,caseinfolist[1]):   #goldenid个数
            #读取每个id
            print 'Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j):',Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j)
            print 'Data.casebegin+caseid:',Data.casebegin+caseid
            print 'Data.argbegin+j :',Data.argbegin+j
            if Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j)!='None':
                #       ?         当前goldenid
                sArg = sArg + caseinfolist[2][j]+'='+Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j)+'&'
                print 'sArg:',sArg
        if sArg[-1] =='&':
            sArg = sArg[0:-1]
            sinput = caseinfolist[0]+sArg
        print 'sinput:',sinput
        return sinput
    elif method == 'POST':
        for j in range(0,caseinfolist[1]):
            if Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j)!='None':
                data[caseinfolist[2][j]] = Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j)
        return data
    
def get_input2(Data, suiteid ,caseid , caseinfolist,method):
    sArg = "/"
    for j in range(0,caseinfolist[1]):
        if Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j)!='None':
            sArg = sArg + Data.read_data(suiteid,Data.casebegin+caseid,Data.argbegin+j) +'/'
    if sArg[-1] =='/':
            sArg = sArg[0:-1]
            sinput = caseinfolist[0]+sArg
    return sinput        
    

def assert_result(sReal,sExcept):
    sReal = str(sReal)
    sExcept = str(sExcept)
    if sReal == sExcept:
        return 'OK'
    else:
        return 'NG'
            
def write_result(Data , suiteid , caseid , resultcol , *result):
    ret = 'OK'
    if len(result)>1:
        ret = 'OK'
        for i in range(0,len(result)):
            if result[i] == 'NG':
                ret = 'NG'
                break
            
        if ret == 'NG':    
            Data.write_data(suiteid , Data.casebegin+caseid , resultcol ,ret,NG_COLOR )
        else:
            Data.write_data(suiteid , Data.casebegin+caseid , resultcol ,ret,OK_COLOR )
            
        Data.allresult.append(ret)
    else:
        if result[0] =='NG':
            Data.write_data(suiteid , Data.casebegin+caseid , resultcol ,result[0],NG_COLOR )
        elif result[0] == 'OK':
            Data.write_data(suiteid , Data.casebegin+caseid , resultcol ,result[0],OK_COLOR )
        else :
            Data.write_data(suiteid , Data.casebegin+caseid , resultcol ,result[0],NT_COLOR )
        Data.allresult.append(result[0])
            
    print("case"+str(caseid+1)+": "+Data.allresult[-1]) 
            
def countflag(flag,ok, ng, nt):   
    calculation  = {'OK':lambda:[ok+1, ng, nt],    
                    'NG':lambda:[ok, ng+1, nt],                        
                    'NT':lambda:[ok, ng, nt+1]}       
    return calculation[flag]()   

def statisticresult(excelobj):  
    allresultlist=excelobj.allresult  
    count=[0, 0, 0]  
    for i in range(0, len(allresultlist)):  
        #print 'case'+str(i+1)+':', allresultlist[i]  
        count=countflag(allresultlist[i],count[0], count[1], count[2])  
    print ('Statistic result as follow:')  
    print ('OK:', count[0])  
    print ('NG:', count[1])  
    print ('NT:', count[2])  
'''          
def get_xmlstring_dict(xmlstring):
    xml = XML2Dict()
    r = xml.fromstring(xmlstring.decode('UTF-8'))
    return r

def get_xmlfile_dict(xmlfile):
    xml = XML2Dict() 
    return xml.prase(xmlfile) 
'''
def save_xml_file(path,xmlstring):
    f = open(path , 'wb')
    f.write(xmlstring)
    f.close()
    
def checkResult(excelobj, sheetid, caseid, real, begincol):
    ret = 'OK'
    exp = excelobj.read_data(sheetid, excelobj.casebegin + caseid, begincol)  #读取result_code值

    #如果检查不一致测将实际结果写入expect字段，格式：exp[real]
    #将return NG
    result = assert_result(real, exp)   #对比系统返回的值和Excel中的预期值
    if result == 'NG':
        writestr = exp + '[' + str(real) + ']'
        excelobj.write_data(sheetid, excelobj.casebegin + caseid, begincol, writestr, NG_COLOR)
        ret = 'NG'
    return ret
            
def checkXmlFile(excelobj, sheetid, caseid, file1 , file2 ):
    ret = 'OK'
    if not file2:
        return ret
    if not filecmp.cmp(file1,file2):
        ret = 'NG'
        '''
        excelobj.write_data(sheetid, excelobj.casebegin + caseid, excelobj.Key_Col, 'Error', NG_COLOR)
    else:
        excelobj.write_data(sheetid, excelobj.casebegin + caseid, excelobj.Key_Col, 'OK', OK_COLOR)
        '''
    return ret
            
#strip except
def delcomment(excelobj, suiteid, iRow, iCol, str):  
    startpos = str.find('[')  
    if startpos>0:  
        str = str[0:startpos].strip()  
        excelobj.write_data(suiteid, iRow, iCol, str, OK_COLOR)  
    return str  
   
def Time():
    tim=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
    return tim

 