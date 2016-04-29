#-*-coding:utf-8-*-
from interface_frame import *
import xml.etree.ElementTree as ms
import interface_frame
import json
from httplib2 import Response
import dbclass

expectxmldir = os.getcwd() + '/testdir/except/'
realxmldir = os.getcwd() + '/testdir/real/'

def run(interface_name , suiteid ,checkmethod = 1,method = "GET"):
    '''运行用例'''
    #打印信息
    print '   '+interface_name+'   '+' test start'
    #读取指定的文件格式
    return_method = excelobj.read_data(suiteid,2,12).upper()  #JSON
    print 'return_method:',return_method

    #返回内容格式
    content_type = excelobj.read_data(suiteid,2,13) #json
    print 'content_type:',content_type

    is_teardown = excelobj.read_data(suiteid,2,14)
    #关键字，指定的判断对象
    check_key = excelobj.read_data(suiteid,4,13)   #goldenId
    print 'check_key:',check_key

    #预期个实际值得存放位置
    global expectxmldir , realxmldir
    expectdir = expectxmldir + interface_name.strip()
    realdir = realxmldir + interface_name.strip()
    if os.path.exists(expectdir) == 0: #创建目录
        os.makedirs(expectdir)
    if os.path.exists(realdir) == 0:
        os.makedirs(realdir)
    excelobj.del_testrecord(suiteid)   #designatedGolden表单第14列的写入NT
    casecount = excelobj.get_ncase(suiteid)    #获取用例数
    caseinfolist = get_caseinfo(excelobj, suiteid)   #获取地址、2和goldenid
    
    for caseid in range(0,casecount):
        #判断用例是否可执行状态
        if excelobj.read_data(suiteid,excelobj.casebegin+caseid,2) == 'N':
            #不可运行就在testresult列写入NT
            write_result(excelobj, suiteid, caseid, excelobj.resultCol, 'NT')
            #excelobj.write_data(suiteid , excelobj.casebegin+caseid ,2 ,'NT' , color = NT_COLOR)
            continue
        if method == 'GET':
            url = get_input(excelobj, suiteid, caseid, caseinfolist,method)  #获得每个id访问地址
            resp,xmlstring,times = HTTP_Get(url)
            print 'resp:',resp
            #在第4行15列写入有3位小数的时间差，也就是响应时间
            excelobj.write_data(suiteid,excelobj.casebegin+caseid,excelobj.resTimeCol,str(round(times,3)))
            #����ֶ�
            if return_method == 'XML':   #判断设置的文件格式
                #��֤��Ӧ�����У�������֤�ֶ�
                #result_code = get_xmlstring_dict(xmlstring)['staus']
                result_code = resp['status']  #Status 属性指定服务器返回的状态行的值。使用该属性修改服务器返回的状态行。
                #如果返回值和预期不一致写入实际返回值，并返回NG
                ret1 = checkResult(excelobj, suiteid, caseid, result_code, excelobj.retCol)
                expectpath = expectdir + '/' + str(caseid+1)+'.xml'
                if checkmethod =='save' :
                    save_xml_file(expectpath, xmlstring)    #保存返回的内容为xml，在expect目录下
                #save real xml result
                realpath = realdir+'/'+str(caseid+1)+'.xml'
                save_xml_file(realpath, xmlstring)    #保存返回的内容为xml，在real目录下
            elif return_method == "JSON":
                #josn contant
                data = json.loads(xmlstring)
                d_num = len(data)+ 1
                if xmlstring !='':
                    if len(data)!=0:
                        lens = len(data)
                        if type(data) == list:
                            for x in range(0,lens):
                                flag = 0
                                real_json = data[x][check_key]  #取goldenID值
                                global exp
                                exp = excelobj.read_data(suiteid, excelobj.casebegin + caseid, 12)
                                if real_json == exp:   #和预期值进行对比
                                    ret3 = checkResult(excelobj, suiteid, caseid, real_json,excelobj.exp_data_col)
                                    flag = 1
                                    break
                            if flag == 0:  #即返回值和预期不一致
                                ret3 = 'NG' 
                        else:
                            real_json = data[check_key]
                            ret3 = checkResult(excelobj, suiteid, caseid, real_json,excelobj.exp_data_col)
                    else:
                        ret3 = 'NG'
                        real_json = ''
                else:
                    ret3 = 'NG'
                    real_json = ''
                result_code = resp['status']
                ret1 = checkResult(excelobj, suiteid, caseid, result_code, excelobj.retCol)
                expectpath = expectdir + '/' + str(caseid+1)+'.json'
                if checkmethod =='save' :
                    save_xml_file(expectpath, xmlstring)  
                #save real json result
                realpath = realdir+'/'+str(caseid+1)+'.json'
                save_xml_file(realpath, xmlstring)
            elif return_method == 'bool'.upper():
                result_code = resp['status']  #状态
                ret1 = checkResult(excelobj, suiteid, caseid, result_code, excelobj.retCol)
                expectpath = expectdir + '/' + str(caseid+1)+'.json'
                if checkmethod =='save' :
                    save_xml_file(expectpath, xmlstring)  
                #save real json result
                realpath = realdir+'/'+str(caseid+1)+'.json'
                save_xml_file(realpath, xmlstring)
                expect_bool = excelobj.read_data(suiteid,excelobj.casebegin+caseid,excelobj.exp_data_col)
                ret2 = checkResult(excelobj, suiteid, caseid, xmlstring.decode('utf-8'),excelobj.exp_data_col)
                
                
        elif method == 'POST':
            url = caseinfolist[0]   #地址
            data = get_input(excelobj, suiteid, caseid, caseinfolist,method)
            resp, xmlstring,times = HTTP_Post(url,data,content_type)
            print 'resp:',resp
            #在第4行15列
            excelobj.write_data(suiteid,excelobj.casebegin+caseid,excelobj.resTimeCol,str(round(times,3)))
            if return_method == "JSON":
                #result_code = json.JSONDecoder().decode(xmlstring.decode('utf-8'))['code']
                result_code = resp["status"]
                ret1 = checkResult(excelobj, suiteid, caseid, result_code, excelobj.retCol)
                expectpath = expectdir + '/' + str(caseid+1)+'.json'
                if checkmethod =='save' :
                    save_xml_file(expectpath, xmlstring)  
                #save real xml result
                realpath = realdir+'/'+str(caseid+1)+'.json'
                save_xml_file(realpath, xmlstring)
                
            elif return_method == 'bool'.upper():
                result_code = resp['status']
                ret1 = checkResult(excelobj, suiteid, caseid, result_code, excelobj.retCol)
                expectpath = expectdir + '/' + str(caseid+1)+'.json'
                if checkmethod =='save' :
                    save_xml_file(expectpath, xmlstring)  
                #save real json result
                realpath = realdir+'/'+str(caseid+1)+'.json'
                save_xml_file(realpath, xmlstring)
                expect_bool = excelobj.read_data(suiteid,excelobj.casebegin+caseid,excelobj.exp_data_col)
                ret2 = checkResult(excelobj, suiteid, caseid, xmlstring.decode('utf-8'),excelobj.exp_data_col)
        
        
            
        try:
            if return_method == 'bool'.upper():    
                pass
            else:
                ret2 = checkXmlFile(excelobj, suiteid, caseid, realpath, expectpath)
        except Exception:
            print('no except data')
            
        write_result(excelobj, suiteid, caseid, excelobj.resultCol,ret1,ret2,ret3)
        #tear down datebase
        db_type = excelobj.read_data(suiteid,2,9)
        list1 = excelobj.read_data(suiteid,2,10)
        if db_type != '':
            exp = excelobj.read_data(suiteid, excelobj.casebegin + caseid, 12)
            db = dbclass.dbClass('10.10.20.108','qatest','qatest123qwe','jyallqa')
            if db_type == 'DELETE':
                sql = db_type + " FROM "+ list1 + " WHERE " + check_key + "=" + "'%s'"%exp + " ORDER BY createtime DESC LIMIT 1"
            if db_type == 'UPDATE':
                tear_date = excelobj.read_data(suiteid, excelobj.casebegin + caseid, 19)
                sql = db_type +' '+ list1 + " set " + check_key + "=" + "'%s'"%tear_date + " where "+ check_key + "=" + "'%s'"%exp + " ORDER BY operatetime DESC LIMIT 1"
            sta = db.delete(sql)
            
        
    print interface_name+'    ' + ' Test End!'
    
if __name__ == '__main__':
    try:              
        #Add testsuite begin  
        #excelobj = create_excel(os.getcwd()+'/TestDir/testinterface.xlsx')
        excelobj = create_excel(os.getcwd()+'/JGJ/JGJ.xlsx')
        run('designatedGolden','designatedGolden','save','GET')
        #run('test','test','save','GET')
        #run("getgoldenByuserIdAndCountyId", 'getgoldenByuserIdAndCountyId','save','GET') 
        #run('getgoldenBygoldenId','getgoldenBygoldenId','save','GET')
        #run('getgoldenBygoldenTel','getgoldenBygoldenTel','save','GET')
        #run('getareaMangerByGoldenId','getareaMangerByGoldenId','save','GET')
        #run('getmygolden','getmygolden','save','GET')
  
        statisticresult(excelobj)  
        try:
            excelobj.close()
        except:
            print("close error")
    except:
        excelobj.close()
        exit()
        
    