import requests
from bs4 import BeautifulSoup
import xlwt
import time
import xlrd
import smtplib
from xlrd import open_workbook
from xlutils.copy import copy
# 定义程序开始执行时间与结束时间

STARTTime=8
ENDTime=23

head={
    'Host':'www.asxenergy.com.au',
    'Connection':'keep-alive',
    'Cache-Control':'max-age=0',
    'Accept':'text/html, */*; q=0.01',
    'X-Requested-With':'XMLHttpRequest',
    'User-Agent':'MoZilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.154 Safari/537.36 LBBROWSER',
    'Referer':'https://www.asxenergy.com.au/futures_au',
    'Accept-Encoding':'gZip, deflate, sdch',
    'Accept-Language':'Zh-CN,Zh;q=0.8',
    'Cookie':'_ga=GA1.3.686961823.1473602256; tdct=98101ed9685250c359c27bf9430a0e90; pdct=98101ed9685250c359c27bf9430a0e90'
}

def saveFlie(workbook,myDate):
    k=0
    try:
        workbook.save( myDate + '-Data' + '.xls')
        print('文件保存成功')
        #file.write('文件保存成功'+'\n')
        #file.flush()
        k=1
    except:
        print('文件保存失败，请关闭excel文件！')
        time.sleep(10)
        print('尝试重新保存')
        #file.write('尝试重新保存'+'\n')
        #file.flush()
    if(k==0):
        saveFlie(workbook, myDate)


def sendEmail(content):
    sender = 'rmhedge32@gmail.com'  # will be replaced with my real email address
    password = 'Wizards23'  # will be replaced with my real password
    receivers = ['rmhedge32@gmail.com']

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()
    server.login(sender, password)  # Exception here
    try:
        server.sendmail(sender, receivers, content)
        print("content:" + content)
        print("Message sent successfully")
    except:
        print("Failed to send message")
    server.quit()


def changeValue(x , y , Z):
    rb = open_workbook('data.xls')
    wb = copy(rb)
    ws = wb.get_sheet(0)
    ws.write(x, y, Z)
    wb.save('data.xls')


def openData():
    book = xlrd.open_workbook('data.xls')
    sheet1 = sheet1 = book.sheet_by_name('Sheet1')


# 打开文件，保存为日志
#file=open('log.txt','a')
# 打开excel文件，读取里面的内容
workbook  = xlwt.Workbook()



#四大洲的Z值
NZ = 0
VZ = 0
QZ = 0
SZ = 0









mycount=1
myDate=time.strftime("%Y%m%d%H%M%S", time.localtime(time.time()))
while(True):

    book = xlrd.open_workbook('data.xls')
    sheet1 = book.sheet_by_name('Sheet1')
    col = sheet1.col(0)

    # 四大洲的对比值
    NS = sheet1.cell(1, 1).value
    VS = sheet1.cell(1, 2).value
    QS = sheet1.cell(1, 3).value
    SS = sheet1.cell(1, 4).value

    print(str(NS) + " " + str(VS) + " " + str(QS) + " " + str(SS))

    k=0
    nowTime=time.strftime("%H", time.localtime(time.time()))
    if(int(nowTime)>=STARTTime and int(nowTime)<=ENDTime):
        myTime=time.strftime("%Y%m%d%H%M%S", time.localtime(time.time()))
        print('当前时间:' + myTime)
        #file.write('当前时间:' + myTime +'\n')
        #file.flush()
        print('开始读取数据')
        #file.write('开始读取数据'+'\n')
        #file.flush()
        try:
            r=requests.get("https://www.asxenergy.com.au/futures_au/dataset",headers=head)
            k=1
        except:
            print('网络连接失败，尝试重新连接！')
            #file.write('网络连接失败，尝试重新连接！'+'\n')
            #file.flush()
            time.sleep(5)
        if(k==1):
            numList=[]
            tableList=[]
            myList=r.text.split('\n')
            for i in range(len(myList)):
                tempList=[]
                if(myList[i].find('market-dataset-state')>-1):
                    soup=BeautifulSoup(myList[i])
                    souptext=soup.text.strip()
                    if(souptext!=""):
                        tableList.append(souptext)
                if(myList[i].find('instrument')>-1):
                    soup=BeautifulSoup(myList[i])
                    souptext=soup.text.strip()
                    if(souptext!=""):
                        tempList.append(souptext)
                    for j in range(i+1,len(myList)):

                        if(myList[j].find('instrument')>-1):
                            break
                        soup=BeautifulSoup(myList[j])
                        souptext=soup.text.strip()
                        if(souptext!=""):

                            tempList.append(souptext)
                    numList.append(tempList)
            print('读取数据成功')
            #file.write('读取数据成功'+'\n')
            #file.flush()
            mycount=mycount+1
            worksheet = workbook.add_sheet('Sheet_' +myTime)
            startRow=0
            startColumn=0
            countN=0
            startValue = 0
            # 把四个州的标题标记一下
            worksheet.write(0,0,tableList[0])
            worksheet.write(0,7,tableList[1])
            worksheet.write(0,14,tableList[2])
            worksheet.write(0,21,tableList[3])
            mailList=[]




            #把每个周的标记分布成4大列
            for i in range(len(numList)):
                if i % (len(numList)/4)==0:
                    startRow=1
                    startColumn=7*countN
                    countN=countN+1

                    print(startValue)

                    # 公式里的变量
                    A = numList[startValue + 7][1]
                    B = numList[startValue + 7][2]
                    C = numList[startValue + 8][1]
                    D = numList[startValue + 8][2]
                    E = numList[startValue + 9][1]
                    F = numList[startValue + 9][2]
                    G = numList[startValue + 10][1]
                    H = numList[startValue + 10][2]
                    I = numList[startValue + 7][3]
                    J = numList[startValue + 8][3]
                    K = numList[startValue + 9][3]
                    L = numList[startValue + 10][3]
                    M = numList[startValue + 7][6]
                    N = numList[startValue + 8][6]
                    O = numList[startValue + 9][6]
                    P = numList[startValue + 10][6]
                    Q = numList[startValue + 12][1]
                    R = numList[startValue + 12][2]
                    S = numList[startValue + 14][1]
                    T = numList[startValue + 14][2]
                    U = numList[startValue + 12][3]
                    V = numList[startValue + 14][3]
                    W = numList[startValue + 12][6]
                    X = numList[startValue + 14][6]

                    # AB Mean
                    if (A == "-" or B == "-"):
                        if (I == "-"):
                            MeanAB = M
                        else:
                            MeanAB = I
                    else:
                        MeanAB = (float(A) + float(B)) / 2

                    # CD Mean
                    if (C == "-" or D == "-"):
                        if (J == "-"):
                            MeanCD = N
                        else:
                            MeanCD = J
                    else:
                        MeanCD = (float(C) + float(D)) / 2

                    # EF Mean
                    if (E == "-" or F == "-"):
                        if (K == "-"):
                            MeanEF = O
                        else:
                            MeanEF = K
                    else:
                        MeanEF = (float(E) + float(F)) / 2

                    # GH Mean
                    if (G == "-" or H == "-"):
                        if (L == "-"):
                            MeanGH = P
                        else:
                            MeanGH = L
                    else:
                        MeanGH = (float(G) + float(H)) / 2

                    # QR Mean
                    if (Q == "-" or R == "-"):
                        if (U == "-"):
                            MeanQR = W
                        else:
                            MeanQR = U
                    else:
                        MeanQR = (float(Q) + float(R)) / 2

                    # ST Mean
                    if (S == "-" or T == "-"):
                        if (V == "-"):
                            MeanST = X
                        else:
                            MeanST = V
                    else:
                        MeanST = (float(S) + float(T)) / 2

                    print("A:" + A + " " +
                          "B:" + B + " " +
                          "C:" + C + " " +
                          "D:" + D + " " +
                          "E:" + E + " " +
                          "F:" + F + " " +
                          "G:" + G + " " +
                          "H:" + H + " " +
                          "I:" + I + " " +
                          "J:" + J + " " +
                          "K:" + K + " " +
                          "L:" + L + " " +
                          "M:" + M + " " +
                          "N:" + N + " " +
                          "O:" + O + " " +
                          "P:" + P + " " +
                          "Q:" + Q + " " +
                          "R:" + R + " " +
                          "S:" + S + " " +
                          "T:" + T + " " +
                          "U:" + U + " " +
                          "V:" + V + " " +
                          "W:" + W + " " +
                          "X:" + X)

                    print("MeanAB:" + str(MeanAB) + " " +
                          "MeanCD:" + str(MeanCD) + " " +
                          "MeanEF:" + str(MeanEF) + " " +
                          "MeanGH:" + str(MeanGH) + " " +
                          "MeanQR:" + str(MeanQR) + " " +
                          "MeanST:" + str(MeanST) )



                    # 计算公式为
                    Z = (((float(MeanAB) * 2160 + float(MeanCD) * 2184 + float(MeanEF) * 2208 + float(
                        MeanGH) * 2208) / 8760 + float(MeanST)) / 2 + (float(MeanQR) + float(MeanST)) / 2) / 2

                    #New South Wales  Z Value
                    if (startValue == 0):
                        NZ = Z

                    #Victoria Z Value
                    if (startValue == 41):
                        VZ = Z

                    #Queensland Z Value
                    if (startValue == 82):
                        QZ = Z
                    #South Australia
                    if (startValue == 123):
                        SZ = Z



                    worksheet.write(42,startColumn, "Z=" + str(Z))
                    print(Z)

                    startValue = startValue + 41

                #把 NumList 里面的数据写到Excel 里面
                for j in range(len(numList[i])):
                    #行，列，值
                    worksheet.write(startRow,j+startColumn,numList[i][j])
                    findAddress='cells('+str(startRow)+','+str(j+startColumn)+ ')'
                    findValue=numList[i][j]
                    #print('cells('+str(startRow)+','+str(j+startColumn)+ ')value:'+numList[i][j])

                startRow=startRow+1

            saveFlie(workbook,myDate)

            #开始比较
            if (NZ != 0):
                if (NZ > NS*1.01):
                    subject = "New South Wales Increase"
                    text = "New South Wales Increase" + "\n" +\
                               " Z Value is" + "  " + str(NZ) + "\n" \
                               "Set Value is    " + str(NS) + "\n" \
                               "Increase  +" + str(((NZ-NS)/NS)*100) + "%"
                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1,1,NZ)

                if (NZ < NS*0.99):
                    subject = "New South Wales Decline"
                    text = "New South Wales Decline" + "\n" +\
                           " Z Value is " + str(NZ) + "\n" +\
                           "Set Value is " + str(NS) + "\n" +\
                           "Decline  -" + str(((NS-NZ)/NS)*100) + "%"

                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1, 1, NZ)


            if (VZ != 0):
                if ( VZ > VS*1.01):
                    subject = "Victoria Increase "
                    text = "Victoria Increase" + "\n" +\
                           "Z Value is " + str(VZ) + "\n" +\
                           "Set Value is " + str(VS) + "\n" +\
                           "Increase  +" + str(((VZ-VS)/VS)*100) + "%"
                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1, 2, VZ)
                if ( VZ < VS*0.99):
                    subject = "Victoria Decline Decline"
                    text = "Victoria Decline" + "\n" +\
                           "Z Value is " + str(VZ) + "\n" +\
                           "Set Value is " + str(VS) + "\n" +\
                           "Decline  -" + str(((VS-VZ)/VS)*100) + "%"
                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1, 2, VZ)

            if (QZ != 0):
                if ( QZ > QS*1.01):
                    subject = "Queensland Increase"
                    text = "Queensland Increase" + "\n" +\
                           "Z Value is" + str(QZ) + "\n" +\
                           "Set Value is" + str(QS) + "\n" +\
                           "Increase  +" + str(((QZ-QS)/QS)*100) + "%"
                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1,3,QZ)

                if (QZ < QS*0.99):
                    subject = "Queensland Decline"
                    text = "Queensland Decline" + "\n" +\
                              "Z Value is" + str(QZ) + "\n" +\
                              "Set Value is" + str(QS) + "\n" +\
                              "Decline -" + str(((QS-QZ)/QS)*100) + "%"
                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1, 3, QZ)



            if (SZ != 0):
                if ( SZ > SS*1.01):
                    subject = "South Australia Increase"
                    text =  "South Australia Increase" + "\n" +\
                            "Z Value is" + str(SZ) + "\n" +\
                            "Set Value is" + str(SS) + "\n" +\
                            "Increase  +" + str(((SZ-QS)/QS)*100) + "%"
                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1, 4, SZ)
                if (SZ < SS*0.99):
                    subject = "South Australia Decline"
                    text = "South Australia Decline" + "\n" +\
                           "Z Value is" + str(SZ) + "\n" +\
                           "Set Value is" + str(SS) + "\n" +\
                           "Decline  -" + str(((SS-SZ)/SS)*100) + "%"
                    message = 'Subject: %s\n\n%s' % (subject, text)
                    sendEmail(message)
                    changeValue(1, 4, SZ)

            # print(len(numList)/4)
            print('程序5分钟后抓取数据')
            #file.write('程序5分钟后抓取数据'+'\n')
            #file.flush()
            print(str(NZ) +" "+ str(VZ)+ " " + str(QZ)+ " " + str(SZ))
            time.sleep(5)
    else:
        myTime=time.strftime("%Y%m%d%H%M%S", time.localtime(time.time()))
        print('当前时间:' + myTime+'，未到执行时间')
        #file.write('当前时间:' + myTime+'，未到执行时间'+'\n')
        #file.flush()
        time.sleep(6)


