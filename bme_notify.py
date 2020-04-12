#!/usr/local/bin/python
# -*- coding: utf-8 -*-
import sys,os
from os.path import dirname, join, abspath
sys.path.insert(0, abspath(join(dirname(__file__), '..')))

import time
import requests
import schedule
import socket
import asyncio
from lxml import html
from xlutils.copy import copy
import xlsxwriter 
import xlrd 
import xlwt
import json
from config import *

_session = None
CSI_URL = ""
LINE_TOKEN = {}
USERNAME = ""
PASSWORD = ""
SLEEP = 0
SITE_NAME = ""

def init():
    global CSI_URL
    global LINE_TOKEN
    global USERNAME
    global PASSWORD
    global SLEEP
    global SITE_NAME
    with open("config.json") as f:
        data = json.load(f)
        CSI_URL = data['CSI_URL']
        LINE_TOKEN = data['LINE_TOKEN']
        USERNAME = data['USERNAME']
        PASSWORD = data['PASSWORD']
        SLEEP = data['SLEEP']
        SITE_NAME = data['SITE_NAME']

def reply(msg,token):
    HEADER = {'content-type':'application/x-www-form-urlencoded','Authorization':'Bearer '+ token}
    try:
        res=requests.post(LINE_URL, headers=HEADER,data={'message':msg})
    except:
        return 400
    finally:
        return 200
    
def department(department):
    token = ""
    try:
        return LINE_TOKEN[str(department)]
    except:
        return LINE_TOKEN[str(SITE_NAME)]

def status_reply(str1,msg,eng):
    msg_reply = ""
    no = 0
    if str1 in ["รอดำเนินการ","Pending","Waiting"]:
        msg_reply = 'ทางทีม BME โดยคุณ %s ได้รับแจ้งส่งซ่อมเครื่อง : %s \nรหัส : %s \nทางทีมจะทำการแก้ไขปัญหาให้เร็วที่สุด ขอบคุณครับ'%(eng,str(msg['TypeCode']),str(msg['MeCode']))
        no = 1
    elif str1 in ["คืนเครื่องแล้ว","Return equipment back"]:
        msg_reply = 'เครื่อง %s รหัส %s \nได้ทำการดำเนินการแก้ไขเสร็จและคืนเครื่องเป็นที่เรียบร้อยแล้ว \nขอบคุณที่ใช้บริการ BME ครับ'%(msg['TypeCode'],msg['MeCode'])
        if CSI_URL != "":
            msg_reply += "\nฝากทำแบบประเมินให้ด้วยนะครับ ขอบคุณครับ\n Link : {}".format(CSI_URL)
        no = 7
    elif str1 in ["ส่งซ่อมภายนอก","Send out source"]:
        msg_reply = 'ขณะนี้เครื่อง %s รหัส %s \nอยู่ในขั้นตอนดำเนินการส่งซ่อมกับทางบริษัท ครับ'%(msg['TypeCode'],msg['MeCode'])
        no = 2
    elif str1 in ["รอวัสดุเบิก","Spare part required","Waiting for spare part"]:
        msg_reply = 'ขณะนี้เครื่อง %s รหัส %s \nกำลังอยู่ในขั้นตอนการจัดหาอะไหล่เพื่อใช้ในการซ่อม ครับ'%(msg['TypeCode'],msg['MeCode'])
        no = 3
    elif str1 in ["กำลังดำเนินการ","In process"]:
        msg_reply = 'ทางทีม BME  โดยคุณ %s ได้รับแจ้งส่งซ่อมเครื่อง : %s \nรหัส : %s \nทางทีมจะทำการแก้ไขปัญหาให้เร็วที่สุด ขอบคุณครับ'%(eng,msg['TypeCode'],msg['MeCode'])
        no = 4
    elif str1 in ["ซ่อมเสร็จ","Completed"]:
        msg_reply = 'เครื่อง %s รหัส %s \nได้ทำการดำเนินการแก้ไขเสร็จแล้ว \nขอบคุณที่ใช้บริการ BME ครับ'%(msg['TypeCode'],msg['MeCode'])
        no = 5
    elif str1 in ["ยกเลิก","Canceled","Cancle","Cancel"]:
        msg_reply = 'เครื่อง %s รหัส %s \nได้ทำการดำเนินการยกเลิกการแจ้งส่งซ่อมแล้วครับ \n'%(msg['TypeCode'],msg['MeCode'])
        no = 6
    return msg_reply,no

def send_task(ses):
    try:
        print('[{}]- Task send running.'.format(SITE_NAME))
        try:
            result = ses.get(URL_ONLINE_REPAIR,headers=dict(referer = URL_ONLINE_REPAIR))
        except:
            return False
            
        tree = html.fromstring(result.content)
    
        # Repair invoiced (ใบแจ้งซ่อม)
        repairInvoiced = tree.xpath("//table[4]//table//table[2]//tr/td[4]/text()[normalize-space()]")
        # Date time (วันที่ เวลา ที่แจ้งส่งซ่อม)
        dateTime = tree.xpath("//table[4]//table//table[2]//tr/td[7]/text()[normalize-space()]")
        # Department ( หน่วยงานแจ้งซ๋อม )
        _department = tree.xpath("//table[4]//table//table[2]//tr/td[9]/text()[normalize-space()]")
        # Detail ( รายละเอียด )
        detail = tree.xpath("//table[4]//table//table[2]//tr/td[10]/text()[normalize-space()]")

        if len(repairInvoiced) == 0:
            print('[{}]- Empty datas.'.format(SITE_NAME))
            return True
        else:
            packet = {}
            _flag = True
            for i in range(0,len(repairInvoiced)):
                try:
                    _url = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/REQ_02.php?req_no={}".format(str(repairInvoiced[i]))
                    _referer = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/BJOBA_01online.php"
                    res = ses.get(_url,headers=dict(referer = _referer))
                except:
                    return False
                meCode = ""
                typeCode = ""
                detail_ = ""
                try:
                    _tree = html.fromstring(res.content)
                    sender = _tree.xpath('//table[3]//form/table//tr/td/table[2]//tr[8]/td[3]/input/@value')[0]
                    meCode = str(detail[i].split('/')[0].split(':')[1].strip())
                    typeCode = str(detail[i].split('/')[1].strip())
                    if len(sender) < 1:
                        sender = ""
                    detail_ = str(detail[i].split('/')[2].split(':')[1].strip())
                except:
                    meCode = str(detail[i].split('/')[0].split(':')[1].strip())
                    typeCode = str(detail[i].split('/')[1].strip())
                    detail_ = str(detail[i].replace('\xa0','')).split(':')[2].strip()
                    

                packet = {
                    'Invoice': str(repairInvoiced[i]),
                    'Date': str(dateTime[i].replace('\xa0','')),
                    'Department': str(_department[i].replace('\xa0','')),
                    'MeCode' : meCode,
                    'TypeCode': typeCode,
                    'Detail': detail_,
                    'sender': sender,
                    'status': '',
                    'check': 0
                }
                isInsert = write_excel(packet)
                if isInsert == True:
                    msg_send = 'หน่วย : %s ได้แจ้งส่งซ่อมเครื่อง : %s รหัส : %s โดยคุณ%s ในวันที่ %s ด้วยอาการคือ : %s ขอบคุณที่ใช้บริการแจ้งส่งซ่อมผ่านระบบ NSmart ครับ'%(str(packet['Department']),str(packet['TypeCode']),str(packet['MeCode']),str(packet['sender']),str(packet['Date']),str(packet['Detail']))
                    token = department(str(packet['Department']))
                    status_send = reply(msg_send,token)
                    if status_send == 200:
                        print('[%s]- Send to LINE %s successed.'%(SITE_NAME,str(packet['Department'])))
                    else:
                        print('[{}]- Can not send to LINE Notify!.'.format(SITE_NAME))
                        _flag = False
                else:
                    print('[{}]- Can not inserted to Excel!.'.format(SITE_NAME))
                    _flag = False
                packet = {}
            return _flag
    except:
        print('[{}]- Exit send task.'.format(SITE_NAME))
        return False

def doReconnectSync():
    try:
        token = ""
        with requests.Session() as s:
            result = s.get(LOGIN_URL)
            payload = {
                    'user': USERNAME,
                    'pass': PASSWORD,
                    'Submit': "Submit"
            }
            result = s.post(LOGIN_URL,data=payload,headers=dict(referer=LOGIN_URL))
            print('[{}]- Connecting to internet...'.format(SITE_NAME))
            return s
    except:
        print('[{}]- Error can not restarted!.'.format(SITE_NAME))
        return None

async def doReplyTaskAsync():
    while True:
        try:
            await asyncio.sleep(SLEEP)
            ses = doReconnectSync()
            send_task(ses)
            reply_task(ses)
            #os.system('clear')
            print('[%s] - Waiting for %d sec.'%(SITE_NAME,int(SLEEP)))
        except:
            await asyncio.sleep(10)

def write_excel(data):
    i = 1
    ck = True
    while ck != False:
        try:
            rb = xlrd.open_workbook(LOC) 
            sheet = rb.sheet_by_index(0) 
            sheet.cell_value(0, 0)
            if int(sheet.row_values(i)[0]) is None:
                ck = False
                break
            else:
                if sheet.row_values(i)[1] == data['Invoice']:
                    ck = False
                    break
                i+=1
        except Exception as ex:
            wb = copy(rb)
            w_sheet = wb.get_sheet(0)
            w_sheet.write(i,0,i)
            w_sheet.write(i,1,data['Invoice'])
            w_sheet.write(i,2,data['Date'])
            w_sheet.write(i,3,data['Department'])
            w_sheet.write(i,4,data['MeCode'])
            w_sheet.write(i,5,data['TypeCode'])
            w_sheet.write(i,6,data['Detail'])
            w_sheet.write(i,7,data['sender'])
            w_sheet.write(i,8,data['status'])
            w_sheet.write(i,9,data['check'])
            wb.save(LOC)
            return True
            break

def reply_task(ses):
    i = 1
    ck = True
    while ck != False:
        try:
            rb = xlrd.open_workbook(LOC) 
            sheet = rb.sheet_by_index(0) 
            sheet.cell_value(0, 0)
            if int(sheet.row_values(i)[0]) is None:
                ck = False
                break
            else:
                if sheet.row_values(i)[9] != 7:
                    result = {
                        'TypeCode': sheet.row_values(i)[5],
                        'MeCode': sheet.row_values(i)[4]
                    }
                    URL_FIND_REPAIR_INVOICE = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/req_audit.php?s_jobtype=&s_req_no=%s&s_sap_code=&s_aname=&s_req_status=&s_urgentstat=&s_dept=&s_bcode=&s_floorno="%str(sheet.row_values(i)[1])
                    try:
                        result_data = ses.get(URL_FIND_REPAIR_INVOICE, headers = dict(referer = URL_FIND_REPAIR_INVOICE))
                    except Exception as ex:
                        return False
                        break
                    tree = html.fromstring(result_data.content)
                    repairInvoiced = tree.xpath("//table[4]//table//table[2]//tr[3]/td[1]/text()[normalize-space()]")
                    statusCall = tree.xpath("//table[4]//table//table[2]//tr[3]/td[10]/text()[normalize-space()]")
                    
                    if sheet.row_values(i)[8] == "":
                        try:
                            _url = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/REQ_01_audit.php?req_no={}".format(str(sheet.row_values(i)[1]))
                            _referer = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/req_audit.php"
                            res = ses.get(_url, headers = dict(referer = _referer))
                        except:
                            return False
                
                        eng = ""
                        if res.status_code == 200:
                            _tree = html.fromstring(res.content)
                            try:
                                eng = str(_tree.xpath('//table[2]//table//table[2]//tr[2]/td[2]/text()[normalize-space()]')[2]).split('\xa0')[0]
                            except Exception as ex:
                                eng = ''
                        reply_msg,no = status_reply(str(statusCall[0].strip()),result,eng)

                        token = department(str(sheet.row_values(i)[3]))
                        if token != '':
                            status_code = reply(reply_msg,token)
                        
                            if status_code == 200:
                                _flag = True
                                print('[%s]- Reply to LINE %s successed.'%(SITE_NAME,str(sheet.row_values(i)[3])))
                                wb = copy(rb)
                                w_sheet = wb.get_sheet(0)
                                w_sheet.write(i,8,str(statusCall[0].strip()))
                                w_sheet.write(i,9,int(no))
                                wb.save(LOC)
                                ck = False
                                break
                            else:
                                _flag = False
                                print('[{}]- Can not reply to LINE Notify!.'.format(SITE_NAME))
                        else:
                            _flag = False
                            print('[%s]- Department %s can not registed!.'%(SITE_NAME,str(sheet.row_values(i)[3])))

                    if len(statusCall) > 0 and str(statusCall[0].replace("\xa0",'').strip()) not in str(sheet.row_values(i)[8]):
              
                        try:
                            _url = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/REQ_01_audit.php?req_no={}".format(str(sheet.row_values(i)[1]))
                            _referer = "http://nsmart.nhealth-asia.com/MTDPDB01/jobs/req_audit.php"
                            res = ses.get(_url, headers = dict(referer = _referer))
                        except:
                            return False
                
                        eng = ""
                        if res.status_code == 200:
                            _tree = html.fromstring(res.content)
                            try:
                                eng = str(_tree.xpath('//table[2]//table//table[2]//tr[2]/td[2]/text()[normalize-space()]')[2]).split('\xa0')[0]
                            except Exception as ex:
                                eng = ''
                        reply_msg,no = status_reply(str(statusCall[0].strip()),result,eng)
                        token = department(str(sheet.row_values(i)[3]))
                        if token != '':
                            status_code = reply(reply_msg,token)
                            if status_code == 200:
                                _flag = True
                                print('[%s]- Reply to LINE %s successed.'%(SITE_NAME,str(sheet.row_values(i)[3])))
                                wb = copy(rb)
                                w_sheet = wb.get_sheet(0)
                                w_sheet.write(i,8,str(statusCall[0].strip()))
                                w_sheet.write(i,9,int(no))
                                wb.save(LOC)
                            else:
                                _flag = False
                                print('[{}]- Can not reply to LINE Notify!.'.format(SITE_NAME))
                        else:
                            _flag = False
                            print('[%s]- Department %s can not registed!.'%(SITE_NAME,str(sheet.row_values(i)[3])))
                    else:
                        print('[%s]- %d-Empty datas!.'%(SITE_NAME,int(sheet.row_values(i)[1])))
                        _flag = True
                i+=1
        except Exception as ex:
            break

def main():
    try:
        loop = asyncio.get_event_loop()
        cors = asyncio.wait([doReplyTaskAsync()])
        loop.run_until_complete(cors)
    except Exception as e:
        print('[Main]- Error can not started!.{}'.format(e))

if __name__ == "__main__":
    try:
        init()
        print('[{}]- Start Service..'.format(SITE_NAME))
        main()
    except KeyboardInterrupt:
        print('[{}]- CQI Service stop!..'.format(SITE_NAME))
        sys.exit()
