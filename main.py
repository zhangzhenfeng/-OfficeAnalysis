# -*- coding: UTF-8 -*-
import xlrd,os,datetime,smtplib,traceback,binascii
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from log import Logger

# 实际工时的列位置
realHourCol = 4
# 预计工时的列位置
planHourCol = 3
# 完成度的列位置
doneCol = 5
# 任务描述列位置
taskDCol = 6
# 任务名称列位置
taskCol = 7
# 所属项目列位置
projectCol = 8
base_path = "./"
# 日报路径
ribao_path = "01-日报/"
# 代码目录
src_path = ""
# 发件人邮箱
sender = 'xxxx'
# 发件人邮箱密码，该密码为动态密码
sender_pass = 'xxx'
# 收件人邮箱
receivers_1 = {'xxx@qq.com':"xxx"}
nameToPin = {"张振峰":"zhangzhenfeng"}
receivers_2 = {"张振峰":'xxx@qq.com'}
# 腾讯服务器地址
smtp_server = 'smtp.exmail.qq.com'
logfile = "log.log"

def getDate(offset = 0):
    """
    获取日期
    :param offset: 当前日期的偏移量
    :return:
    """
    from datetime import datetime, date, timedelta
    ISOTIMEFORMAT='%Y-%m-%d'
    cur_time = (date.today() + timedelta(days = offset)).strftime(ISOTIMEFORMAT)
    return cur_time

def getTime(offset = 0):
    """
    获取时间
    :param offset: 当前日期的偏移量
    :return:
    """
    from datetime import datetime, date, timedelta
    ISOTIMEFORMAT='%Y-%m-%d %H:%M:%S'
    cur_time = (date.today() + timedelta(days = offset)).strftime(ISOTIMEFORMAT)
    return cur_time

def listxls(path):
    """
    枚举所有日报
    :param path: 日报路径
    :return: 所有日报路径的list
    """
    tmp = []
    lst = os.listdir(path)
    for l in lst:
        if 'xls' in l and '~' not in l:
            tmp.append(os.path.join(path,l))
    return tmp

def checkRealHours(path,sheet,rs,re):
    """
    检查实际工时
    :param path: 日报路径
    :param sheet: xls对象
    :param rs: 起始行
    :param re: 结束行
    :return: 工时数量
    """
    count = 0
    # 读取该天其实行到结束行的内容
    for index in range(rs,re):
        type = sheet.cell(index,realHourCol).ctype
        try:
            rh = float(sheet.cell(index,realHourCol).value)
            count += rh
        except:
            Logger().logger.error("[-]读取日报实际工时异常:文件路径[%s],rs值[%s],re值[%s],单元格类型[%s]" % (str(path),str(rs),str(re),str(type)))
            count = count if count != 0 else 0
    Logger().logger.info("[+][%s]在[%s]的实际工时为【%s】" % (outName(path),str(getDate(-1)),str(count)))
    # 返回该天总时间
    return count

def checkPlanHours(path,sheet,rs,re):
    """
    检查预计工时
    :param path: 日报路径
    :param sheet: xls对象
    :param rs: 起始行
    :param re: 结束行
    :return: 预计总工时
    """
    count = 0
    # 读取该天其实行到结束行的内容
    for index in range(rs,re):
        type = sheet.cell(index,planHourCol).ctype
        try:
            rh = float(sheet.cell(index,planHourCol).value)
            count += rh
        except:
            Logger().logger.error("[-]读取日报预计工时异常:文件路径[%s],rs值[%s],re值[%s],单元格类型[%s]" % (str(path),str(rs),str(re),str(type)))
            count = count if count != 0 else 0
    Logger().logger.info("[+][%s]在[%s]的预计工时为【%s】" % (outName(path),str(getDate()),str(count)))
    # 返回该天总时间
    return count

def readbase(sheet,rs,cs):
    """
    读取单元格数据
    :param sheet:xls对象
    :param rs:起始行
    :param cs:起始列
    :return:单元格内容
    """
    # 读取单元格内容
    value = sheet.cell(rs,cs).value
    return value

def checkdate(path,sheet,rs,cs):
    """
    读取日报时间
    :param path: 日报路径
    :param sheet: xls对象
    :param rs: 起始行
    :param cs: 起始列
    :return: 日报时间
    """

    value = sheet.cell(rs,cs).value
    type =  sheet.cell(rs,cs).ctype
    try:
        __date__ = xlrd.xldate_as_tuple(value,0)
        year = __date__[0]
        month = __date__[1]
        day = __date__[2]
        date = datetime.date(year,month,day)
        Logger().logger.info("[+][%s]在[%s]的实际时间为【%s】" % (outName(path),str(getDate()),str(date)))
        return date
    except:
        Logger().logger.error("[-]读取日报时间异常:文件路径[%s],时间值[%s],单元格类型[%s],异常内容:%s" % (str(path),str(value),str(type),str(traceback.format_exc())))
        return -1

def checkDone(sheet,rs,re,type="t",date=""):
    """
    检查完成度的内容
    :param sheet: xls
    :param rs: 起始行
    :param re: 结束行
    :param type: 标记今天还是昨天
    :param date: 日期
    :return:
    """
    message = ""
    try:
        for row in range(rs,re):
            doneVal = str(readbase(sheet,row,doneCol))
            if type == "t":
                if doneVal != "" and doneVal != None:
                    message += "[%s](今天)的完成度不能有值，因为今天还没开始工作！\n" % date
                    break
            elif type == "y":
                if doneVal == "" or doneVal == None:
                    message += "[%s](昨天)的完成度为空，请填写后提交！\n" % date
                    break
        return message
    except:
        Logger().logger.error("[-]检查完成度的内容异常:rs[%s],re[%s],type[%s],date[%s],异常内容:%s" % (str(rs),str(re),str(type),str(date),str(traceback.format_exc())))

def checkTask(sheet,rs,re,date=getDate()):
    """
    检查任务描述是否有空值
    :param sheet: xls对象
    :param rs: 开始行
    :param re: 结束行
    :param date: 日期
    :return: 错误信息
    """
    message = ""
    try:
        for row in range(rs,re):
            taskVal = readbase(sheet,row,taskCol)
            if taskVal == "" or taskVal == None:
                message += "[%s](今天)的任务名称为空，请填写后提交！\n" % date
                break
        return message
    except:
        Logger().logger.error("[-]检查任务名称是否有空值异常:rs[%s],re[%s],date[%s],异常内容:%s" % (str(rs),str(re),str(date),str(traceback.format_exc())))
def checkProject(sheet,rs,re,date=getDate()):
    """
    检查项目是否为空过
    :param sheet: xls对象
    :param rs: 起始行
    :param re: 结束行
    :param date: 日期
    :return: 错误信息
    """
    message = ""
    try:
        for row in range(rs,re):
            projectVal = readbase(sheet,row,projectCol)
            if projectVal == "" or projectVal == None:
                message += "[%s](今天)的任务所属项目为空，请填写后提交！\n" % date
                break
        return message
    except:
        Logger().logger.error("[-]检查任务所属项目异常:rs[%s],re[%s],date[%s],异常内容:%s" % (str(rs),str(re),str(date),str(traceback.format_exc())))

def checkTaskD(sheet,rs,re,date=getDate()):
    """
    检查任务描述
    :param sheet:xls对象
    :param rs:起始行
    :param re:结束行
    :param date:日期
    :return:错误信息
    """
    message = ""
    try:
        for row in range(rs,re):
            taskDVal = readbase(sheet,row,taskDCol)
            if taskDVal == "" or taskDVal == None:
                message += "[%s](今天)的任务描述为空，请填写后提交！\n" % date
                break
        return message
    except:
        Logger().logger.error("[-]检查任务描述异常:rs[%s],re[%s],date[%s],异常内容:%s" % (str(rs),str(re),str(date),str(traceback.format_exc())))


def sendMail(target,msg):
    """
    发送邮件
    :param target: 收件人
    :param msg: 邮件内容
    :return:
    """
    to = ""
    for t in target:
        to += "" if receivers_1.get(t) is None else receivers_1.get(t) + ','
    message = MIMEText(msg, 'plain', 'utf-8')
    #message['From'] = Header("张振峰<zhangzhenfeng@anyuntec.com>", 'utf-8')
    message['From'] = formataddr(["张振峰", "zhangzhenfeng@anyuntec.com"])
    message['To'] =  Header(to[0:-1], 'utf-8')

    subject = '%s日报反馈' % str(getDate())
    message['Subject'] = Header(subject, 'utf-8')

    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(smtp_server, 25)    # 25 为 SMTP 端口号
        smtpObj.login(sender,sender_pass)
        smtpObj.sendmail(sender, target, message.as_string())
        Logger().logger.info("[+]已成功给%s发送邮件。邮件内容【%s】" % (str(target),msg))
    except smtplib.SMTPException,e:
        Logger().logger.error("[-]无法发送邮件，涉及收件人%s,邮件内容【%s】，异常信息【%s】" % (str(target),msg,str(traceback.format_exc())))
def outName(path):
    f = path.index("日报-") + 7
    t = path.index(".xlsx")
    return path[f:t]
def sub(name,i):
    """
    绩效扣分
    :param name: 姓名
    :param i: 扣分数量
    :return:
    """
    name = nameToPin.get(name)
    obj = None
    count_path = os.path.join(base_path,src_path,"count.txt")
    with open(count_path,'r') as f:
        obj = f.readline()
        if obj is not None and obj != "":
            obj = eval(obj)
            if obj.get(name) is not None and obj.get(name) != "":
                obj[name] = obj[name] - i
            else:
                Logger().logger.error("[-]count.txt文件内容没有[%s]" % (name))
        else:
            Logger().logger.error("[-]count.txt文件内容为空")
    with open(count_path,'w') as f:
        f.write(str(obj))

def todayandyesterday(sheet,cells):
    """
    返回今天，昨天日报的单元格信息
    :param sheet:
    :param cells:
    :return:
    """
    __today__ = getDate()
    __yesterday__ = getDate(-1)
    today = None
    yesterday = None
    for c in cells:
        rs, re, cs, ce = c
        value = sheet.cell(rs,cs).value
        __date__ = xlrd.xldate_as_tuple(value,0)
        __ = str(datetime.date(__date__[0],__date__[1],__date__[2]))
        if today is None and __today__ == __:today = c
        if yesterday is None and __yesterday__ == __:yesterday = c
        if today is not None and yesterday is not None:break

    return today,yesterday

def readxls(paths):
    for path in paths:
        # 邮件内容
        message = ""
        name = outName(path)
        file=xlrd.open_workbook(path)
        sheet=file.sheet_by_index(0)
        cells = sheet.merged_cells
        # 获取今天和昨天日期对应的单元格
        today , yesterday = todayandyesterday(sheet,cells)
        date,ph = '',0
        if today != None and today != "":
            # 获取最新一天的单元格信息，起始行，结束行，起始列，结束列
            t_rs, t_re, t_cs, t_ce = today
            # 日报时间
            date = checkdate(path,sheet,t_rs, t_cs)
            # 预计工时
            ph = checkPlanHours(path,sheet,t_rs, t_re)
            # 检查今天的完成度是否有值
            message += checkDone(sheet,t_rs,t_re,"t",getDate())
            # 检查今天的任务描述是否有值
            message += checkTask(sheet,t_rs,t_re)
            # 检查今天的任务所属项目是否有值
            message += checkProject(sheet,t_rs,t_re)
            # 检查今天的任务描述是否有值
            message += checkTaskD(sheet,t_rs,t_re)
        # 判断日报时间是否和当前日期相同，如果不相同说明未提交日报
        if str(date) != str(getDate()):
            message += "请在早8:40前提交当日日报工作内容，绩效扣除0.01。\n"
            sub(name,1)
        if ph != 8 :
            message += "[%s]日工作预计工时总和非8个小时，请修改提交。\n" % (getDate())
        if yesterday != None and yesterday != "":
            # 昨天日报信息的单元格信息
            y_rs, y_re, y_cs, y_ce = yesterday
            # 检查实际工时
            rh = checkRealHours(path,sheet,y_rs, y_re)
            if rh < 8 and len(cells) > 1:
                message += "[%s]日工作实际工时小于8小时，请修改提交。\n" % (getDate(-1))
            # 检查昨天的完成度是否有值
            message += checkDone(sheet,y_rs,y_re,"y",getDate(-1))
        else:
            message += "[%s]日报缺失，请修改提交。\n" % (getDate(-1))
        Logger().logger.info("[+]预计邮件内容【%s】" % (message))
        if message is not None and message != "":
            message += "\n此邮件为自动发送，请勿回复。"
            sendMail([receivers_2.get(name)],message)


def updatesvn():
    os.system("svn update %s" % os.path.join(base_path,ribao_path))
if __name__ == '__main__':
    import sys
    reload(sys)
    sys.setdefaultencoding('utf-8')
    updatesvn()
    xlspath = listxls(os.path.join(base_path,ribao_path))
    readxls(xlspath)