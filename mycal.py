# -*- coding = utf-8 -*-
# @Time : 2022/12/18 15:57
# @Author : CQU20205644
# @File : mycal.py
# @Software : PyCharm
#./classes.xlsx

from icalendar import Calendar, Event, Alarm
import xlrd
import datetime
#打开excel
# wb = xlrd.open_workbook('./121.xlsx')
# #按工作簿定位工作表
# sh = wb.sheet_by_name('Sheet0')
# totalRow=sh.nrows
# print(sh.nrows)#有效数据行数
# print(sh.ncols)#有效数据列数
# print(tmp)
# print(type(tmp[3]))
# print(len(tmp[3]))


def getWeek(msg):#e.g "1-7,9" --> [0, 1, 2, 3, 4, 5, 6, 8] 周的偏移量
    tmp = msg.split(',')
    res = []
    for i in tmp:
        if '-' in i:
            ttmp = i.split('-')
            for j in range(int(ttmp[0]) - 1, int(ttmp[1])):
                res.append((j))
        else:
            res.append(int(i) - 1)
    #print(res)
    return res

bastTime= datetime.datetime(2023, 2, 13, 0) #基准时间，开学2023年2月13日
startTime=["T083000","T092500","T103000","T112500","T133000","T142500","T152000","T162500","T172000","T190000","T195500","T205000","T214500","T000000"]#空间换时间
endTime = ["T091500","T101000","T111500","T121000","T141500","T151000","T160500","T171000","T180500","T194500","T204000","T213500","T223000","T223000"]
ZHWeek=["一","二","三","四","五","六","日"]
myUID=0




def getClassComponent(week,day,time,className,msg,loc):
    aft_days = bastTime + datetime.timedelta(weeks=week,days=day)
    myTime = aft_days.strftime('%Y%m%d')
    hour=time.split('-')
    begin=myTime+startTime[int(hour[0])-1]
    end=myTime+endTime[int(hour[1])-1]
    tmpClass = Event()
    global myUID
    tmpClass.add('UID', str(myUID))  # ID
    myUID=myUID+1
    tmpClass.add('DTSTART;VALUE=DATE',begin)
    tmpClass.add('DTEND;VALUE=DATE',end)
    tmpClass.add('SUMMARY', className)
    tmpClass.add('DESCRIPTION', msg)
    tmpClass.add('LOCATION', loc)
    # 设置闹钟,提前10min提醒
    alarm = Alarm()
    alarm.add('ACTION', 'AUDIO')
    #alarm.add('TRIGGER;VALUE=DATE-TIME', '19760401T005545Z')
    alarm.add('TRIGGER', datetime.timedelta(minutes=-int(10)))
    alarm.add('DESCRIPTION', loc)
    tmpClass.add_component(alarm)
    return tmpClass

# class_cal = Calendar() #建立你的课表日历
# class_cal.add('VERSION', '2.0')
# class_cal.add('X-WR-CALNAME','cqu2020春课表')
# class_cal.add('X-APPLE-CALENDAR-COLOR', '#5c7651')
# class_cal.add('TZID', 'China Standard Time')
#
# for i in range(2,totalRow):
#     rawMsg=sh.row_values(i)
#     name = rawMsg[0]
#     msg = "教学班号:"+rawMsg[1]+" |"+ rawMsg[4]
#     if len(rawMsg[3])==0:
#         loc="暂无位置信息 |"+ rawMsg[4]
#     else:
#         loc=rawMsg[3]+" |"+ rawMsg[4]
#     Msg2=rawMsg[2]
#     Msg2=Msg2.strip("节")
#     Msg2=Msg2.split('周')
#     day=ZHWeek.index(Msg2[1][2])
#     time=Msg2[1][3:]
#     tmpWeek=getWeek(Msg2[0])
#     for week in tmpWeek:
#         class_cal.add_component(getClassComponent(week,day,time,name,msg,loc))
# f = open('我的课表.ics', 'wb')
# f.write(class_cal.to_ical())
# f.close()


def final():
    try:
        wb = xlrd.open_workbook('./classes.xlsx')
        # #按工作簿定位工作表
        sh = wb.sheet_by_name('Sheet0')
        totalRow=sh.nrows
        sh = wb.sheet_by_name('Sheet0')
        totalRow = sh.nrows
        class_cal = Calendar()  # 建立你的课表日历
        class_cal.add('VERSION', '2.0')
        class_cal.add('X-WR-CALNAME', 'cqu2020春课表')
        class_cal.add('X-APPLE-CALENDAR-COLOR', '#5c7651')
        class_cal.add('TZID', 'China Standard Time')

        for i in range(2, totalRow):
            rawMsg = sh.row_values(i)
            name = rawMsg[0]
            msg = "教学班号:" + rawMsg[1] + " |" + rawMsg[4]
            if len(rawMsg[3]) == 0:
                loc = "暂无位置信息 |" + rawMsg[4]
            else:
                loc = rawMsg[3] + " |" + rawMsg[4]
            Msg2 = rawMsg[2]

            flagcou = Msg2.find("节")
            flagDay = Msg2.find("星期")
            Msg2 = Msg2.strip("节")
            Msg2 = Msg2.split('周')
            if flagDay != -1:
                day = ZHWeek.index(Msg2[1][2])
            else:
                day = 0
            if flagcou != -1:
                time = Msg2[1][3:]
            else:
                time = '0-0'
            # Msg2 = Msg2.strip("节")
            # Msg2 = Msg2.split('周')
            # day = ZHWeek.index(Msg2[1][2])
            # time = Msg2[1][3:]
            tmpWeek = getWeek(Msg2[0])
            for week in tmpWeek:
                class_cal.add_component(getClassComponent(week, day, time, name, msg, loc))
        f = open('我的课表.ics', 'wb')
        f.write(class_cal.to_ical())
        f.close()
    except:
        print("ERR!")

final()
