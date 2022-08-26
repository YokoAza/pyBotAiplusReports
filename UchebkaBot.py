from this import s
from telegram import *
from telegram.ext import * 
from requests import *
import pandas as pd
import requests as req
import json
import xlsxwriter 
import urllib.parse
import re
import openpyxl
from pprint import pprint
from datetime import datetime
import operator

updater = Updater(token="5272838015:AAE5P_kf_g3_xzSurPI4z1pvOCPt3mG9ULk")
dispatcher = updater.dispatcher
params = "take=7000&statuses="+urllib.parse.quote('Занимается,Заморозка,Регистрация', safe='~()*!\'');
authkeyAst ='OSYKFlLNdjeHVtUvvUfzmxoJwVNPEKUFN52WeBDWV7gc86eq9aiGybPG1vqx%2FOKqVF0nafbcDpkTA1rQpZDcGg%3D%3D'
authkeyAlm ='VdqvXSXu%2Fq1DWiLefLBUihGMn7MHlvSP59HIHoHH7%2BLEtHB5dtznB6sqyJIPjH5w'
authkeyShym = "NgeceftL80JZBQ4YOI9GAWMUpbOLgvOaHbOhjtzFcEu3y04LAct6gJ3%2BLEnn3w4yIsVt%2BkLeAJfCoq5%2FF5jnaA%3D%3D"

alm = "Алматы"
ast = "Астана"
shym = "Шымкент"
city = ""
otv = "Ответственные"
iin = "ИИН"
klass = "Классы"
school = "Школы"
branch = "Отделения"
office = "Филиалы"
blocks = "Блоки обучения"
time = "Время обучения"
foto = "Фото"
contacts = "Контакты"
amocrmid = "AmoCrm ID"
studentId = ""
res = 'Результаты Ученика'
def startCommand(update: Update, context: CallbackContext):       
    context.bot.send_message(chat_id=update.effective_chat.id, text="Здравствуйте! Это бот учебного отдела Айплюс. Выберите в меню команд, что вы бы хотели сделать.")


def cityCommand(update: Update, context: CallbackContext):       
    buttons = [[KeyboardButton(alm)], [KeyboardButton(ast)],[KeyboardButton(shym)] ]
    context.bot.send_message(chat_id=update.effective_chat.id, text="Вам нужен отчёт по какому городу?", reply_markup=ReplyKeyboardMarkup(buttons))

def resultCommand(update: Update, context: CallbackContext):       
    context.bot.send_message(chat_id=update.effective_chat.id, text="Здравствуйте! Введите пожалуйста id нужного вам ученика и город")
  
    

def messageHandler(update: Update, context: CallbackContext):
    buttons = [[KeyboardButton(res)], [KeyboardButton(otv)], [KeyboardButton(iin)],[KeyboardButton(klass)] 
    ,[KeyboardButton(school)],[KeyboardButton(branch)],[KeyboardButton(office)]
    ,[KeyboardButton(blocks)],[KeyboardButton(time)],[KeyboardButton(foto)],[KeyboardButton(contacts)],[KeyboardButton(amocrmid)]]
    global city
    global studentId
    global authkeyAlm
    global authkeyAst
    global authkeyShym

    if alm in update.message.text:
        city = alm
        context.bot.send_message(chat_id=update.effective_chat.id, text="Вам нужен какой отчёт?(Это займёт какое то время)", reply_markup=ReplyKeyboardMarkup(buttons))

    elif ast in update.message.text:
        city = ast
        context.bot.send_message(chat_id=update.effective_chat.id, text="Вам нужен какой отчёт?(Это займёт какое то время)", reply_markup=ReplyKeyboardMarkup(buttons))
    elif shym in update.message.text:
        city = shym
        context.bot.send_message(chat_id=update.effective_chat.id, text="Вам нужен какой отчёт?(Это займёт какое то время)", reply_markup=ReplyKeyboardMarkup(buttons))
    elif re.match("^[0-9]{1,10}$", update.message.text):
        studentId = update.message.text
        context.bot.send_message(chat_id=update.effective_chat.id, text="Введите период времени, за который вам нужны результаты (Пример - начальная дата,конечная дата: 2022-01-01,2022-02-02)")
    if re.match("^([0-9]+(-[0-9]+)+),[0-9]{4}-[0-9]{2}-[0-9]{2}$", update.message.text):
        stringDates = update.message.text
        splitStringDates = stringDates.split(',')
        startDate = splitStringDates[0]
        endDate = splitStringDates[1]
        if city == alm:
            authKey = authkeyAlm
            domain = "aiplus"
        elif city == ast:
            authKey = authkeyAst
            domain = "aiplus-astana"
        elif city == shym:
            authKey = authkeyShym
            domain = "aiplus-shymkent"
        response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?id='+studentId+'&authkey='+authKey+'')
        todos = response.json()
        dum = json.dumps(todos)
        dictjson = json.loads(dum)
        #pprint(dictjson)
        try:
            for j in dictjson['Students']:
                clientId = j['ClientId']
                clientId = str(clientId)
                itemNames = {'field': []}
                for y in j['ExtraFields']:
                    itemNames['field'].append(y["Name"])
                    if ("Name",'Отделение') in y.items():
                        if y["Name"] == 'Отделение':
                            groupBranch = y['Value']
                            print(groupBranch)
                    if ("Name",'КЛАСС') in y.items():
                        print("zawel")
                        if y["Name"] == 'КЛАСС':
                            studKlass = y['Value']
                if studKlass == '4' or studKlass == '5':
                    context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")

                    response1 = req.get('https://'+domain+'.t8s.ru//Api/V2/GetEdUnitTestResults?dateFrom='+startDate+'&dateTo='+endDate+'&studentClientId='+clientId+'&TestTypeCategoryName=Attendance list&take=7000&authkey='+authKey)
                    todos1 = response1.json()
                    dum1 = json.dumps(todos1)
                    dictjson1 = json.loads(dum1)
                    resultM = {'Дата': [], 'Дз': [], 'Срез': [], 'Тема': []}
                    resultE = {'Дата': [], 'Дз': [], 'Срез': [], 'Тема': []}
                    resultK = {'Дата': [], 'Дз': [], 'Срез': [], 'Тема': []}
                    resultR = {'Дата': [], 'Дз': [], 'Срез': [], 'Тема': []}
                    response2 = req.get('https://'+domain+'.t8s.ru//Api/V2/GetEdUnitStudents?StudentClientId='+clientId+'&authkey='+authKey+'&take=10000&queryDays=true&dateFrom='+startDate+'&dateTo='+endDate)
                    todos2 = response2.json()
                    dum2 = json.dumps(todos2)
                    dictjson2 = json.loads(dum2)
                    listId = []
                    for i in dictjson2['EdUnitStudents']:
                        if 'EndDate' in i:
                            continue
                        else:        
                            listId.append(i['EdUnitId'])
                    for y in listId:
                        response3 = req.get('https://'+domain+'.t8s.ru//Api/V2/GetEdUnitStudents?StudentClientId='+clientId+'&authkey='+authKey+'&take=10000&queryDays=true&dateFrom='+startDate+'&dateTo='+endDate+'&edUnitId='+str(y))
                        todos3 = response3.json()
                        dum3 = json.dumps(todos3)
                        dictjson3 = json.loads(dum3)
                        for k in dictjson3['EdUnitStudents']:
                            groupName = k["EdUnitName"].split('.')
                            groupSubj = groupName[0]
                            for j in k['Days']:
                                if j['Pass'] is True:
                                    response4 = req.get('https://'+domain+'.t8s.ru//Api/V2/GetEdUnitTestResults?dateFrom='+j['Date']+'&dateTo='+j['Date']+'&edUnitId='+str(y)+'&authkey='+authKey)
                                    todos4 = response4.json()
                                    dum4 = json.dumps(todos4)
                                    dictjson4 = json.loads(dum4)
                                    if groupSubj == 'M':
                                        date = j['Date']
                                        # splitDate = date.split('-')
                                        # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                                        resultM['Дата'].append(date)
                                        resultM['Дз'].append('Отсутствовал')
                                        resultM['Срез'].append('Отсутствовал')
                                        for i in dictjson4['EdUnitTestResults']:
                                            for x in i['Skills']:
                                                if x['SkillName'] == 'Темы':
                                                    tema = x['Score']
                                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                                    df1 = pd.read_excel(xlsx, 'Математика')
                                                    df = pd.DataFrame(df1, columns = ['Номер','ТемаРО','Предмет'])
                                                    for p in df.index:
                                                        if df['Номер'][p] == tema:
                                                            tempTema = df['ТемаРО'][p]
                                                    if tempTema == '':
                                                        resultM['Тема'].append('нет темы')
                                                    else:
                                                        resultM['Тема'].append(tempTema)
                                                    break
                                            break
                                    elif groupSubj == 'E':
                                        date = j['Date']
                                        # splitDate = date.split('-')
                                        # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                                        resultE['Дата'].append(date)
                                        resultE['Дз'].append('Отсутствовал')
                                        resultE['Срез'].append('Отсутствовал')
                                        for i in dictjson4['EdUnitTestResults']:
                                            for x in i['Skills']:
                                                if x['SkillName'] == 'Темы':
                                                    tema = x['Score']
                                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                                    df1 = pd.read_excel(xlsx, 'Математика')
                                                    df = pd.DataFrame(df1, columns = ['Номер','ТемаРО','Предмет'])
                                                    for p in df.index:
                                                        if df['Номер'][p] == tema:
                                                            tempTema = df['ТемаРО'][p]
                                                    if tempTema == '':
                                                        resultM['Тема'].append('нет темы')
                                                    else:
                                                        resultM['Тема'].append(tempTema)
                                                    break
                                            break                                            
                                    elif groupSubj == 'K':
                                        date = j['Date']
                                        # splitDate = date.split('-')
                                        # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                                        resultK['Дата'].append(date)
                                        resultK['Дз'].append('Отсутствовал')
                                        resultK['Срез'].append('Отсутствовал')
                                        for i in dictjson4['EdUnitTestResults']:
                                            for x in i['Skills']:
                                                if x['SkillName'] == 'Темы':
                                                    tema = x['Score']
                                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                                    if studKlass == '4':
                                                        tempTema = ''
                                                        df1 = pd.read_excel(xlsx, '4 класс')
                                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                                        for p in df.index:
                                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Казахский язык' and df['Отделение'][p] == groupBranch:
                                                                tempTema = df['Тема'][p]
                                                        if tempTema == '':
                                                            resultK['Тема'].append('нет темы')
                                                        else:
                                                            resultK['Тема'].append(tempTema)
                                                            

                                                    elif studKlass == '5':
                                                        tempTema = ''
                                                        df1 = pd.read_excel(xlsx, '5 класс')
                                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                                        for p in df.index:
                                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Казахский язык' and df['Отделение'][p] == groupBranch:
                                                                tempTema = df['Тема'][p]
                                                        if tempTema == '':
                                                            resultK['Тема'].append('нет темы')
                                                        else:
                                                            resultK['Тема'].append(tempTema)
                                                    break
                                            break
                                    elif groupSubj == 'R':
                                        date = j['Date']
                                        # splitDate = date.split('-')
                                        # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                                        resultR['Дата'].append(date)
                                        resultR['Дз'].append('Отсутствовал')
                                        resultR['Срез'].append('Отсутствовал') 
                                        for i in dictjson4['EdUnitTestResults']:
                                            for x in i['Skills']:
                                                if x['SkillName'] == 'Темы':
                                                    tema = x['Score']
                                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                                    if studKlass == '4':
                                                        tempTema = ''

                                                        df1 = pd.read_excel(xlsx, '4 класс')
                                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                                        for p in df.index:
                                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Русский язык' and df['Отделение'][p] == groupBranch:
                                                                tempTema = df['Тема'][p]
                                                        if tempTema == '':
                                                            resultR['Тема'].append('нет темы')
                                                        else:
                                                            resultR['Тема'].append(tempTema)
                                                    elif studKlass == '5':
                                                        tempTema = ''
                                                        df1 = pd.read_excel(xlsx, '5 класс')
                                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                                        for p in df.index:
                                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Русский язык' and df['Отделение'][p] == groupBranch:
                                                                tempTema = df['Тема'][p]
                                                        if tempTema == '':
                                                            resultR['Тема'].append('нет темы')
                                                        else:
                                                            resultR['Тема'].append(tempTema)
                                                    break
                                            break
                    for i in dictjson1['EdUnitTestResults']:
                        groupName = i["EdUnitName"].split('.')
                        groupKlass = groupName[1]
                        if groupBranch == 'KO':
                            groupBranch = 'КО'
                        elif groupBranch == 'RO':
                            groupBranch = 'РО'
                        if groupName[0] == 'M':
                            tempTema = ''
                            count = 0
                            date = i['Date']
                            # splitDate = date.split('-')
                            # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                            resultM['Дата'].append(date)

                            for k in i['Skills']:
                                if k['SkillName'] == 'Оценка учителя':
                                    if k['Score'] > 10:
                                        k['Score'] = 'не было'
                                    resultM['Дз'].append(k['Score'])
                                elif k['SkillName'] == "Срез":
                                    if k['Score'] == 11:
                                        k['Score'] = 'не писал ученик'
                                    elif k['Score'] == 12:
                                        k['Score'] = 'не писала группа'
                                    resultM['Срез'].append(k['Score'])
                                elif k['SkillName'] == "Темы":
                                    tema = k['Score']
                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                    df1 = pd.read_excel(xlsx, 'Математика')
                                    df = pd.DataFrame(df1, columns = ['Номер','ТемаРО','Предмет'])
                                    for p in df.index:
                                        if df['Номер'][p] == tema:
                                            tempTema = df['ТемаРО'][p]
                                    if tempTema == '':
                                        resultM['Тема'].append('нет темы')
                                    else:
                                        resultM['Тема'].append(tempTema)
                                elif k['SkillName'] == "Ранг":
                                    print('k')
                                
                                count+=1
                            if count !=4:
                                resultM['Тема'].append('нет темы')

                        elif groupName[0] == 'E':
                            tempTema = ''
                            count = 0
                            date = i['Date']
                            # splitDate = date.split('-')
                            # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                            resultE['Дата'].append(date)                        
                            for k in i['Skills']:
                                if k['SkillName'] == 'Оценка учителя':
                                    if k['Score'] > 10:
                                        k['Score'] = 'не было'
                                    resultE['Дз'].append(k['Score'])
                                elif k['SkillName'] == "Срез":
                                    if k['Score'] == 11:
                                        k['Score'] = 'не писал ученик'
                                    elif k['Score'] == 12:
                                        k['Score'] = 'не писала группа'
                                    resultE['Срез'].append(k['Score'])
                                elif k['SkillName'] == "Темы":
                                    tema = k['Score']
                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                    df1 = pd.read_excel(xlsx, 'Английский')
                                    df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет'])
                                    for p in df.index:
                                        if df['Номер'][p] == tema:
                                            
                                            tempTema = df['Тема'][p]
                                    if tempTema == '':
                                        resultE['Тема'].append('нет темы')
                                    else:
                                        resultE['Тема'].append(tempTema)
                                elif k['SkillName'] == "Ранг":
                                    print('k')
                                count+=1
                            if count !=4:
                                resultE['Тема'].append('нет темы')
                        elif groupName[0] == 'K':
                            count = 0
                            date = i['Date']
                            # splitDate = date.split('-')
                            # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                            resultK['Дата'].append(date) 
                            for k in i['Skills']:
                                if k['SkillName'] == 'Оценка учителя':
                                    if k['Score'] > 10:
                                        k['Score'] = 'не было'
                                    resultK['Дз'].append(k['Score'])
                                elif k['SkillName'] == "Срез":
                                    if k['Score'] == 11:
                                        k['Score'] = 'не писал ученик'
                                    elif k['Score'] == 12:
                                        k['Score'] = 'не писала группа'
                                    resultK['Срез'].append(k['Score'])
                                elif k['SkillName'] == "Темы":
                                    tema = k['Score']
                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                    if groupKlass == '4':
                                        tempTema = ''
                                        df1 = pd.read_excel(xlsx, '4 класс')
                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                        for p in df.index:
                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Казахский язык' and df['Отделение'][p] == groupBranch:
                                                tempTema = df['Тема'][p]
                                        if tempTema == '':
                                            resultK['Тема'].append('нет темы')
                                        else:
                                            resultK['Тема'].append(tempTema)
                                            

                                    elif groupKlass == '5':
                                        tempTema = ''
                                        df1 = pd.read_excel(xlsx, '5 класс')
                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                        for p in df.index:
                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Казахский язык' and df['Отделение'][p] == groupBranch:
                                                tempTema = df['Тема'][p]
                                        if tempTema == '':
                                            resultK['Тема'].append('нет темы')
                                        else:
                                            resultK['Тема'].append(tempTema)
                                            
                                elif k['SkillName'] == "Ранг":
                                    print('k')
                                count+=1
                            if count !=4:
                                resultK['Тема'].append('нет темы')
                        elif groupName[0] == 'R':
                            count = 0
                            date = i['Date']
                            # splitDate = date.split('-')
                            # resultDate = splitDate[2] + '.' + splitDate[1] + '.' + splitDate[0]
                            resultR['Дата'].append(date) 
                            for k in i['Skills']:
                                if k['SkillName'] == 'Оценка учителя':
                                    if k['Score'] > 10:
                                        k['Score'] = 'не было'
                                    resultR['Дз'].append(k['Score'])
                                elif k['SkillName'] == "Срез":
                                    if k['Score'] == 11:
                                        k['Score'] = 'не писал ученик'
                                    elif k['Score'] == 12:
                                        k['Score'] = 'не писала группа'
                                    resultR['Срез'].append(k['Score'])
                                elif k['SkillName'] == "Темы":
                                    tema = k['Score']
                                    xlsx = pd.ExcelFile('Attendance_4-6_кл.xlsx')
                                    if groupKlass == '4':
                                        tempTema = ''

                                        df1 = pd.read_excel(xlsx, '4 класс')
                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                        for p in df.index:
                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Русский язык' and df['Отделение'][p] == groupBranch:
                                                tempTema = df['Тема'][p]
                                        if tempTema == '':
                                            resultR['Тема'].append('нет темы')
                                        else:
                                            resultR['Тема'].append(tempTema)
                                    elif groupKlass == '5':
                                        tempTema = ''
                                        df1 = pd.read_excel(xlsx, '5 класс')
                                        df = pd.DataFrame(df1, columns = ['Номер','Тема','Предмет','Отделение'])
                                        for p in df.index:
                                            if df['Номер'][p] == tema and df['Предмет'][p] == 'Русский язык' and df['Отделение'][p] == groupBranch:
                                                tempTema = df['Тема'][p]
                                        if tempTema == '':
                                            resultR['Тема'].append('нет темы')
                                        else:
                                            resultR['Тема'].append(tempTema)
                                elif k['SkillName'] == "Ранг":
                                    print('k')
                                count+=1
                            if count !=4:
                                resultR['Тема'].append('нет темы')
                    
                    
                    
                    dfInfoM = pd.DataFrame(resultM)
                    dfInfoM[['Дата']] = dfInfoM[['Дата']].apply(pd.to_datetime)
                    dfInfoM = dfInfoM.sort_values(by=['Дата'])
                    dfInfoM['Дата'] = dfInfoM['Дата'].dt.strftime('%d.%m.%Y')
                    dfInfoM['Дата'] = dfInfoM['Дата'].astype(str)
                    dfInfoM = dfInfoM.transpose()
                    pprint(dfInfoM)
                    dfInfoE = pd.DataFrame(resultE)
                    dfInfoE[['Дата']] = dfInfoE[['Дата']].apply(pd.to_datetime)
                    dfInfoE = dfInfoE.sort_values(by=['Дата'])
                    dfInfoE['Дата'] = dfInfoE['Дата'].dt.strftime('%d.%m.%Y')
                    dfInfoE['Дата'] = dfInfoE['Дата'].astype(str)
                    dfInfoE = dfInfoE.transpose()
                    pprint(dfInfoE)
                    dfInfoK = pd.DataFrame(resultK)
                    dfInfoK[['Дата']] = dfInfoK[['Дата']].apply(pd.to_datetime)
                    dfInfoK = dfInfoK.sort_values(by=['Дата'])
                    dfInfoK['Дата'] = dfInfoK['Дата'].dt.strftime('%d.%m.%Y')
                    dfInfoK['Дата'] = dfInfoK['Дата'].astype(str)
                    dfInfoK = dfInfoK.transpose()
                    pprint(dfInfoK)
                    dfInfoR = pd.DataFrame(resultR)
                    dfInfoR[['Дата']] = dfInfoR[['Дата']].apply(pd.to_datetime)
                    dfInfoR = dfInfoR.sort_values(by=['Дата'])
                    dfInfoR['Дата'] = dfInfoR['Дата'].dt.strftime('%d.%m.%Y')
                    dfInfoR['Дата'] = dfInfoR['Дата'].astype(str)
                    dfInfoR = dfInfoR.transpose()
                    pprint(dfInfoR)

                    salary_sheets = {'Математика': dfInfoM, 'Английский': dfInfoE, 'Казахский': dfInfoK,'Русский': dfInfoR}
                    writer = pd.ExcelWriter('Результаты ученика.xlsx', engine='xlsxwriter')

                    for sheet_name in salary_sheets.keys():
                        salary_sheets[sheet_name].to_excel(writer, sheet_name=sheet_name)
                    workbook = writer.book
                    # Light red fill with dark red text.
                    format1 = workbook.add_format({'bg_color':   '#FFC7CE'})

                    # Light yellow fill with dark yellow text.
                    format2 = workbook.add_format({'bg_color':   '#FFEB9C'})

                    # Green fill with dark green text.
                    format3 = workbook.add_format({'bg_color':   '#C6EFCE'})
                    
                    format4 = workbook.add_format()
                    listOfSheets = ['Математика','Английский','Русский','Казахский']
                    for w in listOfSheets:
                        worksheet = writer.sheets[w]

                        worksheet.conditional_format('A1:XFD1048576', 
                                                {'type': 'blanks',
                                                'stop_if_true': True,
                                                'format': format4})

                        worksheet.conditional_format('B3:X1000', 
                                                {'type':     'cell',
                                                'criteria': 'between',
                                                'minimum':    8,
                                                'maximum':    10,
                                                'format':   format3})
                        worksheet.conditional_format('B3:X1000', 
                                                {'type':     'cell',
                                                'criteria': 'between',
                                                'minimum':    5,
                                                'maximum':    8,
                                                'format':   format2})
                        worksheet.conditional_format('B3:X1000', 
                                                {'type':     'cell',
                                                'criteria': 'between',
                                                'minimum':    0,
                                                'maximum':    5,
                                                'format':   format1})           
                                                
                    writer.save()
                    
                    chat_id = update.message.chat_id
                    document = open('Результаты ученика.xlsx', 'rb')
                    context.bot.send_document(chat_id, document)
                else:
                    context.bot.send_message(chat_id=update.effective_chat.id, text="Данный ученик не является учеником 4-5 классов. Попробуйте заново - /result")
        except Exception as e:
            print(e)
            context.bot.send_message(chat_id=update.effective_chat.id, text="При получении оценок произошла ошибка.Попробуйте заново - /result")

    if city == alm and update.message.text == res:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Введите пожалуйста id нужного вам ученика и город")
    elif city == ast and update.message.text == res:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Введите пожалуйста id нужного вам ученика и город")
    elif city == shym and update.message.text == res:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Введите пожалуйста id нужного вам ученика и город")
        
    if city == alm:
        authKey = authkeyAlm
        domain = "aiplus"
        
        try:

            if otv in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                try:
                    result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                    for i in dictjson["Students"]:
                        itemNames = {'field': []}
                        studentid = i["Id"]
                        StudName = i["FirstName"] +' '+ i["LastName"]
                        StudStatus = i["Status"]
                        #Ответственный(ментор)
                        try:
                            for h in i["Assignees"]:
                                try:
                                    StudMentor = h["FullName"]
                                except KeyError:
                                    continue
                        except:
                            result['id'].append(studentid)
                            result['name'].append(StudName)
                            result['status'].append(StudStatus)
                            break
                    dfInfo = pd.DataFrame(result)
                    dfInfo = dfInfo.transpose
                    dfInfo.to_excel("Отчёт по ответственным(Алматы).xlsx")
                    chat_id = update.message.chat_id
                    document = open('Отчёт по ответственным(Алматы).xlsx', 'rb')
                    context.bot.send_document(chat_id, document)
                except:
                    context.bot.send_message(chat_id = update.effective_chat.id,text ="Проблемных учеников не найдено")
            elif iin in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []} 
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    #ИИН
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])
                            #print(itemNames.values())
                            if ("Name",'ИИН') in j.items():
                                if j["Name"] == 'ИИН':
                                    if re.match("^[0-9]{11}|[0-9]{12}$", j["Value"]):
                                        continue
                                    else:
                                        result['id'].append(studentid)
                                        result['name'].append(StudName)
                                        result['status'].append(StudStatus)
                                        result['Mentor'].append(StudMentor)
                                        break
                    except KeyError:
                    
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)

                    if 'ИИН' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
        
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по ИИН(Алматы).xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по ИИН(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif klass in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:     
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'КЛАСС' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по классам("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по классам(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif amocrmid in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:     
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'id amoCRM' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по AmoCRM ID(Алматы).xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по AmoCRM ID(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif school in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:
                        
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'Школа обучения' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по школам("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по школам(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif branch in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:
                    
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'Отделение' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по отделению("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по отделению(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif office in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        if len(i["OfficesAndCompanies"]) > 1:
                            result['id'].append(studentid)
                            result['name'].append(StudName)
                            result['status'].append(StudStatus)
                            result['Mentor'].append(StudMentor)
                            break
                    except KeyError:
                    
                        break
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по филиалам("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по филиалам(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif blocks in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])
                            if ("Name",'КЛАСС') in j.items():
                                if j["Name"] == 'КЛАСС':
                                    if j["Value"] == '4' or j["Value"] == '5':
                                        try:
                                            for j in i["ExtraFields"]:
                                                itemNames['field'].append(j["Name"])      
                                        except KeyError:
                                            result['id'].append(studentid)
                                            result['name'].append(StudName)
                                            result['status'].append(StudStatus)
                                            result['Mentor'].append(StudMentor)
                                            break
                                        list_of_values = []
                                        for key, value in itemNames.items():
                                            list_of_values.append(value)
                                        if 'Блок Английский язык' not in list_of_values[0] or "Блок Математика" not in list_of_values[0]:
                                            result['id'].append(studentid)
                                            result['name'].append(StudName)
                                            result['status'].append(StudStatus)
                                            result['Mentor'].append(StudMentor)
                    except KeyError:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по блоку обучения(Алматы).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по блоку обучения(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif time in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:
                        
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'Время обучения' not in list_of_values[0] or 'Время прихода' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по времени обучения(Алматы).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по времени обучения(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif foto in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        if len(i["PhotoUrls"]) == 0:
                            result['id'].append(studentid)
                            result['name'].append(StudName)
                            result['status'].append(StudStatus)
                            result['Mentor'].append(StudMentor)
                            break
                    except KeyError:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break 
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по фото(Алматы).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по фото(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif contacts in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for h in i["Agents"]:
                            continue
                    except:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по контактам(Алматы).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по контактам(Алматы).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
        except Exception as e:
            print(e)
            context.bot.send_message(chat_id=update.effective_chat.id, text="При получении отчёта произошла ошибка.Попробуйте заново - /city")
    elif city == ast:
        authKey = authkeyAst
        domain = "aiplus-astana"
        
        try:

            if otv in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                try:
                    result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                    for i in dictjson["Students"]:
                        itemNames = {'field': []}
                        studentid = i["Id"]
                        StudName = i["FirstName"] +' '+ i["LastName"]
                        StudStatus = i["Status"]
                        #Ответственный(ментор)
                        try:
                            for h in i["Assignees"]:
                                try:
                                    StudMentor = h["FullName"]
                                except KeyError:
                                    continue
                        except:
                            result['id'].append(studentid)
                            result['name'].append(StudName)
                            result['status'].append(StudStatus)
                            break
                    dfInfo = pd.DataFrame(result)
                    dfInfo.to_excel("Отчёт по ответственным(Астана).xlsx")
                    chat_id = update.message.chat_id
                    document = open('Отчёт по ответственным(Астана).xlsx', 'rb')
                    context.bot.send_document(chat_id, document)
                except:
                    context.bot.send_message(chat_id = update.effective_chat.id,text ="Проблемных учеников не найдено")

            elif iin in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []} 
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    #ИИН
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])
                            #print(itemNames.values())
                            if ("Name",'ИИН') in j.items():
                                if j["Name"] == 'ИИН':
                                    if re.match("^[0-9]{11}|[0-9]{12}$", j["Value"]):
                                        continue
                                    else:
                                        result['id'].append(studentid)
                                        result['name'].append(StudName)
                                        result['status'].append(StudStatus)
                                        result['Mentor'].append(StudMentor)
                                        break
                    except KeyError:
                        
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)

                    if 'ИИН' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
        
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по ИИН(Астана).xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по ИИН(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif klass in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:
                        
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'КЛАСС' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по классам("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по классам(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif amocrmid in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:     
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'id amoCRM' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по AmoCRM ID(Астана).xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по AmoCRM ID(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif school in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:
                        
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'Школа обучения' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по школам("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по школам(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif branch in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:
                        
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'Отделение' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по отделению("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по отделению(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif office in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        if len(i["OfficesAndCompanies"]) > 1:
                            result['id'].append(studentid)
                            result['name'].append(StudName)
                            result['status'].append(StudStatus)
                            result['Mentor'].append(StudMentor)
                            break
                    except KeyError:
                        
                        break
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel("Отчёт по филиалам("+city+").xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по филиалам(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif blocks in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])
                            if ("Name",'КЛАСС') in j.items():
                                if j["Name"] == 'КЛАСС':
                                    if j["Value"] == '4' or j["Value"] == '5':
                                        try:
                                            for j in i["ExtraFields"]:
                                                itemNames['field'].append(j["Name"])      
                                        except KeyError:
                                            result['id'].append(studentid)
                                            result['name'].append(StudName)
                                            result['status'].append(StudStatus)
                                            result['Mentor'].append(StudMentor)
                                            break
                                        list_of_values = []
                                        for key, value in itemNames.items():
                                            list_of_values.append(value)
                                        if 'Блок Английский язык' not in list_of_values[0] or "Блок Математика" not in list_of_values[0]:
                                            result['id'].append(studentid)
                                            result['name'].append(StudName)
                                            result['status'].append(StudStatus)
                                            result['Mentor'].append(StudMentor)
                    except KeyError:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по блоку обучения(Астана).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по блоку обучения(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif time in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for j in i["ExtraFields"]:
                            itemNames['field'].append(j["Name"])      
                    except KeyError:
                        break
                    list_of_values = []
                    for key, value in itemNames.items():
                        list_of_values.append(value)
                    if 'Время обучения' not in list_of_values[0] or 'Время прихода' not in list_of_values[0]:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по времени обучения(Астана).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по времени обучения(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif foto in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        if len(i["PhotoUrls"]) == 0:
                            result['id'].append(studentid)
                            result['name'].append(StudName)
                            result['status'].append(StudStatus)
                            result['Mentor'].append(StudMentor)
                            break
                    except KeyError:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break 
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по фото(Астана).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по фото(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            elif contacts in update.message.text:
                context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
                response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
                todos = response.json()
                dum = json.dumps(todos)
                dictjson = json.loads(dum)
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        continue
                    try:
                        for h in i["Agents"]:
                            continue
                    except:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break
                dfInfo = pd.DataFrame(result)
                dfInfo.to_excel('Отчёт по контактам(Астана).xlsx')
                chat_id = update.message.chat_id
                document = open('Отчёт по контактам(Астана).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
        except Exception as e:
            print(e)
            context.bot.send_message(chat_id=update.effective_chat.id, text="При получении отчёта произошла ошибка.Попробуйте заново - /city")
    elif city == shym:
        authKey = authkeyShym
        domain = "aiplus-shymkent"
        
        if otv in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            try:
                result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
                for i in dictjson["Students"]:
                    itemNames = {'field': []}
                    studentid = i["Id"]
                    StudName = i["FirstName"] +' '+ i["LastName"]
                    StudStatus = i["Status"]
                    #Ответственный(ментор)
                    try:
                        for h in i["Assignees"]:
                            try:
                                StudMentor = h["FullName"]
                            except KeyError:
                                continue
                    except:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        break
                dfInfo = pd.DataFrame(result)
                dfInfo = dfInfo.transpose
                dfInfo.to_excel("Отчёт по ответственным(Шымкент).xlsx")
                chat_id = update.message.chat_id
                document = open('Отчёт по ответственным(Шымкент).xlsx', 'rb')
                context.bot.send_document(chat_id, document)
            except:
                context.bot.send_message(chat_id = update.effective_chat.id,text ="Проблемных учеников не найдено")
        elif iin in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []} 
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                #ИИН
                try:
                    for j in i["ExtraFields"]:
                        itemNames['field'].append(j["Name"])
                        #print(itemNames.values())
                        if ("Name",'ИИН') in j.items():
                            if j["Name"] == 'ИИН':
                                if re.match("^[0-9]{11}|[0-9]{12}$", j["Value"]):
                                    continue
                                else:
                                    result['id'].append(studentid)
                                    result['name'].append(StudName)
                                    result['status'].append(StudStatus)
                                    result['Mentor'].append(StudMentor)
                                    break
                except KeyError:
                    break
                list_of_values = []
                for key, value in itemNames.items():
                    list_of_values.append(value)

                if 'ИИН' not in list_of_values[0]:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel("Отчёт по ИИН(Шымкент).xlsx")
            chat_id = update.message.chat_id
            document = open('Отчёт по ИИН(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif klass in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    for j in i["ExtraFields"]:
                        itemNames['field'].append(j["Name"])      
                except KeyError:
                
                    break
                list_of_values = []
                for key, value in itemNames.items():
                    list_of_values.append(value)
                if 'КЛАСС' not in list_of_values[0]:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel("Отчёт по классам("+city+").xlsx")
            chat_id = update.message.chat_id
            document = open('Отчёт по классам(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif amocrmid in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    for j in i["ExtraFields"]:
                        itemNames['field'].append(j["Name"])      
                except KeyError:     
                    break
                list_of_values = []
                for key, value in itemNames.items():
                    list_of_values.append(value)
                if 'id amoCRM' not in list_of_values[0]:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel("Отчёт по AmoCRM ID(Шымкент).xlsx")
            chat_id = update.message.chat_id
            document = open('Отчёт по AmoCRM ID(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif school in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    for j in i["ExtraFields"]:
                        itemNames['field'].append(j["Name"])      
                except KeyError:
                    break
                list_of_values = []
                for key, value in itemNames.items():
                    list_of_values.append(value)
                if 'Школа обучения' not in list_of_values[0]:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel("Отчёт по школам("+city+").xlsx")
            chat_id = update.message.chat_id
            document = open('Отчёт по школам(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif branch in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    for j in i["ExtraFields"]:
                        itemNames['field'].append(j["Name"])      
                except KeyError:
                    break
                list_of_values = []
                for key, value in itemNames.items():
                    list_of_values.append(value)
                if 'Отделение' not in list_of_values[0]:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel("Отчёт по отделению("+city+").xlsx")
            chat_id = update.message.chat_id
            document = open('Отчёт по отделению(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif office in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    if len(i["OfficesAndCompanies"]) > 1:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break
                except KeyError:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
                    break
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel("Отчёт по филиалам("+city+").xlsx")
            chat_id = update.message.chat_id
            document = open('Отчёт по филиалам(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif blocks in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    for j in i["ExtraFields"]:
                        itemNames['field'].append(j["Name"])
                        if ("Name",'КЛАСС') in j.items():
                            if j["Name"] == 'КЛАСС':
                                if j["Value"] == '4' or j["Value"] == '5':
                                    try:
                                        for j in i["ExtraFields"]:
                                            itemNames['field'].append(j["Name"])      
                                    except KeyError:
                                        result['id'].append(studentid)
                                        result['name'].append(StudName)
                                        result['status'].append(StudStatus)
                                        result['Mentor'].append(StudMentor)
                                        break
                                    list_of_values = []
                                    for key, value in itemNames.items():
                                        list_of_values.append(value)
                                    if 'Блок Английский язык' not in list_of_values[0] or "Блок Математика" not in list_of_values[0]:
                                        result['id'].append(studentid)
                                        result['name'].append(StudName)
                                        result['status'].append(StudStatus)
                                        result['Mentor'].append(StudMentor)
                except KeyError:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
                    break
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel('Отчёт по блоку обучения(Шымкент).xlsx')
            chat_id = update.message.chat_id
            document = open('Отчёт по блоку обучения(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif time in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    for j in i["ExtraFields"]:
                        itemNames['field'].append(j["Name"])      
                except KeyError:
                    break
                list_of_values = []
                for key, value in itemNames.items():
                    list_of_values.append(value)
                if 'Время обучения' not in list_of_values[0] or 'Время прихода' not in list_of_values[0]:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel('Отчёт по времени обучения(Шымкент).xlsx')
            chat_id = update.message.chat_id
            document = open('Отчёт по времени обучения(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif foto in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    if len(i["PhotoUrls"]) == 0:
                        result['id'].append(studentid)
                        result['name'].append(StudName)
                        result['status'].append(StudStatus)
                        result['Mentor'].append(StudMentor)
                        break
                except KeyError:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
                    break 
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel('Отчёт по фото(Шымкент).xlsx')
            chat_id = update.message.chat_id
            document = open('Отчёт по фото(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        elif contacts in update.message.text:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Отчёт готовится")
            response = req.get('https://'+domain+'.t8s.ru//Api/V2/GetStudents?'+params+'&authkey='+authKey)
            todos = response.json()
            dum = json.dumps(todos)
            dictjson = json.loads(dum)
            result = {'id': [], 'name': [], 'status': [], 'Mentor': []}
            for i in dictjson["Students"]:
                itemNames = {'field': []}
                studentid = i["Id"]
                StudName = i["FirstName"] +' '+ i["LastName"]
                StudStatus = i["Status"]
                #Ответственный(ментор)
                try:
                    for h in i["Assignees"]:
                        try:
                            StudMentor = h["FullName"]
                        except KeyError:
                            continue
                except:
                    continue
                try:
                    for h in i["Agents"]:
                        continue
                except:
                    result['id'].append(studentid)
                    result['name'].append(StudName)
                    result['status'].append(StudStatus)
                    result['Mentor'].append(StudMentor)
                    break
            dfInfo = pd.DataFrame(result)
            dfInfo.to_excel('Отчёт по контактам(Шымкент).xlsx')
            chat_id = update.message.chat_id
            document = open('Отчёт по контактам(Шымкент).xlsx', 'rb')
            context.bot.send_document(chat_id, document)
        

def unknown(update: Update, context: CallbackContext):
	update.message.reply_text(
		"Вы набрали неправильную команду" % update.message.text)

dispatcher.add_handler(CommandHandler("start", startCommand))
dispatcher.add_handler(CommandHandler("city", cityCommand))
dispatcher.add_handler(CommandHandler("result", resultCommand))

dispatcher.add_handler(MessageHandler(Filters.text, messageHandler))
updater.dispatcher.add_handler(MessageHandler(Filters.text, unknown))


updater.start_polling()