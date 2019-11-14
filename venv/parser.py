import xml.etree.cElementTree as ET
import xlrd, xlwt
import os
import shutil





#########################################-_i_-Работаем с xml_-i-_#######################################################

# 1. Для модификации указываем файл 'flow.xml' который должен лежать в одной папке с данным скриптом
# 2. Вычитываем структуру xml файла в переменную root. В итоге root становится -->
# --> многомерным массивом(ширина массива зависит от конкретного xml файла)
tree = ET.parse('flow.xml')
root = tree.getroot()
i = 1
q = 1
rowColor=''
countRows=0
countCopys=0
premission = False
stepRename = 0
godRowCount = 0
checkRowColor=""
alarmGodRowCount = False
breakStep=False
# объявляем переменные для замены в необходимых элементах массива root
code = ""
name = ""
ip = ""
port = ""
sub3NumElem = 0
sub4NumElem = 0
sub5NumElem = 0
websocketSub3NumElem = 0
websocketSub4NumElem = 0
websocketSub5NumElem = 0

# Перебираем елементы верхнего уровня в массиве элементов root
for elem in root:
# Перебираем елементы второго уровня в массиве элементов root
    for subelem in elem:
# Перебираем елементы третьего уровня в массиве элементов root
        for sub2 in subelem:
            # Объявляем тригеры для поиска интересующих нас элементов и замены их значений на значение из источника(БД)
            mainNameTrig=False
# Перебираем елементы червёртого уровня в массиве элементов root
            for sub3 in sub2:
                # Ищем текст "Copy of " в имени элементов четвёртого уровня эелементов root
                if "Copy of " in str(sub3.text):
                    countCopys = countCopys + 1

# открываем файл источник на чтение
readWorkbook = xlrd.open_workbook('conf.xls', formatting_info=True)
# выбираем активный лист
sheet = readWorkbook.sheet_by_name("Sheet1")
# Определяем количество строк в xls файле
countRows = int(sheet.nrows)

for checkNum in range(sheet.nrows):
    try:
        checkRowColor = str(readWorkbook.colour_map[sheet.cell(checkNum, 1).xf_index])
        if (checkRowColor == "(0, 204, 255)") or (checkRowColor == "(192, 192, 192)"):
            godRowCount = godRowCount + 1
    except:
        print("Ошибка открытия файла. Для продолжения корректной работы парсера, необходимо пересохранить conf.xls")
        premission = False

# print("countCopys " + str(countCopys) + " countRows " + str(countRows))
if countCopys == godRowCount:
    print("Количество скопированных элементов NIFI равно количеству строк в xls источнике.")
    premission = True
if countCopys > godRowCount:
    print("Внимание!!! Количество скопированных элементов NIFI больше количества строк в xls источнике. Продолжаем работу.")
    alarmGodRowCount = True
    premission = True
if countCopys < godRowCount:
    print("Внимание!!! Количество скопированных элементов NIFI меньше количества строк в xls источнике. Продолжаем работу.")
    premission = True

print("godRowCount "+str(godRowCount))
if godRowCount < 1:
    print("Нет доступных строк в источнике данных(conf.xls) для обновления конфигурации NIFI")
    premission = False

if premission:
    # Перебираем елементы верхнего уровня в массиве элементов root
    for elem in root:
    # Перебираем елементы второго уровня в массиве элементов root
        for subelem in elem:
    # Перебираем елементы третьего уровня в массиве элементов root
            for sub2 in subelem:
                # Объявляем тригеры для поиска интересующих нас элементов и замены их значений на значение из источника(БД)
                mainNameTrig=False
                ipTrig=False
                portTrig = False
                sqlTrig = False
                hashNameTrig = False
                websocketTrig = False
                sub3NumElem = 0
    # Перебираем елементы червёртого уровня в массиве элементов root
                for sub3 in sub2:
                    if breakStep:
                        break
                    # Ищем текст "Copy of " в имени элементов четвёртого уровня эелементов root
                    if "Copy of " in str(sub3.text):
                        jettyId = ''

                        # открываем файл источник на чтение
                        readWorkbook = xlrd.open_workbook('conf.xls', formatting_info=True)
                        # выбираем активный лист
                        sheet = readWorkbook.sheet_by_index(0)
                        # Определяе цвет строки
                        for numRow in range(sheet.nrows):
                            try:
                                rowColor = str(readWorkbook.colour_map[sheet.cell(i, 1).xf_index])
                                # print("rowColor: " + str(rowColor))
                            except:
                                print("Ошибка получения элемента листа книги в функции определения строки. ")
                                print("Служебная информация:")
                                print("numRow " + str(numRow) + ", i " + str(i) + ", range(sheet.nrows) "
                                      + str(range(sheet.nrows)) + ", countRows " + str(countRows))
                            if (rowColor == "(0, 204, 255)") or (rowColor == "(192, 192, 192)"):

                                ###############################################################

                                bookWrite = xlwt.Workbook()
                                sheetWrite = bookWrite.add_sheet('Sheet1')
                                for y in range((i+1), countRows):
                                    cell = xlwt.easyxf('pattern: pattern solid;')
                                    cell.pattern.pattern_fore_colour = 1
                                    sheetWrite.write(y, 0, str(sheet.row_values(y)[0]), cell)
                                    sheetWrite.write(y, 1, (str(sheet.row_values(y)[1]).split("."))[0], cell)
                                    sheetWrite.write(y, 2, str(sheet.row_values(y)[2]), cell)
                                    sheetWrite.write(y, 3, (str(sheet.row_values(y)[3]).split("."))[0], cell)
                                    sheetWrite.write(y, 4, str(sheet.row_values(y)[4]), cell)
                                    sheetWrite.write(y, 5, str(sheet.row_values(y)[5]), cell)
                                for x in range(0, i+1):
                                    cell = xlwt.easyxf('pattern: pattern solid;')
                                    if alarmGodRowCount and (x == (countRows)):
                                        breakStep=True
                                        cell.pattern.pattern_fore_colour = 2
                                        sheetWrite.write(x-1, 6, "Элементов Copy of в конфигурации NIFI больше чем "
                                                               "доступных строк в источнике(xls файле)", cell)
                                        print('Записана ошибка в xls файл: "Элементов Copy of в конфигурации NIFI '
                                              'больше чем доступных строк в источнике(xls файле)"')
                                        bookWrite.save('modified.xls')
                                        break
                                    else:
                                        cell.pattern.pattern_fore_colour = 42
                                    try:
                                    # print(x)
                                        if x!=countRows:
                                            sheetWrite.write(x, 0, str(sheet.row_values(x)[0]), cell)
                                            sheetWrite.write(x, 1, (str(sheet.row_values(x)[1]).split("."))[0], cell)
                                            sheetWrite.write(x, 2, str(sheet.row_values(x)[2]), cell)
                                            sheetWrite.write(x, 3, (str(sheet.row_values(x)[3]).split("."))[0], cell)
                                            sheetWrite.write(x, 4, str(sheet.row_values(x)[4]), cell)
                                            sheetWrite.write(x, 5, str(sheet.row_values(x)[5]), cell)
                                    except:
                                        print("x "+str(x))
                                bookWrite.save('modified.xls')
                                ###############################################################
                                # print("white line " + str(numRow) + " i " + str(i))
                                break
                            i = i + 1 # !!!!!!!!!!!!!! Изменить на определение строки элемента без заливки!!!!!!!!!
                        try:
                            # записываем в переменные значения из источника(xls)
                            if i>0 and i!=countRows:
                                code = (str(sheet.row_values(i-1)[1]).split("."))[0]
                                name = ("mm" + code)
                                ip = str(sheet.row_values(i-1)[2])
                                port = (str(sheet.row_values(i-1)[3]).split("."))[0]
                                # print("i " + str(i) + " code " + code + " name " + name + " ip " + ip + " port " + port)
                                sub3.text=code # меняем имя элемена "Copy of ********" на значение из источника данных
                                # print("New name MM: "+sub3.text)
                                print("i "+str(i))
                            if i>0 and i==countRows:
                                code = (str(sheet.row_values(i-1)[1]).split("."))[0]
                                name = ("mm" + code)
                                ip = str(sheet.row_values(i-1)[2])
                                port = (str(sheet.row_values(i-1)[3]).split("."))[0]
                                # print("i " + str(i) + " code " + code + " name " + name + " ip " + ip + " port " + port)
                                sub3.text = code  # меняем имя элемена "Copy of ********" на значение из источника данных
                                # print("New name MM: "+sub3.text)
                                print("i == " + str(i))
                        except:
                            print("Ошибка открытия файла. Для продолжения корректной работы парсера,"
                                  " необходимо пересохранить conf.xls")

                        mainNameTrig=True # Активируем тригер, что группа процессоров с именем содержащим "Copy of " найдена
                        i = i + 1 # !!!!!!!!!!!!!! Изменить на определение строки элемента без заливки!!!!!!!!!
                        stepRename = stepRename + 1
    # Если тригер true - перебираем элементы пятого уровня в массиве root
                    if mainNameTrig:
                        sub4NumElem = 0
                        sub5NumElem = 0
                        for sub4 in sub3:
                            # Ищем элемент с именем "JettyWebSocketClient"
                            if (str(sub4.tag)=="name") and ("JettyWebSocketClient" in str(sub4.text)):
                                sub4.text = ("JettyWebSocketClient" + code) # Перезаписываем элеент "JettyWebSocketClient" + код ММ
                                # print("sub4 tag: " + str(sub4.tag) + ", value: " + str(sub4.text))
                                # print("sub3 val 11 "+ str(sub3[sub4NumElem-1].text)+' sub4NumElem '+str(sub4NumElem-1))
                                jettyId = str(sub3[sub4NumElem-1].text)
                                # sub3[websocketSub4NumElem][websocketSub5NumElem].text = jettyId
                                # print("websocketSub3NumElem " + str(websocketSub3NumElem))
                                # print("websocketSub4NumElem "+str(websocketSub4NumElem))
                                # print(" websocketSub5NumElem "+str(websocketSub5NumElem))
                                # print(" sub_4[websocketSub4NumElem] " + str(sub2[4]))
                                # print(" sub_4-19[websocketSub4NumElem] " + str(sub2[4][18]))
                                # print(" sub_4-19-5[websocketSub4NumElem] " + str(sub2[4][18][0].text))
                                # print(" sub_4-19-5[websocketSub4NumElem] " + str(sub2[4][18][1].text))
                                sub2[websocketSub3NumElem][websocketSub4NumElem-1][1].text = jettyId
                                # print(" sub_4-19-5[websocketSub4NumElem] " + str(sub2[4][18][2].text))
                                # print(" sub_4-19-5[websocketSub4NumElem] " + str(sub2[4][18][3].text))
                                # print(" sub_4-19-5[websocketSub4NumElem] " + str(sub2[4][18][4].text))
                            sub4NumElem = sub4NumElem + 1
    # Перебираем элементы шестого уровня в массиве root
                            for sub5 in sub4:
                                if (str(sub5.tag) == "name") and ("websocket-controller-service-id" in str(sub5.text)):
                                    websocketSub3NumElem = sub3NumElem
                                    websocketSub4NumElem = sub4NumElem
                                    websocketSub5NumElem = sub5NumElem
                                    # print("sub5!!!!!!!!!!!!!!!!!!!!")
                                # Ищем элемент 'Server_IP', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='Server_IP':
                                    ipTrig=True
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - меняем значение на значение из источника и сбрасивыем тригер
                                if ipTrig:
                                    sub5.text=ip
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    ipTrig=False
                                # Ищем элемент 'Server_PORT', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='Server_PORT':
                                    portTrig=True
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - меняем значение на значение из источника и сбрасивыем тригер
                                if portTrig:
                                    sub5.text=port
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    portTrig=False
                                # Ищем элемент 'putsql-sql-statement', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='putsql-sql-statement':
                                    sqlTrig=True
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - модифицируем значение и подставляем в строку sql запроса название ММ из источника и сбрасивыем тригер
                                if sqlTrig:
                                    sub5.text=("INSERT INTO " + name + " (tag, data, controller_time, server_time) VALUES ('${tag}', '${data}', '${controllerTime}', '${now():format('yyyy-MM-dd HH:mm:ss')}')'")
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    sqlTrig=False
                                # Ищем элемент 'HASHNAME_PROPERTY', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='HASHNAME_PROPERTY':
                                    hashNameTrig=True
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - меняем значение на значение из источника и сбрасивыем тригер
                                if hashNameTrig:
                                    sub5.text=name
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    hashNameTrig=False
                                # Ищем элемент 'websocket-uri', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='websocket-uri':
                                    websocketTrig=True
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - модифицируем значение подставляя код ММ из источника и сбрасивыем тригер
                                if websocketTrig:
                                    sub5.text=("wss://10.8.37.125/ws/receive/"+code+"/")
                                    # print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    websocketTrig=False
                                sub5NumElem = sub5NumElem + 1
                    sub3NumElem = sub3NumElem + 1
    # Перечитываем структуру массива root в tree
    tree = ET.ElementTree(root)
    with open("updated.xml", "w"): # Открываем файл updated.xml на запись. p.s. этот файл должен находиться в одной папке со скриптом
        tree.write("updated.xml") # Записывае модифицированную структуру массива root в файл updated.xml
    readWorkbook.release_resources()
    del readWorkbook
    os.remove('conf.xls')
    shutil.copyfile("modified.xls", "conf.xls")

