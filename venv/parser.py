import xml.etree.cElementTree as ET
import xlrd, xlwt



#########################################-_i_-Работаем с xml_-i-_#######################################################

# 1. Для модификации указываем файл 'flow.xml' который должен лежать в одной папке с данным скриптом
# 2. Вычитываем структуру xml файла в переменную root. В итоге root становится -->
# --> многомерным массивом(ширина массива зависит от конкретного xml файла)
tree = ET.parse('flow.xml')
root = tree.getroot()
i = 1
# объявляем переменные для замены в необходимых элементах массива root
code = ""
name = ""
ip = ""
port = ""
countCopys=0
premission = False

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
sheet = readWorkbook.sheet_by_index(0)
coontRows = sheet.nrows

print("countCopys " + str(countCopys) + " coontRows " + str(coontRows))
if countCopys == coontRows:
    print("Количество скопированных элементов NIFI равно количеству строк в xls источнике. Продолжаем работу.")
    premission = True
if countCopys > coontRows:
    print("Количество скопированных элементов NIFI больше количества строк в xls источнике")

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
    # Перебираем елементы червёртого уровня в массиве элементов root
                for sub3 in sub2:

                    # Ищем текст "Copy of " в имени элементов четвёртого уровня эелементов root
                    if "Copy of " in str(sub3.text):
                        # открываем файл источник на чтение
                        readWorkbook = xlrd.open_workbook('conf.xls', formatting_info=True)
                        # выбираем активный лист
                        sheet = readWorkbook.sheet_by_index(0)
                        # Определяе цвет строки
                        for numRow in range(sheet.nrows):
                            rowColor = str(readWorkbook.colour_map[sheet.cell(i, 1).xf_index])
                            # print("each color "+rowColor)

                            if rowColor == "(0, 204, 255)":
                                print("white line " + str(numRow) + " i " + str(i))
                                break
                            i = i + 1
                        # print(i)
                        # записываем в переменные значения из источника(БД)
                        code = (str(sheet.row_values(i)[1]).split("."))[0]
                        name = ("mm" + code)
                        ip = str(sheet.row_values(i)[2])
                        port = (str(sheet.row_values(i)[3]).split("."))[0]
                        print("i " + str(i) + " code " + code + " name " + name + " ip " + ip + " port " + port)
                        sub3.text=code # меняем имя элемена "Copy of ********" на значение из источника данных
                        print("New name MM: "+sub3.text)
                        mainNameTrig=True # Активируем тригер, что группа процессоров с именем содержащим "Copy of " найдена
                        i = i + 1
    # Если тригер true - перебираем элементы пятого уровня в массиве root
                    if mainNameTrig:
                        for sub4 in sub3:
                            # Ищем элемент с именем "JettyWebSocketClient"
                            if (str(sub4.tag)=="name") and ("JettyWebSocketClient" in str(sub4.text)):
                                sub4.text = ("JettyWebSocketClient" + code) # Перезаписываем элеент "JettyWebSocketClient" + код ММ
                                print("sub4 tag: " + str(sub4.tag) + ", value: " + str(sub4.text))
    # Перебираем элементы шестого уровня в массиве root
                            for sub5 in sub4:
                                # Ищем элемент 'Server_IP', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='Server_IP':
                                    ipTrig=True
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - меняем значение на значение из источника и сбрасивыем тригер
                                if ipTrig:
                                    sub5.text=ip
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    ipTrig=False
                                # Ищем элемент 'Server_PORT', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='Server_PORT':
                                    portTrig=True
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - меняем значение на значение из источника и сбрасивыем тригер
                                if portTrig:
                                    sub5.text=port
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    portTrig=False
                                # Ищем элемент 'putsql-sql-statement', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='putsql-sql-statement':
                                    sqlTrig=True
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - модифицируем значение и подставляем в строку sql запроса название ММ из источника и сбрасивыем тригер
                                if sqlTrig:
                                    sub5.text=("INSERT INTO " + name + " (tag, data, controller_time, server_time) VALUES ('${tag}', '${data}', '${controllerTime}', '${now():format('yyyy-MM-dd HH:mm:ss')}')'")
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    sqlTrig=False
                                # Ищем элемент 'HASHNAME_PROPERTY', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='HASHNAME_PROPERTY':
                                    hashNameTrig=True
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - меняем значение на значение из источника и сбрасивыем тригер
                                if hashNameTrig:
                                    sub5.text=name
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    hashNameTrig=False
                                # Ищем элемент 'websocket-uri', активируем тригер нахождения элемента, переходим к следующему элементу массива -->
                                if sub5.text=='websocket-uri':
                                    websocketTrig=True
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    continue
                                # --> в элементе к которому перешли - модифицируем значение подставляя код ММ из источника и сбрасивыем тригер
                                if websocketTrig:
                                    sub5.text=("wss://10.8.37.125/ws/receive/"+code+"/")
                                    print("sub5 tag: " + str(sub5.tag) + ", value: " + str(sub5.text))
                                    websocketTrig=False
    # Перечитываем структуру массива root в tree
    tree = ET.ElementTree(root)
    with open("updated.xml", "w"): # Открываем файл updated.xml на запись. p.s. этот файл должен находиться в одной папке со скриптом
        tree.write("updated.xml") # Записывае модифицированную структуру массива root в файл updated.xml

