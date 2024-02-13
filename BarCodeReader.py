# -----Библиотеки-----
from barcode import EAN13 #Есть и другие варианты: EAN-8, EAN-13, EAN-14, UPC-A, JAN, ISBN-10, ISBN-13
from barcode.writer import ImageWriter #Для генерации Штрих-кода
from pyzbar.pyzbar import decode #Для распознавания кода на изображении
import cv2 as cv #Библиотека OpenCV
import serial #Библиотка для работы с последовательным портом
import time #Библиотка для работы со временем
import colorama as cl #Для красивого текста консоли
import art #Тоже для красивого текста
import openpyxl #Для работы с excel-файлами

#----Постоянные----
ArduinoControl = True #Вывод сигналов на последовательный порт
ArduinoPort = 'COM5' #Порт с подключенной дуиной
ArduinoSpeed = 19200 #Скорость передачи данных в бодах

CheckCode = "Codes/IvanVas1.png" #Код для проверки без камеры 1
#CheckCode = "NewCode.png" #Код для проверки без камеры 2
CamWork = False # Включение поиска кода через камеру

excel_filename = 'Данные.xlsx' #Файл с данными для чтения
column_index = 4  # Номер колонки со штрих-кодами

tickTimer = 0.5 # Таймер сна, для более медленного выполнения кода, значения в секундах

#---Переменные---
acsees_list = [] #Список допущенных

#Генерация штрих-кода
def BarCodeGenerate(number = "0000000000000",name = "Default"):
    my_code = EAN13(number, writer=ImageWriter())
    my_code.save(name)
    print(cl.Fore.BLACK + f"Для номера {number} сохранён код под названием: {name}")
    acsees_list.append(number)
    print(cl.Fore.BLACK + f"Список допущенных: {acsees_list}")

#Чтение штрих-кода
def BarcodeReader():
    acsees_list = read_column_from_excel(excel_filename, column_index) #Обновим список допущенных

    #Если включено распознавание при помощи камеры
    if CamWork:
        cap = cv.VideoCapture(0)  # Подключаемся (захватываем) к камере. 0 — это индекс камеры, если их несколько то будет 0 или 1 и т.д.
        ret, img = cap.read()  # Читаем с устройства кадр, метод возвращает флаг ret (True , False) и img — саму картинку (массив numpy)
        #cv.imshow("Camera view", img)
        cv.imwrite("Screen.jpg", img) #Сохраняем изображение
        barcode = cv.imread("Screen.jpg") #Читаем изображение
        detectedBarcodes = decode(barcode)  # Распознаём штрих-код

    #Иначе читаем заготовленный заранее
    elif CamWork == False:
        barcode = cv.imread(CheckCode)  # Читаем изображение
        detectedBarcodes = decode(barcode)  # Распознаём штрих-код

    # Если код не распознан
    if not detectedBarcodes:
        print(cl.Fore.RED + "Штрих-код не распознан, он либо пустой, либо повреждён!")
        if ArduinoControl:
            SendData(0) #Отправим на Ардуино сигнал об этом
    else:
        # Просмотр всех обнаруженных штрих-кодов на изображении
        for barcode in detectedBarcodes:
            # Определяем положение штрих-кода на изображении
            (x, y, w, h) = barcode.rect

            # Пометим прямоугольником распознанную область
            if CamWork:
                cv.rectangle(img, (x - 10, y - 10),
                         (x + w + 10, y + h + 10),
                         (255, 0, 0), 2)

            # Если данные не пустые
            if barcode.data != "":
                # Выведем код и его тип
                data = int(barcode.data)
                print(cl.Fore.BLACK + f"Данные на штрих-коде: {data}")
                print(cl.Fore.BLACK + f"Тип штрих-кода: {barcode.type}")

                if ArduinoControl:
                    SendData(1)  # Отправим на Ардуино сигнал об этом

                acsees_flag = False #Флаг доступа
                for i in range(0, len(acsees_list)):
                    if data == acsees_list[i]:
                        if ArduinoControl:
                            SendData(2) # Отправим на Ардуино сигнал об этом
                        print(cl.Fore.GREEN + "Доступ разрешён!")
                        acsees_flag = True

                if acsees_flag == False:
                    print(cl.Fore.RED + "Доступ запрещён")

#Отправка данных по последовательному порту на Ардуино
def SendData(x):
    ser.write(x)
    print(cl.Fore.BLACK + f"Bytes: {x}")
    time.sleep(0.05)
    data = ser.readline()
    print(cl.Fore.BLACK + f"Data: {data}")

#Чтение значений из колонки со штрих-кодами
def read_column_from_excel(filename, column_index):
    try:
        wb = openpyxl.load_workbook(filename, read_only=True)
        ws = wb.active
    except FileNotFoundError:
        print(f"Файл {filename} не найден.")
        return None

    column_data = []
    for row in ws.iter_rows(min_col=column_index, max_col=column_index, values_only=True):
        for value in row:
            column_data.append(value)
    print(column_data)
    return column_data

if __name__ == "__main__":
    if ArduinoControl:
        ser = serial.Serial(ArduinoPort, ArduinoSpeed, timeout=0.1)

    art.tprint("Barcode Reader")
    print(cl.Fore.MAGENTA + "Новые данные вводим? Y/N")
    choise = input()
    if choise.lower() == "y":
        print(cl.Fore.BLACK  + "Ок, введите 13 цифр")
        num = input()
        print(cl.Fore.BLACK + "Введите имя")
        name = input()
        BarCodeGenerate(num, name)
    else:
        print(cl.Fore.BLACK  + "Ок, тогда пошли дальше")

    while True:
        BarcodeReader()
        time.sleep(tickTimer)
