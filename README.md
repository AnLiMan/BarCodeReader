# BarCodeReader
Проект по чтению и генерации штрих-кодов на python. Для генерации штрих-кодов разработан специальный интерфейс: 
![Screen1](https://github.com/AnLiMan/BarCodeReader/blob/main/Screens/Screen1.png)

введённые данные сохраняются в excel-файл:
![Screen1](https://github.com/AnLiMan/BarCodeReader/blob/main/Screens/Screen2.jpg)

а штрих-код сохраняется в виде отдельного файла. Интерфейс разработан в QT-Designer

## Материал для чтения

1. Распознавание штрих-кодов https://www.geeksforgeeks.org/how-to-make-a-barcode-reader-in-python/ 

2. Генерация штрих-кода https://www.geeksforgeeks.org/how-to-generate-barcode-in-python/
  
3. Работа с Ардуино https://pyserial.readthedocs.io/en/latest/shortintro.html

## Установка библиотек
pip install python-barcode

pip install pyzbar

pip install opencv-python

pip install pyserial

pip install numpy

pip install colorama

pip install art

pip install opencv-python

pip install openpyxl

## Дополнительная информация
Скрипт "BarCodeReader.py" является полностью самостоятельным и умеет как генеририровать штрих-коды,
так и распознавать их при помощи камеры. Excel-файл со списком доступа можно заполнять как вручную, 
так и автоматически при помощи "BarCodeInterface.py".
