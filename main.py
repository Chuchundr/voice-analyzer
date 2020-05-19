# -*- coding: utf-8 -*-
import json
import os
import speech_recognition as sr
from datetime import date
from excel import ExcelClass


os.environ['PYTHONIOENCODING'] = 'ASCII'

rcgnzr = sr.Recognizer()

xlsx = ExcelClass()

with open('json_dump.json', 'r') as f:
    json_dump = json.load(f)

def say():
    with sr.Microphone(device_index=1, sample_rate=44100, chunk_size=2048) as source:
        rcgnzr.adjust_for_ambient_noise(source)
        audio = rcgnzr.listen(source)
    text = rcgnzr.recognize_google(audio, language='ru-RU', )
    return text.lower()


while True:
    try:
        print('Говори')
        text = say()
        print('Ты сказал ' + text)
        if text == '1':
            print('Сайт айди')
            text = say().split()
            print(text)
            for key, value in json_dump['a']['cities'].items():
                if key in text:
                    text.pop(text.index(key))
                    full_text = str(value + ''.join(text))
                    xlsx.write_into_cell(1, xlsx.sheet.max_row+1, full_text)
                    print(full_text)
                    xlsx.write_into_cell(3, xlsx.sheet.max_row, date.today().strftime("%d.%m.%Y"))
        if text == '6':
            print('Имя')
            text = say()
            xlsx.write_into_cell(6, xlsx.sheet.max_row, text.capitalize())
            print(text.capitalize())
        if text == '9':
            print('Время начала')
            text = say()
            xlsx.write_into_cell(9, xlsx.sheet.max_row, text)
            print(text)
        if text == 'удали':
            text = ''.join(say().split()).capitalize()
            print(text)
            xlsx.delete(text)
            print('удалил ' + text)
    except sr.UnknownValueError:
        print("Не понял")
    except sr.RequestError as e:
        print('Ошибка при отправке запроса;{0}'.format(e))


