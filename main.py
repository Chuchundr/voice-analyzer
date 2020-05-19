# -*- coding: utf-8 -*-
import json
import os
import speech_recognition as sr
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
    return text


while True:
    try:
        print('Говори')
        words = say()
        print('Ты сказал ' + words)
        if words.lower() == 'витаминка':
            print('Что писать?')
            data = say().split(' ')
            lower_func = lambda x: x.lower()
            data = list(map(lower_func, data))
            print(data)
            for key, value in json_dump['a']['cities'].items():
                if key in data:
                    data.pop(data.index(key))
                    xlsx.write_into_cell(1, xlsx.sheet.max_row, str(value + ''.join(data)))
                    print(value)
    except sr.UnknownValueError:
        print("Не понял")
    except sr.RequestError as e:
        print('Ошибка при отправке запроса;{0}'.format(e))


