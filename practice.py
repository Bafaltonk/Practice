import os
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import speech_recognition as sr
import webbrowser as wb
import subprocess as sp

a = str()

while a != 'закрыть программу':

    print('Говорите: ')
    mic = sr.Microphone(device_index=2)
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
    try:
        a = (r.recognize_google(audio, language="ru-RU")).lower()
    except:
        pass

    print(a)

    if a == 'открыть диспетчер задач':
        os.startfile('C:\Windows\System32\Taskmgr.exe')
    elif a == 'открыть панель управления':
        os.startfile('C:\Windows\System32\control.exe')
    elif 'поиск в интернете' in a:
        wb.open('https://yandex.ru/search/?lr=213&text={}'.format(a[18:]))
    elif '+' in a:
        a = a.split()
        print("Cумма равна", (int(a[0]) + int(a[2])))
    elif a == 'открыть блокнот host':
        sp.Popen(['notepad.exe', 'C:\windows\system32\drivers\etc\hosts'])
    elif a == 'выключить звук':
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        sound = 100
        proverka = False
        while(proverka == False):
                try:
                    volume.SetMasterVolumeLevel(-sound, None)
                except:
                    sound -= 1
                else:
                    proverka = True
    elif a == 'включить звук':
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volume.SetMasterVolumeLevel(0, None)

print('Программа успешно закрыта.')

