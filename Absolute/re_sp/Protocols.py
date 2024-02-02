import webbrowser, os, time, random, sys,datetime
import keyboard as k
import pyautogui as pw
import aspose.words as aw
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pyttsx3 import init
from win32comext.shell import shell, shellcon   # pip install pywin32
from words2numsrus import NumberExtractor
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

def Voise(phrase):
    data = open("For_Absolute/parametr.txt", "r",encoding='utf-8')
    normal_data=data.readlines()
    data.close
    case_gender=['Ирина','Александр','Анна',"Елена"]
    number = case_gender.index(normal_data[0].replace('\n',''))
    voises = ['','Aleksandr','Anna','Elena']
    tts = init()
    voices = tts.getProperty('voices')
    tts.setProperty('rate', int(normal_data[2].replace('\n',''))+125)  # 150 words per minute
    tts.setProperty('volume', int(normal_data[2].replace('\n',''))/100)  # 80% volume
    for voice in voices:
        if voice.name == voises[int(number)]:
            tts.setProperty('voice', voice.id)
    tts.say(phrase)
    tts.runAndWait()



#Запуск музыки
def P1(command):
    data = open("For_Absolute/parametr.txt", "r",encoding='utf-8')
    normal_data=data.readlines()
    data.close()
    if normal_data[4].replace('\n','')=='':
        o = Options()
        o.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options=o)
        driver.get('https://vk-save.com/playlist625975121_15_ef62918a4863814209?back_url=%2Ffave')
        button = driver.find_element(By.CLASS_NAME, 'ext_action_btn_title.play')
        button.click()
    else:
        webbrowser.open(normal_data[4].replace('\n',''))

#Протокол депрессия
def P2(command):
    P1(command)
    P8(command)
#Протокол суицид
def P3(command):
    os.system("shutdown /p")
#Открой браузер
def P4(command):
    webbrowser.open("https://www.google.ru/?hl=ru")

#Открой телеграмм
def P5(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[5])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1]) # Выполнение команды
    else:
        Voise('Данная программа не найдена') # Ответ об ошибке
#Открой в контакте
def P6(command):
    webbrowser.open("https://vk.com/feed")
#Открой решу егэ
def P7(command):
    webbrowser.open("https://phys-ege.sdamgia.ru/?redir=1")
#Открой стим
def P8(command):
    
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[6])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

#Включи песню
def P9(command):

    non_txt = "включи песню"
    quest = command.replace(non_txt, '')
    o = Options()
    o.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=o)
    driver.get('https://vk-save.com')
    promt = driver.find_element(By.XPATH, '//input')
    promt.click()
    promt.send_keys(quest)
    time.sleep(4)
    song = driver.find_element(By.CLASS_NAME, 'audio')
    song.click()
#Открой файлы
def P10(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[7])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

#Открой ютуб
def P11(command):
    webbrowser.open("https://www.youtube.com")
#Открой редактор кода
def P12(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[13])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

#Создай новый текстовый файл
def P13(command):
    doc = aw.Document()
    Count_of_name = str(random.randrange(1,1000,1))
    name = 'test_file_№'+ Count_of_name +'.docx'
    doc.save(name)
    os.startfile(name)
# Открой камеру
def P14(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[8])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

#Открой твич
def P15(command):
    webbrowser.open("https://www.twitch.tv")
#Выключи звук
def P16(command):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(0, None) #Где 0.5 это 50% громкость
#Включи звук на максимум
def P17(command):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(1, None) #Где 0.5 это 50% громкость
#Половина громкости
def P18(command):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(0.5, None) #Где 0.5 это 50% громкость
#Открой переводчик
def P19(command):
    webbrowser.open("https://translate.yandex.ru/?lang=ru-en")
#Случайное число до 10
def P20(command):    
    Voise(str(random.randrange(0,10,1)))
#Случайное число до 100
def P21(command): 
    Voise(str(random.randrange(0,100,1)))
#Открой корзину
def P22(confirm=False, show_progress=True, sound=False):
    flags = 0
    if not confirm:
        flags |= shellcon.SHERB_NOCONFIRMATION
    if not show_progress:
        flags |= shellcon.SHERB_NOPROGRESSUI
    if not sound:
        flags |= shellcon.SHERB_NOSOUND
    shell.SHEmptyRecycleBin(None, None, flags)
#Открой генератор презентации
def P23(command):
    webbrowser.open("https://gamma.app")
#Закрой окно
def P24(command):
    pw.hotkey('alt', 'f4')
#Сверни все окна
def P25(command):
    pw.hotkey('winleft', 'd')
#Запусти доту
def P26(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[6])
    file.close
    if link[0]!='0':
        os.startfile('steam://rungameid/570')
    else:
        Voise('Данная программа не найдена')

#Закрой вкладку
def P27(command):
    pw.hotkey('ctrl', 'w')
#Сделай скриншот
def P28(command):
    pw.screenshot('Screen_001.png')
    os.startfile(os.path.realpath('Screen_001.png'))
#протоколь халяль гитлер
def P29(command):
    pass
#Протокол гуль
def P30(command):
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
    os.startfile(os.path.realpath('/For_Absolute/goul.py'))
#Открой дискорд
def P31(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[9])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

#Что такое "(ввод слова)"
def P32(command):

    non_txt = "что такое"
    quest = command.replace(non_txt, '')
    webbrowser.open("https://www.google.com/search?q="+quest)

#Протокол день рождение
def P33(command):
    P1(command)
    P8(command)
    Voise('С днем рождения Капитан')
#Открой торрент
def P34(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[10])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

#Протокол работа
def P35(command):
    P1(command)
    P12(command)
#Протокол один дома
def P36(command):
    P1(command)
    P8(command)
#Протокол отдых
def P37(command):
    P1(command)
#Запусти торент игруха
def P38(command):
    webbrowser.open('https://itorrents-igruha.org/newgames/')
#запусти музыку и решу егэ
def P39(command):
    P1(command)
    P7(command)

#Выключи систему
def P40(command):
    os.system("shutdown /p")
#Спящий режим
def P41(command):
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
#Абсолют отбой
def P42(command):
    sys.exit(1)
#Создай новый текстовый файл
def P43(command):
    doc = aw.Document()
    Count_of_name = str(random.randrange(1,1000,1))
    name = 'test_file_№'+ Count_of_name +'.txt'
    doc.save(name)
    os.startfile(name)
#Открой файл передачи
def P44(command):
    os.startfile('For_Absolute/helper_file_for_speak.txt')
#открой Фотошоп
def P45(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[11])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

# Презагрузи пк
def P46(command): 
   os.system("shutdown /r")

#открой видеоредактор
def P47(command):
    file = open('For_Absolute/parametr.txt','r+',encoding='utf-8')
    a=file.readlines()
    link=(a[12])
    file.close
    if link[0]!='0':
        os.startfile(link[:-1])
    else:
        Voise('Данная программа не найдена')

#русская рулетка
def P48(command):
    d = random.randrange(1,7,1)
    if d == 6:
        os.system("shutdown /p")    

#Смена языка
def P49(command):
    pw.hotkey('shiftleft', 'alt')

#Громкость на 75
def numberr(quest):
    extractor = NumberExtractor()
    proc = extractor.replace_groups(quest)
    proc=int(proc)
    return proc

def P50(command):

    non_txt = "звук на"
    quest = command.replace(non_txt, '')
    if 'максимум' in quest:
        proc = 1
    elif numberr(quest)<=100:
        extractor = NumberExtractor()
        proc = extractor.replace_groups(quest)
        proc=int(proc)/100
    else:
        Voise('Громкость не задана')
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(proc, None) #Где 0.5 это 50% громкость

#
def P51(command):
    pass
#Который час
def P52(command):

    time= datetime.datetime.now()
    time = str(time.hour) + " часов " +str(time.minute) +' минут '+str(time.second) + 'секунд'
    Voise(time)
#Какая дата? 
def P53(command):
    
    nomber=['первое','второе','третье','четвертое','пятое','шестое','седьмое','восьмое','девятое','десятое','одинадцатое','двенадцатое','тринадцатое','четырнадцатое','пятнадцатое','шестнадцатое','семнадцатое','восемнадцатое','девятнадцатое','двацатое','двадцать первое','двадцать второе','двадцать третье','двадцать четвертое','двадцать пятое','двадцать шестое','двадцать седьмое','двадцать восьмое','двадцать девятое','тридцатое','тридцать первое']
    month = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря']

    date= str(datetime.datetime.now().date())
    date = date.replace('-',' ').split()
    print(date)
    today =nomber[int(date[2])-1]+ month[int(date[1])-1]+ date[0]
    Voise(today)

#Сколько дней до Нового года
def P54(command):
    now = datetime.datetime.today()
    NY = datetime.datetime(2025, 1, 1)
    d = NY-now                   
    mm, ss = divmod(d.seconds, 60)
    hh, mm = divmod(mm, 60)
    print('До нового года: {} дней {} часов {} минут {} секунд.'.format(d.days, hh, mm, ss))
    Voise('До нового года: {} дней {} часов {} минут {} секунд.'.format(d.days, hh, mm, ss))
#Протокол Новый год
def P55(command):
    P1(command)
    Voise('C новым годом сэр')
#Открой калькулятор
def P56(command):
    file = open(os.path.realpath('For_Absolute/parametr.txt'),'r+',encoding='utf-8')
    a=file.readlines()
    link=(a[14])
    file.close
    if link[0]!='0':
        os.startfile(link)
    else:
        Voise('Данная программа не найдена')
#пауза
def P57(command):
    k.send('play/pause media')

#Найди в интернете
def P58(command):

    non_txt = "найди в интернете"
    quest = command.replace(non_txt, '')
    webbrowser.open("https://www.google.com/search?q="+quest )

