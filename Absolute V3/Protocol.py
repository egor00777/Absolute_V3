import webbrowser, os, time, random,datetime
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

from main_core import voice_m, voulme_m, speed_m, music_m
def ender_phrase():
    phrases_ok = ['есть сэр','готово','сделано','исполненно',' ','так точно','выполнено', 'как прикажите', "как пожелаете", "прошу", " "]
    Voise(phrases_ok[random.randrange(1,len(phrases_ok),1)])

def Voise(phrase):
    case_gender=['Ирина','Александр','Анна',"Елена"]
    number = case_gender.index(voice_m.replace('\n',''))
    voises = ['','Aleksandr','Anna','Elena']
    tts = init()
    voices = tts.getProperty('voices')
    tts.setProperty('rate', speed_m+125)  # 150 words per minute
    tts.setProperty('volume', voulme_m/100)  # 80% volume
    for voice in voices:
        if voice.name == voises[int(number)]:
            tts.setProperty('voice', voice.id)
    tts.say(phrase)
    tts.runAndWait()


def P0(command,list_programm):
    from GPT import gpt
    gpt(command)



#Запуск музыки
def P1(command,list_programm):
    if music_m=='':
        o = Options()
        o.add_experimental_option("detach", True)

        o.add_argument("--disable-infobars")
        o.add_argument("start-maximized")
        o.add_argument("--disable-extensions")
        # Pass the argument 1 to allow and 2 to block
        o.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 1})
        driver = webdriver.Chrome(options=o)
        driver.get('https://vk-save.com/playlist625975121_22_8d72a7092165bbd32a?back_url=%2F%3Fq%3DMode%253A%2540%2524%2523%2525%255Elast%2520year%2520of%2520life%2540%2524%2523%2525%255E')
        button = driver.find_element(By.CLASS_NAME, 'ext_action_btn_title.play')
        button.click()
        windows = driver.window_handles
        driver.switch_to.window(windows[1])
        driver.close()
    else:
        webbrowser.open(music_m)

#Протокол депрессия
def P2(command,list_programm):
    link=(list_programm[1])
    if link[0]!='0':
        os.startfile(link)
        P1(command,list_programm)
        ender_phrase()
    else:
        Voise('Невозможно выполнить протокол')

#Протокол суицид
def P3(command,list_programm):
    os.system("shutdown /p")
    Voise('отключение через 3')
    time.sleep(1)
    Voise('2')
    time.sleep(1)
    Voise('1')

#Открой браузер
def P4(command,list_programm):
    webbrowser.open("https://www.google.ru/?hl=ru")
    ender_phrase()
#Открой телеграмм
def P5(command,list_programm):
    link=(list_programm[0])
    if link[0]!='0':
        os.startfile(link) # Выполнение команды
        ender_phrase()
    else:
        Voise('Данная программа не найдена') # Ответ об ошибке

#Открой вконтакте
def P6(command,list_programm):
    webbrowser.open("https://vk.com/feed")
    ender_phrase()
#Открой решу егэ
def P7(command,list_programm):
    webbrowser.open("https://phys-ege.sdamgia.ru/?redir=1")
    ender_phrase()
#Открой стим
def P8(command,list_programm):

    link=(list_programm[1])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

#Включи песню
def P9(command,list_programm):

    non_txt = "включи песню"
    quest = command.replace(non_txt, '')
    o = Options()
    o.add_experimental_option("detach", True)
    o.add_argument("--disable-infobars")
    o.add_argument("start-maximized")
    o.add_argument("--disable-extensions")
    # Pass the argument 1 to allow and 2 to block
    o.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 1})
    driver = webdriver.Chrome(options=o)
    driver.get('https://vk-save.com')
    promt = driver.find_element(By.XPATH, '//input')
    promt.click()
    windows = driver.window_handles
    driver.switch_to.window(windows[1])
    driver.close()
    driver.switch_to.window(windows[0])
    promt.send_keys(quest)
    time.sleep(4)
    song = driver.find_element(By.CLASS_NAME, 'audio')
    song.click()

#Открой файлы
def P10(command,list_programm):
    link=(list_programm[2])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        os.startfile('C:')

#Открой ютуб
def P11(command,list_programm):
    webbrowser.open("https://www.youtube.com")
    ender_phrase()
#Открой редактор кода
def P12(command,list_programm):

    link=(list_programm[8])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

#Создай новый текстовый файл
def P13(command,list_programm):
    doc = aw.Document()
    Count_of_name = str(random.randrange(1,1000,1))
    name = 'test_file_№'+ Count_of_name +'.docx'
    doc.save(name)
    os.startfile(name)
    ender_phrase()
# Открой камеру
def P14(command,list_programm):

    link=(list_programm[3])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

#Открой твич
def P15(command,list_programm):
    webbrowser.open("https://www.twitch.tv")
    ender_phrase()
#Выключи звук
def P16(command,list_programm):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(0, None) #Где 0 это 0% громкость
    ender_phrase()


# Отправить уведомление
def P17(command,list_programm):
    pass
    
#Половина громкости
def P18(command,list_programm):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(0.5, None) #Где 0.5 это 50% громкость
    ender_phrase()
#Открой переводчик
def P19(command,list_programm):
    webbrowser.open("https://translate.yandex.ru/?lang=ru-en")
    ender_phrase()

#Случайное число до 10
def P20(command,list_programm):    
    Voise(str(random.randrange(0,10,1)))
#Случайное число до 100
def P21(command,list_programm): 
    Voise(str(random.randrange(0,100,1)))

#Очисти корзину
def P22(confirm=False, show_progress=True, sound=False):
    flags = 0
    if not confirm:
        flags |= shellcon.SHERB_NOCONFIRMATION
    if not show_progress:
        flags |= shellcon.SHERB_NOPROGRESSUI
    if not sound:
        flags |= shellcon.SHERB_NOSOUND
    shell.SHEmptyRecycleBin(None, None, flags)
    ender_phrase()
#Открой генератор презентации
def P23(command,list_programm):
    webbrowser.open("https://gamma.app")
    ender_phrase()
#Закрой окно
def P24(command,list_programm):
    pw.hotkey('alt', 'f4')
    ender_phrase()
#Сверни все окна
def P25(command,list_programm):
    pw.hotkey('winleft', 'd')
    ender_phrase()
#Запусти доту
def P26(command,list_programm):

    link=(list_programm[1])
    if link[0]!='0':
        os.startfile('steam://rungameid/570')
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

#Закрой вкладку
def P27(command,list_programm):
    pw.hotkey('ctrl', 'w')
    ender_phrase()
#Сделай скриншот
def P28(command,list_programm):
    pw.screenshot('Screen_001.png')
    os.startfile(os.path.realpath('Screen_001.png'))
    ender_phrase()

# погода
def P29(command,list_programm):
    import requests
    import json
    from main_core import city_m
    city = city_m
    url = 'https://api.openweathermap.org/data/2.5/weather?q='+city+'&units=metric&lang=ru&appid=79d1ca96933b0328e1c7e3e7a26cb347'
    weather_data = requests.get(url).json()
    weather_data_structure = json.dumps(weather_data, indent=2)
    temperature = round(weather_data['main']['temp'])
    # выводим значения на экран
    Voise('Сейчас в городе' +city+" "+ str(temperature)+ '°')

# включи аниме
def P30(command,list_programm):
     webbrowser.open("https://anilib.me/?section=all-updates")
     ender_phrase()
#Открой дискорд
def P31(command,list_programm):

    link=(list_programm[4])

    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

#Найди в интернете
def P32(command,list_programm):
    non_txt = "найди в интернете"
    quest = command.replace(non_txt, '')
    webbrowser.open("https://www.google.com/search?q="+quest )
    ender_phrase()
    

#Протокол день рождение
def P33(command,list_programm):
    P2(command,list_programm)
    Voise('С днем рождения Капитан')
#Открой торрент
def P34(command,list_programm):
    print(list_programm)
    link=(list_programm[5])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

#Протокол работа
def P35(command,list_programm):
    link=(list_programm[8])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')
    P1(command,list_programm)
    
#Протокол один дома
def P36(command,list_programm):
    P1(command,list_programm)
    link=(list_programm[1])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')
#Протокол отдых
def P37(command,list_programm):
    P1(command,list_programm)
#Запусти торент игруха
def P38(command,list_programm):
    webbrowser.open('https://itorrents-igruha.org/newgames/')
    ender_phrase()
#запусти музыку и решу егэ
def P39(command,list_programm):
    P1(command,list_programm)
    webbrowser.open("https://phys-ege.sdamgia.ru/?redir=1")

#Выключи систему
def P40(command,list_programm):
    
    Voise('отключение через 3')
    time.sleep(1)
    Voise('2')
    time.sleep(1)
    Voise('1')
    os.system("shutdown /p")
#Спящий режим
def P41(command,list_programm):
    os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
#Абсолют отбой
def P42(command,list_programm):
    Voise('отключение программы через 3')
    time.sleep(1)
    Voise('2')
    time.sleep(1)
    Voise('1')
    os.kill(os.getpid(),9)

#Создай новый текстовый файл
def P43(command,list_programm):
    doc = aw.Document()
    Count_of_name = str(random.randrange(1,1000,1))
    name = 'test_file_№'+ Count_of_name +'.txt'
    doc.save(name)
    os.startfile(name)
    ender_phrase()
# включи вотсап
def P44(command,list_programm):
     webbrowser.open("https://web.whatsapp.com/")
     ender_phrase()
#открой Фотошоп
def P45(command,list_programm):

    link=(list_programm[6])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

# Презагрузи пк
def P46(command,list_programm): 
    Voise('перезагрузка через 3')
    time.sleep(1)
    Voise('2')
    time.sleep(1)
    Voise('1')
    os.system("shutdown /r")

#открой видеоредактор
def P47(command,list_programm):

    link=(list_programm[7])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')

#русская рулетка
def P48(command,list_programm):
    d = random.randrange(1,7,1)
    if d == 6:
        os.system("shutdown /p")    
    else: Voise(str(d)+" не повезло")

#Смена языка
def P49(command,list_programm):
    pw.hotkey('shiftleft', 'alt')
    ender_phrase()
#Громкость на 75
def numberr(quest):
    extractor = NumberExtractor()
    proc = extractor.replace_groups(quest)
    proc=int(proc)
    return proc
def P50(command,list_programm):

    non_txt = "звук на"
    quest = command.replace(non_txt, '')
    if 'максимум' in quest:
        proc = 1
    elif "половину" in quest:
        proc = 0.5
    elif 'четверть' in quest:
        proc = 0.25
    elif 'две четвери' in quest:
        proc = 0.5
    elif 'три четвери' in quest:
        proc =0.75

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
    volume.SetMasterVolumeLevelScalar(proc, None) 
    ender_phrase()
# включи почту
def P51(command,list_programm):
     webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
     ender_phrase()
#Который час
def P52(command,list_programm):
    time= datetime.datetime.now()
    time = str(time.hour) + " часов " +str(time.minute) +' минут '+str(time.second) + 'секунд'
    Voise(time)
#Какая дата? 
def P53(command,list_programm):
    
    nomber=['первое','второе','третье','четвертое','пятое','шестое','седьмое','восьмое','девятое','десятое','одинадцатое','двенадцатое','тринадцатое','четырнадцатое','пятнадцатое','шестнадцатое','семнадцатое','восемнадцатое','девятнадцатое','двацатое','двадцать первое','двадцать второе','двадцать третье','двадцать четвертое','двадцать пятое','двадцать шестое','двадцать седьмое','двадцать восьмое','двадцать девятое','тридцатое','тридцать первое']
    month = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря']

    date= str(datetime.datetime.now().date())
    date = date.replace('-',' ').split()
    print(date)
    today =nomber[int(date[2])-1]+ month[int(date[1])-1]+ date[0]
    Voise(today)

#Сколько дней до Нового года
def P54(command,list_programm):
    now = datetime.datetime.today()
    NY = datetime.datetime(2025, 1, 1)
    d = NY-now                   
    mm, ss = divmod(d.seconds, 60)
    hh, mm = divmod(mm, 60)
    Voise('До нового года: {} дней {} часов {} минут {} секунд.'.format(d.days, hh, mm, ss))
#Протокол Новый год
def P55(command,list_programm):
    Voise('C новым годом сэр')
    P1(command,list_programm)
#Открой калькулятор
def P56(command,list_programm):

    link=(list_programm[9])
    if link[0]!='0':
        os.startfile(link)
        ender_phrase()
    else:
        Voise('Данная программа не найдена')
#пауза
def P57(command,list_programm):
    pw.press("playpause")
    ender_phrase()

#Что такое "(ввод слова)"
def P58(command,list_programm):
    non_txt = "что такое"
    quest = command.replace(non_txt, '')
    webbrowser.open("https://www.google.com/search?q="+quest)
    ender_phrase()
    

#Предыдущий трек
def P59(command,list_programm):
    pw.press("prevtrack")
    ender_phrase()

#Следущий трек
def P60(command,list_programm):
    pw.press("nexttrack")
    ender_phrase()