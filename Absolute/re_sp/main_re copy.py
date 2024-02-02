from os import system, path, chdir
chdir(path.dirname(path.abspath(__file__))[:-6])
print(path.dirname(path.abspath(__file__))[:-6])



from vosk import Model, KaldiRecognizer
from pyaudio import PyAudio, paInt16
from random import randrange
import Protocols as pr
from os.path import realpath
from os import  path, chdir
import numpy as np
from keyboard import add_hotkey
from pyttsx3 import init
from translate import Translator
import json
import os
from functools import lru_cache
import regex as reg
import re
from requests import get
import tensorflow as tf
from tqdm import tqdm

import webbrowser, os, time, random, sys, datetime
import keyboard as k
import pyautogui as pw
import aspose.words as aw
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from win32comext.shell import shell, shellcon   # pip install pywin32
from words2numsrus import NumberExtractor
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

chdir(path.dirname(path.abspath(__file__))[:-6])
print(path.dirname(path.abspath(__file__))[:-6])



def pppp(promt):
        global prompt
        
        print(promt)
        translator = Translator(from_lang='Russian', to_lang='English')
        prompt = translator.translate(promt.replace('ответь на вопрос',''))
        print(prompt)

        import fire
        fire.Fire(main)
        pr.Voise(perevod(output_text))


methods={
    'ответь на вопрос':pppp,

    "включи музыку":pr.P1,
    "протокол депрессия":pr.P2,
    "протокол суицид":pr.P3,
    'включи браузер':pr.P4,
    'включи телеграмм':pr.P5,
    'включи вконтакте':pr.P6,
    'включи решу егэ':pr.P7,
    'включи стим':pr.P8,
    'включи песню':pr.P9,
    'включи файлы':pr.P10,
    'включи ютуб':pr.P11,
    'включи редактор кода':pr.P12,
    'создай новый текстовый файл':pr.P13,
    'включи камеру':pr.P14,
    'включи твич':pr.P15,
    'выключи звук':pr.P16,
    'включи звук на максимум':pr.P17,
    'половина громкости':pr.P18,
    'включи переводчик':pr.P19,
    'случайное число до десяти':pr.P20,
    'случайное число до ста':pr.P21,
    'очисти корзину':pr.P22,
    'включи генератор презентации':pr.P23,
    'закрой окно':pr.P24,
    'сверни все окна':pr.P25,
    'запусти доту':pr.P26,
    'закрой вкладку':pr.P27,
    'сделай скриншот':pr.P28,
    'None':pr.P29,
    'протокол гуль':pr.P30,
    'включи дискорт':pr.P31,
    'что такое':pr.P32,
    'протокол день рождения':pr.P33,
    'включи торрент':pr.P34,
    'протокол работа':pr.P35,
    'протокол один дома':pr.P36,
    'протокол отдых':pr.P37,
    'включи торрент игруха':pr.P38,
    'включи музыку и решу егэ':pr.P39,
    'выключи систему':pr.P40,
    'спящий режим':pr.P41,
    'отбой':pr.P42,
    'создай новый лист':pr.P43,
    'включи файл передачи':pr.P44,
    'включи фотошоп':pr.P45,
    'презагрузи пк':pr.P46,
    'включи видеоредактор':pr.P47,
    'русская рулетка':pr.P48,
    'смена языка':pr.P49,
    'звук на':pr.P50,
    'None':pr.P51,
    'который час':pr.P52,
    'какая дата':pr.P53,
    'сколько дней до нового года':pr.P54,
    'протокол новый год':pr.P55,
    'включи калькулятор':pr.P56,
    'пауза':pr.P57,
    'найди в интернете':pr.P58,

}





W1=['запусти','включи','активируй','открой']
W2=['громкость','звук']
W3=['сколько время','который час','скажи время']
W4=['какая дата','что за число сегодня','какое число сегодня']
W5=['ответите', 'ответьте','ответьтите','ответить']


phrases_ok = ['есть сэр','готово','сделано','исполненно','']





def simi(promt):
    W0= W1+W2+W3+W4+W5
    for x in range(len(W0)):
        if W0[x] in promt:
            for i in range(len(W1)): promt = promt.replace(W1[i], 'включи')
            for i in range(len(W2)): promt = promt.replace(W2[i], 'звук')
            for i in range(len(W3)): promt = promt.replace(W3[i], 'который час')
            for i in range(len(W4)): promt = promt.replace(W4[i], 'какая дата')
            for i in range(len(W5)): promt = promt.replace(W5[i], 'ответь')
    
    return promt




#pr.P53(nomber, month)
        
def check(promt, methods):
    promt = simi(promt)
    print(promt)
    if any(x in promt for x in methods):
        for x in methods:
            if 'ответь на вопрос' in promt:
                methods['ответь на вопрос'](promt)
                break
            elif x in promt:
                methods[x]()
                break
    elif promt=='':
        pass
    else:
        if all([x !=promt for x in methods]):
            pr.Voise('Протокол не найден')


    


def prosmotr():
    from os import walk, path, chdir
    from psutil import disk_partitions

    chdir('C:/')
    a=disk_partitions()
    chdir(path.dirname(path.abspath(__file__))[:-6])
    cont=[]
    for i in a:
        cont+=str(i)[18]

    list=[]
    list2=[]

    link_programm= ['Telegram.exe','steam.exe','Этот компьютер - Ярлык.lnk','obs64.exe','Discord.exe', 'uTorrent.exe', 'Photoshop.exe','Adobe Premiere Pro.exe','Code.exe','calc.exe']

    for disk in cont:
        for root, dirs, files in walk(disk+':'):  
            for filename in files:
                if filename in link_programm:
                    list.append(filename)
                    list2.append(root)


    f=open(path.realpath('For_Absolute/parametr.txt'), 'a' ,encoding='utf-8')
    for x in link_programm:
        if x in list:
            f.write('\n'+list2[list.index(x)]+'/'+x)
        else:
            f.write('\n'+'0')
    f.close()
    # print(list)
    # print(list2)










@lru_cache()
def bytes_to_unicode():
    """
    Returns list of utf-8 byte and a corresponding list of unicode strings.
    The reversible bpe codes work on unicode strings.
    This means you need a large # of unicode characters in your vocab if you want to avoid UNKs.
    When you're at something like a 10B token dataset you end up needing around 5K for decent coverage.
    This is a significant percentage of your normal, say, 32K bpe vocab.
    To avoid that, we want lookup tables between utf-8 bytes and unicode strings.
    And avoids mapping to whitespace/control characters the bpe code barfs on.
    """
    bs = list(range(ord("!"), ord("~") + 1)) + list(range(ord("¡"), ord("¬") + 1)) + list(range(ord("®"), ord("ÿ") + 1))
    cs = bs[:]
    n = 0
    for b in range(2**8):
        if b not in bs:
            bs.append(b)
            cs.append(2**8 + n)
            n += 1
    cs = [chr(n) for n in cs]
    return dict(zip(bs, cs))


def get_pairs(word):
    """Return set of symbol pairs in a word.
    Word is represented as tuple of symbols (symbols being variable-length strings).
    """
    pairs = set()
    prev_char = word[0]
    for char in word[1:]:
        pairs.add((prev_char, char))
        prev_char = char
    return pairs


class Encoder:
    def __init__(self, encoder, bpe_merges, errors="replace"):
        self.encoder = encoder
        self.decoder = {v: k for k, v in self.encoder.items()}
        self.errors = errors  # how to handle errors in decoding
        self.byte_encoder = bytes_to_unicode()
        self.byte_decoder = {v: k for k, v in self.byte_encoder.items()}
        self.bpe_ranks = dict(zip(bpe_merges, range(len(bpe_merges))))
        self.cache = {}

        # Should have added re.IGNORECASE so BPE merges can happen for capitalized versions of contractions
        self.pat = reg.compile(r"""'s|'t|'re|'ve|'m|'ll|'d| ?\p{L}+| ?\p{N}+| ?[^\s\p{L}\p{N}]+|\s+(?!\S)|\s+""")

    def bpe(self, token):
        if token in self.cache:
            return self.cache[token]
        word = tuple(token)
        pairs = get_pairs(word)

        if not pairs:
            return token

        while True:
            bigram = min(pairs, key=lambda pair: self.bpe_ranks.get(pair, float("inf")))
            if bigram not in self.bpe_ranks:
                break
            first, second = bigram
            new_word = []
            i = 0
            while i < len(word):
                try:
                    j = word.index(first, i)
                    new_word.extend(word[i:j])
                    i = j
                except:
                    new_word.extend(word[i:])
                    break

                if word[i] == first and i < len(word) - 1 and word[i + 1] == second:
                    new_word.append(first + second)
                    i += 2
                else:
                    new_word.append(word[i])
                    i += 1
            new_word = tuple(new_word)
            word = new_word
            if len(word) == 1:
                break
            else:
                pairs = get_pairs(word)
        word = " ".join(word)
        self.cache[token] = word
        return word

    def encode(self, text):
        bpe_tokens = []
        for token in reg.findall(self.pat, text):
            token = "".join(self.byte_encoder[b] for b in token.encode("utf-8"))
            bpe_tokens.extend(self.encoder[bpe_token] for bpe_token in self.bpe(token).split(" "))
        return bpe_tokens

    def decode(self, tokens):
        text = "".join([self.decoder[token] for token in tokens])
        text = bytearray([self.byte_decoder[c] for c in text]).decode("utf-8", errors=self.errors)
        return text


def get_encoder(model_name, models_dir):
    with open(os.path.join(models_dir, model_name, "encoder.json"), "r") as f:
        encoder = json.load(f)
    with open(os.path.join(models_dir, model_name, "vocab.bpe"), "r", encoding="utf-8") as f:
        bpe_data = f.read()
    bpe_merges = [tuple(merge_str.split()) for merge_str in bpe_data.split("\n")[1:-1]]
    return Encoder(encoder=encoder, bpe_merges=bpe_merges)


def download_gpt2_files(model_size, model_dir):
    assert model_size in ["124M", "355M", "774M", "1558M"]
    for filename in [
        "checkpoint",
        "encoder.json",
        "hparams.json",
        "model.ckpt.data-00000-of-00001",
        "model.ckpt.index",
        "model.ckpt.meta",
        "vocab.bpe",
    ]:
        url = "https://openaipublic.blob.core.windows.net/gpt-2/models"
        r = get(f"{url}/{model_size}/{filename}", stream=True)
        r.raise_for_status()

        with open(os.path.join(model_dir, filename), "wb") as f:
            file_size = int(r.headers["content-length"])
            chunk_size = 1000
            with tqdm(
                ncols=100,
                desc="Fetching " + filename,
                total=file_size,
                unit_scale=True,
                unit="b",
            ) as pbar:
                # 1k for chunk_size, since Ethernet packet size is around 1500 bytes
                for chunk in r.iter_content(chunk_size=chunk_size):
                    f.write(chunk)
                    pbar.update(chunk_size)


def load_gpt2_params_from_tf_ckpt(tf_ckpt_path, hparams):
    def set_in_nested_dict(d, keys, val):
        if not keys:
            return val
        if keys[0] not in d:
            d[keys[0]] = {}
        d[keys[0]] = set_in_nested_dict(d[keys[0]], keys[1:], val)
        return d

    params = {"blocks": [{} for _ in range(hparams["n_layer"])]}
    for name, _ in tf.train.list_variables(tf_ckpt_path):
        array = np.squeeze(tf.train.load_variable(tf_ckpt_path, name))
        name = name[len("model/") :]
        if name.startswith("h"):
            m = re.match(r"h([0-9]+)/(.*)", name)
            n = int(m[1])
            sub_name = m[2]
            set_in_nested_dict(params["blocks"][n], sub_name.split("/"), array)
        else:
            set_in_nested_dict(params, name.split("/"), array)

    return params


def load_encoder_hparams_and_params(model_size, models_dir):
    assert model_size in ["124M", "355M", "774M", "1558M"]

    model_dir = os.path.join(models_dir, model_size)
    tf_ckpt_path = tf.train.latest_checkpoint(model_dir)
    if not tf_ckpt_path:  # download files if necessary
        os.makedirs(model_dir, exist_ok=True)
        download_gpt2_files(model_size, model_dir)
        tf_ckpt_path = tf.train.latest_checkpoint(model_dir)

    encoder = get_encoder(model_size, models_dir)
    hparams = json.load(open(os.path.join(model_dir, "hparams.json")))
    params = load_gpt2_params_from_tf_ckpt(tf_ckpt_path, hparams)

    return encoder, hparams, params


def gelu(x):
    return 0.5 * x * (1 + np.tanh(np.sqrt(2 / np.pi) * (x + 0.044715 * x**3)))

def softmax(x):
    exp_x = np.exp(x - np.max(x, axis=-1, keepdims=True))
    return exp_x / np.sum(exp_x, axis=-1, keepdims=True)

def layer_norm(x, g, b, eps: float = 1e-5):
    mean = np.mean(x, axis=-1, keepdims=True)
    variance = np.var(x, axis=-1, keepdims=True)
    return g * (x - mean) / np.sqrt(variance + eps) + b

def linear(x, w, b):
    return x @ w + b

def ffn(x, c_fc, c_proj):
    return linear(gelu(linear(x, **c_fc)), **c_proj)

def attention(q, k, v, mask):
    return softmax(q @ k.T / np.sqrt(q.shape[-1]) + mask) @ v

def mha(x, c_attn, c_proj, n_head):
    x = linear(x, **c_attn)
    qkv_heads = list(map(lambda x: np.split(x, n_head, axis=-1), np.split(x, 3, axis=-1)))
    causal_mask = (1 - np.tri(x.shape[0], dtype=x.dtype)) * -1e10
    out_heads = [attention(q, k, v, causal_mask) for q, k, v in zip(*qkv_heads)]
    x = linear(np.hstack(out_heads), **c_proj)
    return x

def transformer_block(x, mlp, attn, ln_1, ln_2, n_head):
    x = x + mha(layer_norm(x, **ln_1), **attn, n_head=n_head)
    x = x + ffn(layer_norm(x, **ln_2), **mlp)
    return x

def gpt2(inputs, wte, wpe, blocks, ln_f, n_head):
    x = wte[inputs] + wpe[range(len(inputs))]
    for block in blocks:
        x = transformer_block(x, **block, n_head=n_head)
    return layer_norm(x, **ln_f) @ wte.T

def generate(inputs, params, n_head, n_tokens_to_generate):
    from tqdm import tqdm
    for _ in tqdm(range(n_tokens_to_generate), "generating"):
        logits = gpt2(inputs, **params, n_head=n_head)
        next_id = np.argmax(logits[-1])
        inputs.append(int(next_id))
    return inputs[len(inputs) - n_tokens_to_generate :]

def main(n_tokens_to_generate: int = 30, model_size: str = "124M", models_dir: str = "models"):
    global output_text

    encoder, hparams, params = load_encoder_hparams_and_params(model_size, models_dir)
    input_ids = encoder.encode(prompt)
    assert len(input_ids) + n_tokens_to_generate < hparams["n_ctx"]
    output_ids = generate(input_ids, params, hparams["n_head"], n_tokens_to_generate)
    output_text = encoder.decode(output_ids)
    return output_text

    
def perevod(text):

    translator = Translator(from_lang='English', to_lang='Russian')
    text = translator.translate(text)
    print(text)
    return text

def minuse(command):
    sim = ["}","{",":","\n",'"']
    for i  in range(len(sim)):
        command = command.replace(sim[i], '')
    if any([x in command for x in sim]) == True:
        minuse(command)
    else:
        return command


def mikrofone():
    global counters_of_mik
    if counters_of_mik==0:
        counters_of_mik+=1
        pr.Voise('Микрофон выключен')
    else:
        counters_of_mik-=1
        pr.Voise('Микрофон включен')



def strem_packet():
    global command
    model = Model(realpath("re_sp/ru_model")) # полный путь к модели
    rec = KaldiRecognizer(model, 8000)
    p = PyAudio()
    stream = p.open(
        format=paInt16, 
        channels=1, 
        rate=8000, 
        input=True, 
        frames_per_buffer=8000
    )

    stream.start_stream() 
    print('READY')

    data1 = open(realpath("For_Absolute\parametr.txt"), "r+",encoding='utf-8')
    normal_data=data1.readlines()
    if len(normal_data)>=3:
        pr.Voise('Необходимые пути найдены')
    else:
        prosmotr()
    data1.close
    name = normal_data[1][:-1]
    te= '"text"'


    pr.Voise('Здравствуйте, меня зовут {}. Жду ваших указаний сэр'.format(name))
    while True:
        data = stream.read(16000)
        if len(data) == 0:
            break
        command= str(rec.Result() if rec.AcceptWaveform(data) else rec.PartialResult())
        if te in command and name in command and counters_of_mik==0 and command!='':
            # print(realpath("re_sp/text_command_red.txt"))
            command = command.replace('"text"', "") 
            command = command.replace(name, "") 
            command =minuse(command)
            check(command, methods)
            command=''

if __name__ == "__main__":
    counters_of_mik =0
    add_hotkey('ctrl+0', mikrofone)
    strem_packet()
    