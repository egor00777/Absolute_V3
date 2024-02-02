#!/usr/bin/env python3

from vosk import Model, KaldiRecognizer, SetLogLevel
import sys
import os
import wave

SetLogLevel(0)

if len(sys.argv) == 2:
    wf = wave.open(sys.argv[1], "rb")
else:
    wf = wave.open("decoder-test.wav", "rb")

if wf.getnchannels() != 1 or wf.getsampwidth() != 2 or wf.getcomptype() != "NONE":
    print ("Audio file must be WAV format mono PCM.")
    exit (1)

model = Model(".")
rec = KaldiRecognizer(model, wf.getframerate())
rec.SetWords(True)

while True:
    data = wf.readframes(4000)
    if len(data) == 0:
        break
    if rec.AcceptWaveform(data):
        print(rec.Result())
    else:
        print(rec.PartialResult())

print(rec.FinalResult())
