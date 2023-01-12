import os
from time import sleep
import pydub
import speech_recognition as sr
import win32com.client as win32
from os import path
from pydub import AudioSegment
from pathlib import Path
from os import remove


# Removendo o TXT transcrito caso já exista
try:
    # Caso ele nao exista, o erro retornado será ignorado
    os.remove(f"{Path.home()}\\Desktop\\Resposta Transcricao\\transcricao.txt")
    os.remove(f"{Path.home()}\\Desktop\\Resposta Transcricao\\audio.wav")

except:
    pass


# Convertendo arquivo mp3 para wav
try:

    src = (f"{Path.home()}\\Desktop\\Resposta Transcricao\\audio.mp3")

    sound = AudioSegment.from_mp3(src)

    sound.export(f"{Path.home()}\\Desktop\\Resposta Transcricao\\audio.wav", format="wav")

except:

# Passando direto caso o arquivo seja wav

    pass



# Pegando o caminho do arquivo na pasta do configurada
AUDIO_FILE = f"{Path.home()}\\Desktop\\Resposta Transcricao\\audio.wav"


# Usar o arquivo de áudio como fonte de áudio
r = sr.Recognizer()
with sr.AudioFile(AUDIO_FILE) as source:
    audio = r.record(source)


# Recognize usando Sphinx
try:
    print("Saida: " + r.recognize_sphinx(audio))
    arquivo = Path.home() / 'Desktop' / 'Resposta Transcricao' / 'transcricao.txt'

    # Escrevendo o retorno do audio no arquivo
    arquivo.touch()
    arquivo.write_text(r.recognize_sphinx(audio))

except sr.UnknownValueError:
    arquivo.write_text("Não foi possivel entender o audio")

except sr.RequestError as e:
    arquivo.write_text("Error; {0}".format(e))



#Comando para converte o script python em um executavel, com todas as bibliotecas "pyinstaller --onefile main.py"

#Caso queira converter em executavel, algumas maquinas exigem as instalação ffmpeg na sua maquina

#pydub.AudioSegment.converter = "C:/Program Files/ffmpeg/bin/ffmpeg.exe"