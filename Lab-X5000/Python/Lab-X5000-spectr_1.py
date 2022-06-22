import json
import codecs
import matplotlib.pyplot as plt
from tkinter import *
from tkinter import filedialog as fd

##прописываем адрес файла самостоятельно
##with codecs.open(r'c:\Users\DAA\Documents\Python\Samples\m82.json',
##                 'r',
##                 encoding='utf-8') as j:
##    data = json.load(j)
                     


##выбираем файл в окне
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
file_name = fd.askopenfilename(filetypes = [('*.json files', '.json')])
with codecs.open(file_name,
                 'r',
                 encoding='utf-8') as j:
    data = json.load(j)

offset = data['spectra'][0]['wavelength_calibration']['positionVsWavelength'][1]
energy_to_channel = data['spectra'][0]['wavelength_calibration']['wavelengthVsPosition'][1]
x = [(offset*(-0.1)+i*energy_to_channel*0.001) for i in range(1, 1+len(data['spectra'][0]['values']))]

###вариант с ограничением максимальной энергии
###в зависимости от энергии возбуждения
###требуется доработка
##x = [(offset+i*energy_to_channel/1000) for i in range(1, 1+data['measurement_conditions']['voltage_kV'])]

plt.plot(x, data['spectra'][0]['values'])
plt.show() #использовать если запускаем сам файл в проводнике
plt.show(block=False) #block=False - не блокирует работу с командной строкой
