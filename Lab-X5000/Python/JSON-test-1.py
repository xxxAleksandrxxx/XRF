##import json
##import codecs
##with codecs.open(r'c:\Users\DAA\Documents\Python\Samples\json-test-data-1.json',
##                 'r',
##                 encoding='utf-8') as j:
##    data = json.load(j)
##    print(data)

#этот код работает:
##import json
##import codecs
##with codecs.open(r'c:\Users\DAA\Documents\Python\Samples\json-test-data-1.json',
##                 'r',
##                 encoding='utf-8') as j:
##    data = json.load(j)
##print(data)
##print("data['article'][0]['id']: ", data['article'][0]['id'])

#этот код работает:
##import json
##import codecs
##with codecs.open(r'c:\Users\DAA\Documents\Python\Samples\m2802-short.json',
##                 'r',
##                 encoding='utf-8') as j:
##    data = json.load(j)
##print(data)

import json
import codecs
import matplotlib.pyplot as plt
with codecs.open(r'c:\Users\DAA\Documents\Python\Samples\m82.json',
                 'r',
                 encoding='utf-8') as j:
    data = json.load(j)
offset = data['spectra'][0]['wavelength_calibration']['wavelengthVsPosition'][0]
energy_to_channel = data['spectra'][0]['wavelength_calibration']['wavelengthVsPosition'][1]
x = [(offset+i*energy_to_channel/1000) for i in range(1, 1+len(data['spectra'][0]['values']))]
###вариант с ограничением максимальной энергии
###в зависимости от энергии возбуждения
##x = [(offset+i*energy_to_channel/1000) for i in range(1, 1+data['measurement_conditions']['voltage_kV'])]

plt.plot(x, data['spectra'][0]['values'])
plt.show()

##import codecs
##with codecs.open('c:\Users\DAA\Documents\Python\Samples\m2802.json', encoding='utf-8') as data:
##    print(data)

##import codecs
##file_name = codecs.open(r'c:\Users\DAA\Documents\Python\Samples\m2802-short.json',
##                        encoding='utf-8')
##data = file_name.read()
##print(data)
##file_name.close()

##import json
##data = {
##    "president": {
##        "name": "Zaphod Beeblebrox",
##        "species": "Betelgeusian"
##    }
##}
##
##with open(r'c:\Users\DAA\Documents\Python\Samples\data_file1.json',
##          'w') as write_file:
##	  json.dump(data, write_file)
