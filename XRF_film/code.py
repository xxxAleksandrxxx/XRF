import csv
from matplotlib import pyplot as plt
from statistics import stdev as stdev
from statistics import mean as mean


'''
f= open('dataset/M-10-10.csv', newline = '')
reader = csv.reader(f, delimiter = ',')
X, Y = [], []
for _ in range(9):     #пропускаем первые строки с описанием условий измерений
    reader.__next__()  #и переходим к данным
for i in reader:
    X.append(int(i[0].strip())/100)
    Y.append(int(i[1].strip())/100)
f.close()
'''
#извлекаем данные из файла
def get_data(file):
    with open(file, newline = '') as f:
        reader = csv.reader(f, delimiter = ',')
        X, Y = [], []
        for _ in range(9):     #пропускаем первые строки с описанием условий измерений
            reader.__next__()  #и переходим к данным
        for i in reader:
            X.append(int(i[0].strip())/100)
            Y.append(int(i[1].strip())/100)
    return X, Y

# диапазон ROI для серы, откуда суммируем интенсивность
ROI_start = 217
ROI_stop = 252 + 1



'''
#График одного батча
#
fig = plt.figure(figsize = (6, 2))
#обсчет повторных измерений
sulphur_int = []
for i in range(1, 11):
    X, Y = get_data(f'dataset/M-1-{i}.csv')
    sulphur_int.append(sum(i for i in Y[ROI_start: ROI_stop]))
    #строим график
    char_1 = fig.add_subplot(121)
    x, y = X[ROI_start:ROI_stop], Y[ROI_start:ROI_stop]
    plt.xlabel('Energy, keV')
    plt.ylabel('Intensity, cps')
    char_1.plot(x, y)
#
print(sulphur_int)
sko_int = stdev(sulphur_int)
print('StDeviation_batch:', sko_int) 
#
#график сумм интенсивностей из ROI
char_2 = fig.add_subplot(122)
char_2.scatter(range(len(sulphur_int)), sulphur_int)
mean_s = mean(sulphur_int)
min_s = min(sulphur_int)-mean_s*0.02
max_s = max(sulphur_int)+mean_s*0.02

char_2.set_ylim([min_s, max_s])  # настраиваем границы по y для интенсивности по S
#
fig.subplots_adjust(bottom = 0.27, wspace = 0.5)  #настройка вывода подграфиков
#
plt.show()
#
'''

#обсчет повторов одной серии
#функция выдает среднее значение серии
def count_repeats(file_directory, file_label, repeats, roi_start, roi_stop):       #передаем имя файла с номером серии, но без номера повтора; также передаем количество повторов и ROI
    fig = plt.figure(figsize = (figsize_x, figsize_y))
    sulphur_int = []
    for i in range(1, repeats + 1):  #считаем интенсивность из области ROI, строим графики
        X, Y = get_data(f'{file_directory}/{file_label}{i}.csv')
        sulphur_int.append(sum(i for i in Y[roi_start: roi_stop])) #добавляем значение суммарной интенсивности
        #график со спектром
        char_1 = fig.add_subplot(121)
        x, y = X[roi_start:roi_stop], Y[roi_start:roi_stop]
        plt.xlabel('Energy, keV')
        plt.ylabel('Intensity, cps')
        char_1.plot(x, y)
    #график с интенсивностями (площадь под графиком спектра)
    char_2 = fig.add_subplot(122)  
    char_2.scatter(range(len(sulphur_int)), sulphur_int)
    mean_s = mean(sulphur_int)
    min_s = min(sulphur_int)-mean_s*0.02
    max_s = max(sulphur_int)+mean_s*0.02
    char_2.set_ylim([min_s, max_s])  # настраиваем границы по y для интенсивности по S
    fig.subplots_adjust(bottom = 0.27, wspace = 0.5)  #настройка вывода подграфиков
    #plt.show()
    print(f'{file_label} mean:', mean_s, 'stdev:', stdev(sulphur_int), )
    plt.savefig(f'charts/{file_label}.jpg')
    return mean_s


#обсчет всех серий
def count_batches(file_directory, file_label, num_batches, num_repeats, roi_start, roi_stop):
    sulphur_all_batches = []
    for i in range(1, num_batches +1):
        sulphur_all_batches.append(count_repeats(file_directory, f'{file_label}-{i}-', num_repeats, roi_start, roi_stop))
    return mean(sulphur_all_batches), stdev(sulphur_all_batches)
                                   
'''
# именования файлов: '{file_directory}/{file_label}-{batch}-{series}.csv'
'''

folder = 'dataset'
#file_name = 'M-1-' для проверки count_repeats
file_name = 'M' #для проверки count_batches

rep = 10
batches = 10
ROI_start = 217
ROI_stop = 252 + 1

figsize_x, figsize_y = 6, 2

names = ['2.5', 'M']
for i in names:
    counted = count_batches(folder, i, batches, rep, ROI_start, ROI_stop)
    
    print('среднее по сериям: ', counted[0])
    print('СКО по сериям:     ', counted[1])
#count_repeats(folder, file_name, rep, ROI_start, ROI_stop)


'''
#графи спектра
plt.subplot(2, 1, 2)
start = ROI_start #150
stop = ROI_stop #281
x, y = X[start:stop], Y[start:stop]
plt.plot(x, y)
plt.xlabel('Energy, keV')
plt.ylabel('Intensity, cps')
plt.grid()
plt.show()
'''
