'''
Date: 2022.02.21
Part 1: Mann kendall test
Part 2: Moving t-test
'''


# Import moduls
import pandas as pd 
import matplotlib.pyplot as plt
from pandas import DataFrame
import numpy as np
# 1 MK test
# 1.1 MK test
#  "Datasets.xlsx"

df = pd.read_excel('dataset.xlsx',engine='openpyxl') 
df.head()
data = df['data']
y = data.to_list()

n = len(data)
Sk = np.zeros(n)
UFk = np.zeros(n)

s = 0 
for i in range(2,n):
    for j in range(1,i):
        if y[i]>y[j]:
            s += 1
    Sk[i] = s
    E = i * (i - 1) / 4
    Var = i * (i - 1) * (2 * i + 5) / 72
    UFk[i] = (Sk[i] - E) / np.sqrt(Var)

y2 = np.zeros(n)
Sk2 = np.zeros(n)
UBk = np.zeros(n)
s = 0
y2 = y[::-1]
for i in range(2,n):
    for j in range(1,i):
        if y2[i] < y2[j]:
            s += 1
    Sk2[i] = s
    E = i * (i - 1) / 4
    Var = i *(i - 1)*(2 * i + 5) / 72
    UBk[i] = -(Sk2[i] - E) / np.sqrt(Var)

UBk2 = UBk[::-1]

# 1.2 plt UF UK
plt.figure(figsize=(10,5))
plt.plot(range(1 ,n+1),UFk,label = 'UFk',color = 'orange')
plt.plot(range(1 ,n+1),UBk2,label = 'UBk',color = 'cornflowerblue')
plt.ylabel('UFk-UBk')
x_lim = plt.xlim()
plt.plot(x_lim,[-1.96,-1.96],'m--',color = 'r')
plt.plot(x_lim, [0,0],'m--')
plt.plot(x_lim,[1.96,1.96],'m--',color = 'r')
plt.show()

# 'UF':UFk,'UB':UBk2 
# 1.3
print(UFk)

# 2 Moving t-test
from matplotlib import pyplot as plt
from tqdm import tqdm
import pandas as pd
import numpy as np
import os
import xlwt

# v:degree of freedom
def get_tvalue(v, sig_level):
    t_values = pd.read_excel('t_values.xlsx',engine='openpyxl')
    return t_values[t_values['n'] == v][sig_level].values

# Pass in the time series and the data to be checked
def huaTTest(data, step, sig_level):
    datacount = len(data)
    v = 2*step-2  
    t_value = get_tvalue(v, sig_level)  # t value
    n1 = step
    n2 = step
    t = np.zeros(len(data))
    c = 1.0 / n1 + 1.0 / n2

    for i in range(step-1, datacount-step):
        data1 = data[i-step+1:i+1]
        data2 = data[i+1:i+step+1]
        # mean
        x1_mean = data1.mean()
        x2_mean = data2.mean()
        # variance
        s1 = data1.var()
        s2 = data2.var()
        sw2 = (n1*s1 + n2*s2)/(n1+n2-2.0)
        t[i - step + 1] = (x1_mean - x2_mean) / np.sqrt(sw2 * c)
    return t, t_value

# read data
df = pd.read_excel('dataset.xlsx',engine='openpyxl') 

time = np.array(df['Year'])
data = np.array(df['data'])
datacount = len(data)

# step range
if datacount % 2 == 0:
    steps = range(2, int(datacount/2))
else:
    steps = range(2, int(datacount/2+1))

dict = {}  # sig_level
# 按照可取步长计算
for step in tqdm(steps, ncols=80, desc=u'滑动T检验进度'):
    t, t_value = huaTTest(data, step, 0.05)

    t = t[t != 0]

    sig_values = []  # significant values
    for i in range(len(t)):
        if np.abs(t[i]) > t_value:
            sig_values.append(time[i + step - 1])
    if sig_values:
        # put results to dictionary
        dict_key = 'step'+str(step)
        dict[dict_key] = sig_values

# write excel
workbook = xlwt.Workbook()
sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)

keys_list = list(dict)
for head in range(len(dict.keys())):
    sheet1.write(0, head, label=keys_list[head])

value_at_index = list(dict.values())

max_row = max([len(value_at_index[i]) for i in range(len(value_at_index))])
for row in range(max_row):
    for col in range(len(keys_list)):
        if row < len(value_at_index[col]):
            sheet1.write(row+1, col, label=int(value_at_index[col][row]))  # 将np.int32转换为int

# write to xls
workbook.save(r'steps___&.xls')

