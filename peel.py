import pandas as pd
import openpyxl

path = ''
wb = openpyxl.load_workbook('/content/drive/MyDrive/WorkColab/peel/剝藥.xlsx')
ws = wb['(2025)4月']

#txt文檔
colspecs = [(0,5), (5,13), (13,64), (64, 72)] #text columns location
m1 = pd.read_fwf('/content/drive/MyDrive/WorkColab/peel/m1.txt',
                 colspecs=colspecs, encoding='big5')
m2 = pd.read_fwf('/content/drive/MyDrive/WorkColab/peel/m2.txt',
                 colspecs=colspecs, encoding='big5')


def dataClean(start, end):
  data = []     #[['item1', 'data1'], ['item2', 'data2'], ...]
  for row in ws[start : end]:
    row_data = []
    for cell in row:
        row_data.append(cell.value)
    data.append(row_data)
  return data

def setCol(df, M):
  df.columns = ['藥盒', '藥碼', '藥名', '周耗量']
  df['藥盒M'] =  M + df['藥盒'].astype(str)
  return df

#**********藥槽與代號**********
data1 = dataClean('a2', 'e100')
data2 = dataClean('g2', 'k77')
data3 = dataClean('g79', 'k92')
data = data1 + data2 + data3

dataX = []
for row in data:
  dataX.append([row[0], row[4]])
df_resp = pd.DataFrame(dataX, columns=['藥盒', '代號'])

pat_cleak ={'詠臻':17.0, '孟庭':'D4', '聖傑':'D5', '宛真':'D6'}
df_resp['代號'] =df_resp['代號'].replace(pat_cleak)
df_resp['藥盒'] =df_resp['藥盒'].astype('int32')

#**********代號與負責人**********
respList = dataClean('m2', 'n31')
df_list = pd.DataFrame(respList, columns=['代號', '負責人'])
df_clerk = pd.DataFrame({'代號':['D1', 'D2' , 'D3', 'D4', 'D5' , 'D6'],
                         '負責人':['姿樺', '乃綸' , '庭伃', '孟庭' , '聖傑', '宛真']})
df_list = pd.concat([df_list, df_clerk], ignore_index=True)
df_list.drop_duplicates(keep='last', inplace=True)

#**********藥包機周耗量**********
M1 = setCol(m1,'M1- ')
M2 = setCol(m2,'M2- ')
df = pd.concat([M1, M2])
df.dropna(axis=0, inplace=True)

peel = pd.merge(df, df_resp, how='left', on='藥盒')
peel.dropna(axis=0, inplace=True)
peel['周耗量'] =  peel['周耗量'].astype(int)

result = pd.merge(peel, df_list, how='inner', on='代號')
result['代號'] = result['代號'].astype(str).str.split('.').str.get(0)
# result[['藥盒M', '藥名', '耗量', '負責人']].to_csv('peel.csv', index=0)
pat = {'1':'01', '2':'02', '3':'03', '4':'04',
       '5':'05', '6':'06', '7':'07', '8':'08', '9':'09'}
result['代號'] = result['代號'].replace(pat)

result.sort_values(by=['藥盒', '藥盒M'], inplace=True)
result[['藥盒M', '藥名', '代號', '周耗量', '負責人']].to_html('peel.html',
                                              header=True, encoding='utf-8' , index=0)