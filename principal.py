import csv
import pandas as pd
from Model import Data
import numpy as np
import matplotlib.pyplot as plt


nameOfSourceArchive  = "dados.xlsx"
realizadoSheet = pd.read_excel(nameOfSourceArchive,usecols=list(range(1,13)))
monthList = []
orcadoSheet = pd.read_excel(nameOfSourceArchive,index_col=None,sheet_name=1)
dataList = {}
for number in range(12):
    data = Data()
    data.feito = realizadoSheet.values[1][number]  
    dataList[realizadoSheet.values[0][number]+""] = data
    monthList.append(orcadoSheet.values[number][0]+"")
    dataList[orcadoSheet.values[number][0]+""].orcado = orcadoSheet.values[number][1]
with open('relatorio.csv', 'w') as csvfile:
    filewriter = csv.writer(csvfile, delimiter=',',quotechar='|', quoting=csv.QUOTE_MINIMAL)
    filewriter.writerow(['Mês', 'Orçado','Realizado','Diferença'])
    for month in monthList:
        filewriter.writerow([month.capitalize(), dataList[month].orcado,dataList[month].feito, dataList[month].getDifference()])
print("Arquivo CSV Gerado =)")

N = 12
listLucro = []
listFeito  = []
for month in monthList:
    listLucro.append(0 if dataList[month].getDifference()<0 else dataList[month].getDifference())
    listFeito.append(dataList[month].feito)
ind = np.arange(N)    
width = 0.7      


plt.figure().set_figwidth(25)
p1 = plt.bar(ind, listLucro, width,bottom=listFeito)
p2 = plt.bar(ind, listFeito, width)
plt.legend((p1[0], p2[0]), ('Orçado', 'Realizado'))


plt.ylabel('$')
plt.xlabel("Mês")
plt.title('Gráfico Orçamento')
plt.xticks(ind, monthList)
plt.yticks(np.arange(0, 450, 100))
plt.savefig('grafico.png')
print("Grafico Gerado =)")
plt.show()