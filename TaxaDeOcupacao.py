import pandas as pd
from datetime import datetime, timedelta

class subtrairDatas:

	def __init__(self, tempo1, tempo2):
		if(tempo1 < tempo2):
			self.diff = tempo2 - tempo1
		else:
			self.diff = tempo1 - tempo2		

	def getMinutos(self):
	 	minutos = 0
	 	if(self.diff.days >= 1):
	 		minutos = self.diff.days * 24 * 60
	 	minutos += self.diff.seconds / 60
	 	return minutos

#Crio um dataframe para cada Aba da planilha
dfB1 = pd.read_excel (r'C:\temp\EstadiasPorBerco.xlsx', sheet_name='Berco 1')
dfB2 = pd.read_excel (r'C:\temp\EstadiasPorBerco.xlsx', sheet_name='Berco 2')
dfB3 = pd.read_excel (r'C:\temp\EstadiasPorBerco.xlsx', sheet_name='Berco 3')
dfB4 = pd.read_excel (r'C:\temp\EstadiasPorBerco.xlsx', sheet_name='Berco 4')
#Adiciono o coluna para diferenciar qual berço antes de juntar todos os dados.
dfB1['Berco'] = "Berco 1"
dfB2['Berco'] = "Berco 2"
dfB3['Berco'] = "Berco 3"
dfB4['Berco'] = "Berco 4"
#Gero apenar um dataframe com os dados de todos os berços
dfBercos =  pd.concat([dfB1,dfB2,dfB3, dfB4])
#dfBercos =  pd.concat([dfB1, dfB4])
#Crio uma lista para separar as datas de atracação e desatracacao
#Quando atracar irei somar um navio e desatracar subtrair
listManobras = []
for index, row in dfBercos.iterrows():
	listManobras.append({"DataManobra": row["Inicio"], "qtdNavios": 1})
	listManobras.append({"DataManobra": row['Fim'], "qtdNavios": -1})

dfManobras = pd.DataFrame(listManobras) 
dfManobras = dfManobras.sort_values("DataManobra")

dataAnterior = datetime.strptime('01/05/2019 00:00', '%d/%m/%Y %H:%M')
qtdNavios = 0
listTaxaOcupacao = []
#Percorro todas as manobras para calcular a qtd de minutos por qtd de navios e por dia
for index, row in dfManobras.iterrows():
	#Caso a manobra seja anterior a data de inicio, apenas atualizo a qtd de navios
	if(dataAnterior > row["DataManobra"]):
		qtdNavios += row["qtdNavios"]		
	else:		
		datAtual = row["DataManobra"]		
		diferenca = datAtual-dataAnterior
		#Caso o dia seja o mesmo apenas subtraio uma data da outra  e salvo os minutos
		if(datAtual.day == dataAnterior.day):		
			minutos = subtrairDatas(datAtual, dataAnterior).getMinutos()
			listTaxaOcupacao.append({"dataManobra": dataAnterior, "qtdNavios": qtdNavios, "tempo": minutos})
			print("1 - Anterior: " + dataAnterior.strftime("%d/%m/%Y, %H:%M:%S")  + 
				" Atual: "+ datAtual.strftime("%d/%m/%Y, %H:%M:%S") + "Tempo: " + str(minutos))
			dataAnterior = datAtual
			qtdNavios += row["qtdNavios"]			
		else:			
			#Caso tenha mais de 1 dia sem manobra gero o calculo do tempo dos dias anteriores			
			while (datAtual.day != dataAnterior.day):
				#Busco o proximo dia e zero as horas, minutos e segundos
				datCont = dataAnterior + timedelta(days=1)
				datCont = datCont.replace(hour=0, minute=0, second=0, microsecond=0)
				minutos = subtrairDatas(datCont, dataAnterior).getMinutos()
				listTaxaOcupacao.append({"dataManobra": dataAnterior, "qtdNavios": qtdNavios, "tempo": minutos})
				print("2 - Anterior: " + dataAnterior.strftime("%d/%m/%Y, %H:%M:%S")  + 
				" Atual: "+ datCont.strftime("%d/%m/%Y, %H:%M:%S") + "Tempo: " + str(minutos))
				dataAnterior = datCont					
			#Apos atualizar os dias anteriores gero o calculo do dia atual
			datAtual = row["DataManobra"]
			minutos = subtrairDatas(datAtual, dataAnterior).getMinutos()
			listTaxaOcupacao.append({"dataManobra": dataAnterior, "qtdNavios": qtdNavios, "tempo": minutos})
			print("3 - Anterior: " + dataAnterior.strftime("%d/%m/%Y, %H:%M:%S")  + 
				" Atual: "+ datAtual.strftime("%d/%m/%Y, %H:%M:%S") + "Tempo: " + str(minutos))
			dataAnterior = datAtual
			qtdNavios += row["qtdNavios"]

dfTaxaOcupacao = pd.DataFrame(listTaxaOcupacao) 
dfTaxaOcupacao = dfTaxaOcupacao.sort_values("dataManobra")
#print (dfTaxaOcupacao)

excel = pd.ExcelWriter('teste.xlsx', engine='xlsxwriter')
dfTaxaOcupacao.to_excel(excel, sheet_name='Taxa de Ocupacao')
dfManobras.to_excel(excel, sheet_name='Manobras')
excel.save()