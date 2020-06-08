import re, glob, os, openpyxl
from collections import OrderedDict

def walk_it():
	data_ret=openpyxl.load_workbook('/Users/Ileam Montoya/Dropbox/Claro/Optimization - Daily/Data retrieval/dashboard.xlsx')
	tabs=['DSLAMS > 80','Interfaces Recurrentes > 95','Interfaces Recurrentes >70< 95','Enlaces > 40','ServiceApp > 40','Hubspoke > 80%', 'Interfaces Fotonico > 95']
	week_dicts={'01':7,'02':8,'03':9,'04':10,'05':11,'06':12,'07':13,'08':14,'09':15,'10':16,'11':17,'12':18,'13':19,
				'14':20,'15':21,'16':22,'17':23,'18':24,'19':25,'20':26,'21':27,'22':28,'23':29,'24':30,'25':31,'26':32,
				'27':33,'28':34,'29':35,'30':36,'31':37,'32':38,'33':39,'34':40,'35':41,'36':42,'37':43,'38':44,'39':45,
				'40':46,'41':47,'42':48,'43':49,'44':50,'45':51,'46':52,'47':53,'48':54,'49':55,'50':56,'51':57,'52':58,
				'53':59,'54':60,'55':61,'56':62,'57':63,'58':64,'59':65,'60':66,'61':67,'62':68}
	for filename in sorted(glob.glob('/Users/Ileam Montoya/Documents/Analisis KPI/Dashboard/Dashboard_Fija_W17.xlsx')):
		print(filename)
		week_text=re.search('W(.+).xlsx',filename)
		wb=openpyxl.load_workbook(filename)
		for tab in tabs:

			data_sheet=data_ret[tab]
			columns=()
			if tab in ['DSLAMS > 80','Hubspoke > 80%','Interfaces Recurrentes > 95','Interfaces Recurrentes >70< 95']:
				columns=(3,4,7)
			else:
				columns=(3,4,6)
		
			sheet=wb[tab]
			row, column = 3, 1
			eq_int=OrderedDict()
			salir=sheet.cell(row=row,column=column).value
			while salir != None:
				semana = week_text.group(1)
				utilizacion = sheet.cell(row=row,column=columns[2]).value
				equipo = sheet.cell(row=row,column=columns[0]).value
				interface = sheet.cell(row=row,column=columns[1]).value
				pais = sheet.cell(row=row,column=2).value
				if pais:
					pais = pais.strip()
				ip = sheet.cell(row=row,column=5).value
				interface_text=f'{equipo}{interface}'
				if interface_text not in eq_int:
					eq_int[interface_text]=[(semana,utilizacion,equipo,interface,pais,ip)]
				else:
					eq_int[interface_text].append((week_text.group(1),sheet.cell(row=row,column=columns[2]).value))
				row+=1
				salir=sheet.cell(row=row,column=1).value
		

			#Activar si solo se esta corriendo la ultima semana
			equipo_row=OrderedDict()
			data_row, data_column= 3,5
			salir=data_sheet.cell(row=data_row,column=data_column).value
			while salir != None:
				equipo_row[salir]=data_row
				data_row+=1
				salir=data_sheet.cell(row=data_row,column=data_column).value
			
			print('NEW WEEK', tab)
			for key in eq_int:
				print(key,eq_int[key])


			data_row, data_column= 1, 1
			for key in eq_int:
				if key in equipo_row:
					data_row=equipo_row[key]
				else:
					data_row=len(equipo_row)+3
					equipo_row[key]=data_row

					print(key," NEW")
				for data_point in eq_int[key]:
					data_column=week_dicts[str(data_point[0])]
					if data_sheet.cell(row=data_row,column=data_column).value == None:
						data_sheet.cell(row=data_row,column=1).value=data_point[4]
						data_sheet.cell(row=data_row,column=2).value=data_point[5]
						data_sheet.cell(row=data_row,column=3).value=data_point[2]
						data_sheet.cell(row=data_row,column=4).value=data_point[3]
						data_sheet.cell(row=data_row,column=5).value=key
						data_sheet.cell(row=data_row,column=data_column).value=data_point[1]					
					else:
						data_sheet.cell(row=data_row,column=data_column).value=data_point[1]
			

	data_ret.save('/Users/Ileam Montoya/Dropbox/Claro/Optimization - Daily/Data retrieval/dashboard.xlsx')

if __name__ == "__main__":

	walk_it()