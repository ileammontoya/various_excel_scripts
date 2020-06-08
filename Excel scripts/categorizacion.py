import re, glob, os, openpyxl
from collections import OrderedDict
from copy import copy

def load_categorizacion():

	wb=openpyxl.load_workbook('/Users/Ileam Montoya/Dropbox/Claro/Optimization - Daily/Dashboards/Categorizacion/Base_Categorizacion.xlsx')
	sheet=wb['Unificacion']

	eq_int=OrderedDict()
	row, column = 2, 1	
	salir=sheet.cell(row=row,column=column).value
	while salir != None:

		eq_regex = re.compile(r'^([_A-Za-z0-9-]+)')
		equipo = eq_regex.search(sheet.cell(row=row,column=1).value).group()
		int_ip = sheet.cell(row=row,column=2).value
		jerarquia = sheet.cell(row=row,column=5).value
		servicio = sheet.cell(row=row,column=7).value
		propagacion = sheet.cell(row=row,column=8).value

		interface=(equipo,int_ip,jerarquia,servicio,propagacion)
		interface_text='{}{}'.format(interface[0],interface[1])

		eq_int.setdefault(interface_text,interface)
		row+=1
		salir=sheet.cell(row=row,column=1).value
	return eq_int

def cat_to_current_week(eq_int):

	tabs=[	('EQ LINK CMTS > 70',11), ('CMTS DHCP POOL >80', 11), ('IPPOOL > 90',9), ('DSLAMS > 80',11), ('MSAN > 80',11), ('CORE >80',13), ('REFLECTOR >80',12), ('MPLS P >80',14),
			('MPLS PE >80',13), ('Isla de App > 80%',8), ('Equipos > 80%',14), ('ServiceApp > 40',9), ('Enlaces > 40',10), ('Hubspoke > 80%',10), ('Interfaces Fotonico > 95',9),
			('Interfaces Recurrentes > 95',10), ('Interfaces Recurrentes >70< 95',10)]
	current_week = 'Dashboard_Fija_W16 - Guatemala - Honduras.xlsx'
	uni=openpyxl.load_workbook(current_week)

	for tab in tabs:
		print(tab)
		sheet=uni[tab[0]]
		
		if tab[0] == 'Equipos > 80%':
			row, column = 4, 3
		else:
			row, column = 3, 3

		old_border = copy(sheet.cell(row=row-1,column=tab[1]).border)
		old_fill = copy(sheet.cell(row=row-1,column=tab[1]).fill)

		for plus, text in [(2,'JERARQUIA')]:
			sheet.cell(row=row-1,column=tab[1]+plus).value = text
			sheet.cell(row=row-1,column=tab[1]+plus).border = old_border
			sheet.cell(row=row-1,column=tab[1]+plus).fill = old_fill

		salir=sheet.cell(row=row,column=column).value
		while salir != None:
			eq_regex = re.compile(r'^([_A-Za-z0-9-]+)')
			equipo = eq_regex.search(sheet.cell(row=row,column=3).value).group()
			int_ip = sheet.cell(row=row,column=4).value
			pais = sheet.cell(row=row,column=2).value
			paises = ['GT','HN','Guatemala','Honduras','GUATEMALA','HONDURAS']
			# paises = ['GT','HN','Guatemala','Honduras', 'Nicaragua', 'NI', 'El Salvador', 'ESV', 'Costa Rica', 'CR','Panama']
			interface=(equipo,int_ip)
			interface_text='{}{}'.format(interface[0],interface[1])
			if interface_text in eq_int and pais in paises:
				sheet.cell(row=row,column=tab[1]+2).value = eq_int[interface_text][2]
				# sheet.cell(row=row,column=tab[1]+3).value = eq_int[interface_text][3]
				# sheet.cell(row=row,column=tab[1]+4).value = eq_int[interface_text][4]
			row+=1
			salir=sheet.cell(row=row,column=3).value

	uni.save(current_week)


if __name__ == "__main__":

	eq_int = load_categorizacion()
	print('Lista de equipos/interfaces tiene {} items'.format(len(eq_int.keys())))
	cat_to_current_week(eq_int)