#!/usr/bin/python
# -*- coding: latin-1 -*-
"""
Created on 08/03/2021

@author: Jonathan Servia Mandome
"""
from parts.word.elements.pagenumber import PageNumber
from shape import AlternateContent, TextBox


def LineaToDc(linea, columnas):
	dc = {}
	for key, posicion in columnas.items():
		dc[key] = eval(repr(linea[posicion]))
	return dc


def CertificadoSAT(partes, args):
	import document
	import os
	# import win32print,win32api

	path_temp = 'C:/Users/Jonathan/Desktop/pyword/'
	f = 'rtf.docx'

	_doc_word = document.Document(path_temp, f)
	_doc_word._debug = True
	_doc_word.EmptyDocument()
	body = _doc_word.get_body()
	header = _doc_word.get_DefaultHeader()
	t = header.add_table([['1', document.new_page_number(header, 'Página ')]],
							column_width=[5000, 5000], horizontal_alignment=['c', 'r'])

	# x = AlternateContent(header, text='ccccccccc')
	r_position=[]
	r_position.append({'orientation':'horizontal', 'position':0, 'relative': 'column'})
	r_position.append({'orientation':'vertical', 'position':0, 'relative': 'page'})
	header.add_paragraph( TextBox(header, 'aaaaaaaaaaa', (500, 250), rotation=270, r_position=r_position,
									horizontal_alignment='c', font_format='b', font_size=12) )
	'''style = doc_word.get_Part('style').get_doc_defaluts().get_rpr_defaults()
	style.get_values()['sz'] = 18
	style.get_values()['szCs'] = 18
	doc_word.get_Part('style').get_doc_defaluts().set_rpr_defaults(style)'''

	# p.get_properties().set_pstyle('Principal')

	texto_antes = '002'
	txt_antes = open(path_temp + '%s.rtf' % texto_antes, 'r').read()
	lineas_1 = [
		['Nombre', 'deno_ccl'],
		['Dirección', 'dir_ccl'],
		['Población', 'pob_ccl'],
		['E-mail', 'email_ccl'],
		['REPRESENTANTE RESPONSABLE'],
		['Nombre y apellidos', 'contacto_ccl'],
		['Cargo', 'cargo_ccl']
	]

	# rt = open('c:/users/jonathan/desktop/002.rtf','r').read()
	cw = ['25%', '75%']


	def Formatea(paragraph):
		paragraph.set_spacing({'after': '60', 'before': '60', 'line': '160'})

	paragraph_d = _doc_word.Image(header, 'c:/users/jonathan/desktop/out/iso.jpg', 2000, 1300)
	paragraph_i = _doc_word.Image(header, 'c:/users/jonathan/desktop/out/logo.png', 1800, 1200)

	datos_intec = _doc_word.Table(header, column_width=['100%'])
	datos_intec.get_properties().set_cell_margin({'start': {'w': 50}})
	nombre_empresa = 'P_NME'
	par, num_row = datos_intec.add_paragraph(nombre_empresa, horizontal_alignment='l', font_format='b')
	Formatea(par)
	data = 'P_DOM'
	_cell_ = datos_intec.get_row(num_row).get_cell(0)
	par1 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par1)
	data = 'P_CDP' + ' ' + 'P_POB'
	par2 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par2)
	data = 'P_RI'
	par3 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par3)
	data = 'Tel: %s    e-mail:%s' % ('P_TEL', 'P_MAIL')
	par4 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par4)
	datos_intec.set_foreground_colour('885500')

	table = header.add_table([[paragraph_i, datos_intec, paragraph_d]], column_width=[2000, 6000, 2000],
								horizontal_alignment=['l', 'l', 'r'])
	table.get_row(0).get_cell(0).get_properties().set_vertical_alignment('center')

	def FormateaTitulo(paragraph):

		paragraph.set_font_format('b')

		paragraph.set_font_size(14)
		paragraph.set_spacing({'before': '180', 'after': '80'})

	##
	def Formatea(paragraph):
		paragraph.set_spacing({'after': '60', 'before': '60', 'line': '160'})
	datos_intec = _doc_word.Table(header, column_width=[6000])
	nombre_empresa = 'P_NME'
	par, num_row = datos_intec.add_paragraph(nombre_empresa, horizontal_alignment='l', font_format='b')
	Formatea(par)
	data = 'P_DOM'
	_cell_ = datos_intec.get_row(num_row).get_cell(0)
	par1 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par1)
	data = 'P_CDP' + ' P_POB'
	par2 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par2)
	data = 'P_RI'
	par3 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par3)
	data = 'Tel: %s    e-mail:%s' % ('P_TEL', 'P_MAIL')
	par4 = _cell_.add_paragraph(data, horizontal_alignment='l', font_format='b')
	Formatea(par4)

	for i in range(0):
		body.add_paragraph('')
	p = body.add_paragraph('drmdrm')
	p.get_properties().set_keep_next(True)
	p.get_properties().set_pstyle('Descripcin')

	tabla_1 = body.add_table(lineas_1, column_width=cw, borders={'all': {'sz': 4}})
	tabla_1.get_properties().set_cell_margin({'start': {'w': 50}})
	tabla_1.get_properties().set_table_caption('titulo')
	tabla_1.set_spacing({'after': '80', 'before': '80', 'line': '180'})
	tabla_1.get_row(4).get_cell(0).get_properties().set_grid_span(2)
	tabla_1.get_row(4).get_cell(0).get_properties().set_table_cell_width('100%')
	for row in tabla_1.get_rows():
		for cell in row.get_cells():
			cell.get_properties().set_vertical_alignment('center')
	# tabla_1.get_row(4).get_cell(0).add_rtf(rt)
	for row in tabla_1.get_rows():
		for cell in row.get_cells():
			cell.get_properties().set_vertical_alignment('center')

	'''t = body.add_table([[None, '2'], ['t1', 't2'], borders={'all': {'sz': 4}})
	t.get_row(1).get_cell(0).add_rtf(txt_antes)
	t.set_foreground_colour('AA2233')'''

	body.add_paragraph('')
	body.add_paragraph('3333').SetFormatList()
	body.add_paragraph('3333').SetFormatList()
	body.add_paragraph('3333').SetFormatList('2', '1')

	body.add_paragraph('999').SetFormatList('2')
	body.add_paragraph('999').SetFormatList()
	body.add_paragraph('3333').SetFormatList()

	_doc_word.Save()

	os.system('start ' + path_temp + f)
	return _doc_word


if __name__ == '__main__':
	import document
	import os

	partes = ['0026085']
	args = {}
	#_doc_word = CertificadoSAT(partes, args)


	path_temp = 'C:/Users/Jonathan/Desktop/pyword/'
	f = 'rtf.docx'
	_doc_word = document.Document(path_temp, f)
	_doc_word._debug = True
	_doc_word.EmptyDocument()
	body = _doc_word.get_body()
	header = _doc_word.get_DefaultHeader()
	r_position=[]
	r_position.append({'orientation': 'horizontal', 'position': -5000, 'relative': 'column'})
	r_position.append({'orientation': 'vertical', 'position': 6000, 'relative': 'paragraph'})
	header.add_paragraph(TextBox(header, 'aaaaaaaaaaa', (50, 5000), (8000, 500), rotation=270, r_position=r_position,
								 background_color='FFEFFF', horizontal_alignment='c', font_format='b', font_size=12))

	_doc_word.Save()

	os.system('start ' + path_temp + f)
