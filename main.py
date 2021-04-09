#!/usr/bin/python
# -*- coding: latin-1 -*-
"""
Created on 08/03/2021

@author: Jonathan Servia Mandome
"""
from parts.word.elements.pagenumber import PageNumber


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
	t = header.add_table([['1', PageNumber(header, 'Página ')]],
							column_width=[5000, 5000], horizontal_alignment=['c', 'r'])

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
	body.add_paragraph('drtwnmw´nsrmdrm')
	for i in range(206):
		body.add_paragraph('')
	p = body.add_paragraph('drmdrm')
	p.get_properties().set_keep_next(True)
	p.get_properties().set_pstyle('Descripcin')

	tabla_1 = body.add_table(lineas_1, column_width=cw, borders={'all': {'sz': 4}})
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
	doc_word = CertificadoSAT(partes, args)


