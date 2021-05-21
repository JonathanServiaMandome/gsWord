# coding=utf-8
import document
import shape
from db import error, Trae_Fila, lee, rg_vacio, gpx, cl, copia_rg, FDC, GS_INS, u_libre, Busca_Prox, Fecha, i_selec, \
	Abre_Aplicacion, Abre_Empresa
from db import Int, Num, lista, Num_aFecha, lee_dc
from aa_funciones import Serie, GetColumnasACC_CK, LineaToDc
from shape import TextBox


def CertificadoSAT(partes, args, columnas_checks=GetColumnasACC_CK(), LineaToDc=LineaToDc, Serie=Serie, cl=cl, gpx=gpx):
	#from gsWord import document, shape
	import os

	u_libres = {}

	def Posicion(campo, arp, cl=cl, gpx=gpx):
		fila = Trae_Fila(FDC[gpx[1]][arp][2], campo, clr=-2)
		fmt = Trae_Fila(FDC[gpx[1]][arp][2], campo, clr=2)
		if fila is None:
			error(cl, "No existe el campo '%s' en el diccionario '%s'." % (campo, arp))
		return fila, fmt

	def merge_vars(text, dc, cl=cl):
		for key in dc.keys():
			if key not in text:
				continue
			if type(dc[key]) == list:
				text = text.split(' ')
				for k in range(len(text)):
					if text[k].startswith(key):
						text[k] = text[k].split('/')
						if len(text[k]) == 1:
							text[k] = str(dc[key])
						elif len(text[k]) == 2:
							v = ''
							if Int(text[k][1]) < len(dc[key]):
								v = str(dc[key][Int(text[k][1])])
							text[k] = v
							text[k] = str(dc[key][Int(text[k][1])])
						elif len(text[k]) == 3:
							v = ''
							if Int(text[k][1]) < len(dc[key]) and Int(text[k][2]) < len(dc[key][Int(text[k][1])]):
								v = str(dc[key][Int(text[k][1])][Int(text[k][2])])
							text[k] = v

				text = ' '.join(text)
			else:
				val = dc[key]
				if val is None:
					val = ''
				if hasattr(dc[key], 'keys') and 'IDX' in dc[key].keys():
					val = dc[key]['IDX']
				text = text.replace(key, val)

		return text

	##

	def GetAcciones(cd_parte, parte, idioma, columnas_checks=columnas_checks, _LineaToDc=LineaToDc, cl=cl, gpx=gpx):
		acciones_parte = i_selec(cl, gpx, 'acciones', 'ACC_NPAR', cd_parte, cd_parte)
		_checks = {}
		_contratos_checks = {}
		_tecnicos = []
		dc_tecnicos = {}
		_observaciones = {}

		_tipo_elementos = []
		arts = {}

		periodicidad = ''
		first_action = {}
		tipo_parte = parte['PT_TIPO']['IDX']
		acciones_parte=[]
		c='016232'
		while c < '016255':
			if c=='016250':
				c='016251'
				continue
			acciones_parte.append(c)
			c=Busca_Prox(c)
		for cd_accion in acciones_parte:
			rg_accion = lee_dc(lee_dc, gpx, 'acciones', cd_accion)

			if not periodicidad:
				periodicidad = rg_accion['ACC_TPER']
			if rg_accion['ACC_TPER'] == 'M':
				periodicidad = rg_accion['ACC_TPER']
			cdar = rg_accion['ACC_CDAR']['IDX']
			if cdar not in arts.keys():
				arts[cdar] = lee_dc(lee_dc, gpx, 'articulos', cdar, rels='')
			articulo = arts[cdar]

			if articulo['AR_TIPO'] != 'CC':
				deno_n1 = ''
				deno_n2 = ''
				for ln in FDC[gpx[1]]['articulos'][2]:
					campo, descripcion, tipo, cols, relaciones = ln[:5]
					if campo == idioma['PTCI_N1EI']:
						if relaciones:
							r = lee_dc(lee_dc, gpx, relaciones, articulo[campo])
							deno_n1 = r.get(FDC[gpx[1]][relaciones][2][0][0], '')
						else:
							deno_n1 = ''

					if campo == idioma['PTCI_N2EI']:
						if relaciones:
							r = lee_dc(lee_dc, gpx, relaciones, articulo[campo])
							try:
								deno_n2 = r.get(FDC[gpx[1]][relaciones][2][0][0], '')
							except:
								error(cl, (FDC[gpx[1]][relaciones][2][0][0], r))
						else:
							deno_n2 = ''

				if deno_n1:
					fila_n1 = Trae_Fila(_tipo_elementos, deno_n1, clb=0, clr=-2)
					if fila_n1 is None:
						fila_n1 = len(_tipo_elementos)
						_tipo_elementos.append([deno_n1, []])

					if deno_n2 and deno_n2 not in _tipo_elementos[fila_n1][1]:
						_tipo_elementos[fila_n1][1].append(deno_n2)

			tx_ = ''
			if rg_accion['ACC_NSER']:
				tx_ = idioma['PTCI_NS'] + ' ' + rg_accion['ACC_NSER'] + ' - '
			id_ns_ = cdar + rg_accion['ACC_MAR']['IDX'] + rg_accion['ACC_NSER']
			cd_numero_serie = '%s%s (%s)' % (tx_, articulo['AR_DENO'], id_ns_)
			normativa = ''
			# Se acumula por normativa
			if articulo['AR_TIPO'] == 'CC':
				# Si no hay check del elemento se busca el check del contrato
				cd_check = rg_accion['ACC_CKGC']
				if cd_check:
					rg_check = lee_dc(lee_dc, gpx, 'contratos_checklists', cd_check)
					if rg_check == 1:
						error(cl, "No existe el checklist del contrato '%s'." % cd_check)
					normativa = rg_check['CCK_NOR']['IDX']
			else:
				cd_check = rg_accion['ACC_CKG']['IDX']
				# Si hay check del elemento
				if cd_check:
					rg_check = lee_dc(lee_dc, gpx, 'checklists', cd_check)
					if rg_check == 1:
						error(cl, "No existe el checklist '%s'." % cd_check)
					normativa = rg_check['CK_NOR']['IDX']

			if normativa:
				rg_normativa = lee_dc(lee_dc, gpx, 'normativa', normativa)
				if rg_normativa == 1:
					error(cl, "No existe la normativa '%s'." % normativa)
			else:
				rg_normativa = {}

			_texto_antes, _texto_desupes, _tabla_resumen = '', '', ''
			oficial = 'S'
			for _linea in rg_normativa.get('NOR_CERT', []):

				if _linea[0] != tipo_parte or _linea[1] != oficial:
					continue
				_texto_antes, _texto_desupes, _tabla_resumen = _linea[2:5]

			'''_clave = (rg_normativa.get('NOR_NOR', 0), rg_normativa.get('NOR_CAP', 0), normativa,
			          rg_normativa.get('NOR_TITNOR', {}).get('NORC_DENO', ''),
			          rg_normativa.get('NOR_IMP', ''), _texto_antes, _texto_desupes)'''
			_clave = (rg_normativa.get('NOR_NOR', 0), rg_normativa.get('NOR_CAP', 0),)

			# print _clave
			if _clave not in _checks.keys():
				_checks[_clave] = {'BOLEANOS': [],
				                   'OTROS': [],
				                   'RESUMEN': {
					                   'horizontal': False,
					                   'titulos': [],
					                   'anchos': [],
					                   'h_align': [],
					                   'texto': [],
					                   'lineas': []
				                   },
				                   'DATA': [normativa,
				                            rg_normativa.get('NOR_TITNOR', {}).get('NORC_DENO', ''),
				                            rg_normativa.get('NOR_IMP', ''), _texto_antes, _texto_desupes]
				                   }

			if _tabla_resumen:
				td = eval(_tabla_resumen)
				if len(td) == 2:
					ls_tabla, horizontal = td
					texto_antes_tabla = ''
				elif len(td) == 3:
					ls_tabla, horizontal, texto_antes_tabla = td
				else:
					error(cl, "Nº de argumentos incorectos en la tabla de resumen")
				_checks[_clave]['RESUMEN']['horizontal'] = horizontal == 'S'
				_checks[_clave]['RESUMEN']['texto'] = texto_antes_tabla
				ln_tabla = []
				for ln_ in ls_tabla:
					col_name, diccionario, nombre_campo, columna, relacion, ancho, columna_filtrar, _idx_filtrar = ln_

					if diccionario == 'acciones':
						pos, fmt = Posicion(nombre_campo, diccionario)
						value = rg_accion[nombre_campo]
						if _clave not in first_action.keys():
							first_action[_clave] = True

						if first_action[_clave]:
							_checks[_clave]['RESUMEN']['anchos'].append(str(ancho) + '%')
							_checks[_clave]['RESUMEN']['titulos'].append(col_name)

						if columna_filtrar:
							fmt = fmt.split(' ')
							value = Trae_Fila(value, _idx_filtrar, clb=Int(columna_filtrar), clr=Int(columna))
							if value is None:
								value = ''
							fmt = fmt[Int(columna)]

						h_align = 'j'
						if fmt == 'd':
							h_align = 'c'
							value = Num_aFecha(value)
							if value is None:
								value = ''
						elif fmt in 'i012345':
							h_align = 'r'
						if first_action[_clave]:
							_checks[_clave]['RESUMEN']['h_align'].append(h_align)
						if type(value) == dict:
							value = value['IDX']
						ln_tabla.append(value)

				_checks[_clave]['RESUMEN']['lineas'].append(ln_tabla)

			if _clave not in _observaciones.keys():
				_observaciones[_clave] = []

			key = ''
			grupo = ''

			for k in range(len(rg_accion['ACC_CK'])):
				dc_check = _LineaToDc(rg_accion['ACC_CK'][k], columnas_checks)
				pregunta = dc_check['descripcion']
				respuesta = dc_check['respuesta']
				padre = dc_check['padre']
				numero = dc_check['numero']
				tipo = dc_check['tipo']
				if padre:
					fila_ns = None
					if grupo == 'BOLEANOS':
						fila_padre = Trae_Fila(_checks[_clave][grupo], padre, clb=0, clr=-2)
						if fila_padre is None:
							fila_padre = len(_checks[_clave][grupo])
							_checks[_clave][grupo].append(copia_rg(linea))
					else:

						fila_ns = Trae_Fila(_checks[_clave][grupo], cd_numero_serie, clb=0, clr=-2)
						if fila_ns is None:
							fila_ns = len(_checks[_clave][grupo])
							_checks[_clave][grupo].append([cd_numero_serie, []])
						fila_padre = Trae_Fila(_checks[_clave][grupo][fila_ns][1], padre, clb=0, clr=-2)

						if fila_padre is None:
							fila_padre = len(_checks[_clave][grupo][fila_ns][1])
							_checks[_clave][grupo][fila_ns][1].append(copia_rg(linea))
					# error(cl, (linea, fila_padre, checks[clave][grupo][fila_padre]))

					'''if grupo == 'BOLEANOS':
						_checks[_clave][grupo][fila_padre][1] += '  ' + pregunta + ' ' + respuesta
					else:
						a = _checks[_clave][grupo]
						_checks[_clave][grupo][fila_ns][1][fila_padre][1] += '\n' + pregunta + ' ' + respuesta'''
					if grupo == 'BOLEANOS':
						_checks[_clave][grupo][fila_padre][1] += '\n' + pregunta + ' ' + respuesta
					else:
						a = _checks[_clave][grupo]
						_checks[_clave][grupo][fila_ns][1][fila_padre][1] += '\n' + pregunta + ' ' + respuesta

					continue

				grupo = 'OTROS'
				linea = [numero, pregunta, respuesta]
				if tipo in 'XB':
					grupo = 'BOLEANOS'
					linea = [numero, pregunta, 0, 0, 0]

				if grupo == 'BOLEANOS':
					fila = Trae_Fila(_checks[_clave][grupo], numero, clb=0, clr=-2)
					if fila is None:
						fila = len(_checks[_clave][grupo])
						_checks[_clave][grupo].append(linea)

					if respuesta == 'SI':
						_checks[_clave][grupo][fila][2] += 1
					elif respuesta == 'N/A':
						_checks[_clave][grupo][fila][3] += 1
					elif respuesta == 'NO':
						_checks[_clave][grupo][fila][4] += 1
				else:
					fila = Trae_Fila(_checks[_clave][grupo], cd_numero_serie, clb=0, clr=-2)
					if fila is None:
						fila = len(_checks[_clave][grupo])
						_checks[_clave][grupo].append([cd_numero_serie, []])

					_checks[_clave][grupo][fila][1].append(linea)

			if rg_accion['ACC_OBS']:
				if articulo['AR_TIPO'] != 'CC':
					_observaciones[_clave].append(cd_numero_serie)
				_observaciones[_clave].append(rg_accion['ACC_OBS'])

			for ln_tec in rg_accion['ACC_TNCS']:
				cd_tecnico = ln_tec[0]
				estado = ln_tec[1]

				dc_tecnicos[cd_tecnico] = dc_tecnicos.get(cd_tecnico, lee_dc(lee_dc, gpx, 'personal', cd_tecnico))
				if dc_tecnicos[cd_tecnico] == 1:
					error(cl, "No existe el técnico '%s'." % cd_tecnico)
				rgt = dc_tecnicos[cd_tecnico]

				if rgt['PE_DENO'] not in _tecnicos:
					_tecnicos.append(rgt['PE_DENO'])

			first_action[_clave] = False
		return [_checks, _contratos_checks, _tecnicos, _observaciones, _tipo_elementos, periodicidad]

	##

	def LeeParte(cd_parte, cl=cl, gpx=gpx):
		dc = {}
		parte = lee_dc(lee_dc, gpx, 'partes', cd_parte)
		cliente = lee_dc(lee_dc, gpx, 'clientes', parte['PT_CCL']['IDX'])
		if cliente == 1:
			error(cl, "No existe el cliente '%s'." % parte['PT_CCL']['IDX'])
		parte.update(cliente)

		delegacion = lee_dc(lee_dc, gpx, 'delegaciones', parte['PT_CCL']['IDX'] + parte['PT_DEL']['IDX'])
		if delegacion == 1:
			error(cl,
			      "No existe la delegación '%s' del cliente '%s'." % (parte['PT_DEL']['IDX'], parte['PT_CCL']['IDX']))
		parte.update(delegacion)

		contrato = lee_dc(lee_dc, gpx, 'contratos', parte['PT_CON']['IDX'])
		ls_partes = i_selec(cl, gpx, 'partes', 'PT_CON', parte['PT_CON']['IDX'], parte['PT_CON']['IDX'])
		parte.update(contrato)
		parte['CN_FECU'] = ''
		if ls_partes:
			parte_ = lee_dc(cl, gpx, 'partes', ls_partes[-1], rels='n')
			parte['FECU'] = Num_aFecha(parte_['PT_FECPA'])
		parte['CN_FEC'] = Num_aFecha(parte['CN_FEC'])
		tipo_contrato = lee_dc(lee_dc, gpx, 'contratos_clases', parte['PT_CON']['CN_TIPO'])
		parte.update(tipo_contrato)
		parte['PT_DIREC/1/2'] = ''
		if len(parte['PT_DIREC']) > 1:
			parte['PT_DIREC/1/2'] = parte['PT_DIREC'][1][2]
		fecr = Num_aFecha(parte['PT_FECR'])
		if fecr is None:
			fecr = ''
		else:
			fecr = fecr[3:]
		parte['PT_FECR'] = fecr
		return [dc, parte]

	##

	def FormateaTablaPrincipal(idioma, tabla, color_primario, color_secundario, header=True):

		tabla.set_spacing({'after': '60', 'before': '60'})
		tabla.get_properties().set_cell_margin(
			{'start': {'w': '60'}, 'end': {'w': '60'}, 'top': {'w': '30'}, 'bottom': {'w': '30'}})

		tabla.set_font_size(idioma['PTCI_FSIZE'])
		if header:
			tabla.get_row(0).set_background_colour(color_primario)
			tabla.get_row(0).set_foreground_colour(color_secundario)
			tabla.get_row(0).set_font_format('b')
		for row in tabla.get_rows():
			row.get_cell(0).set_font_format('b')
			for cell in row.get_cells():
				cell.get_properties().set_vertical_alignment('center')
				for p in cell.elements:
					for t in p.elements:
						t.set_font(idioma['PTCI_FONT'])

	##

	def FormateaTitulo(idioma, paragraph):

		paragraph.set_font_format('b')
		paragraph.get_properties().set_keep_next(True)
		paragraph.set_font_size(idioma['PTCI_FSIZE'] + 4)
		paragraph.set_spacing({'before': '180', 'after': '80'})
		for t in paragraph.elements:
			t.set_font(idioma['PTCI_FONT'])

	##

	def FormateaCapitulo(idioma, paragraph):
		paragraph.set_font_format('b')
		paragraph.set_font_size(idioma['PTCI_FSIZE'] + 1)
		paragraph.set_spacing({'before': '180', 'after': '80'})
		for t in paragraph.elements:
			t.set_font(idioma['PTCI_FONT'])

	##

	def FormateaLista(idioma, paragraph):
		paragraph.set_font_size(idioma['PTCI_FSIZE'])
		for t in paragraph.elements:
			t.set_font(idioma['PTCI_FONT'])

	##

	def FormateaElemento(idioma, paragraph):
		paragraph.set_font_format('b')
		paragraph.set_font_size(idioma['PTCI_FSIZE'])
		paragraph.set_spacing({'before': '120', 'after': '120'})
		for t in paragraph.elements:
			t.set_font(idioma['PTCI_FONT'])

	##

	def Footer(doc_word, idioma, parte, _path_s):
		footer = doc_word.get_DefaultFooter()

		#-footer.add_rtf(idioma['PTCI_LOPD'])

	##

	def Header(header, doc_word, idioma, _path_i, _path_d, merge_vars=merge_vars, horizontal=False):
		num_pag = document.new_page_number(header, idioma['PTCI_SUFPAG'])
		num_pag.set_text_separator(idioma['PTCI_SEPPAG'])
		'''p = doc_word.new_paragraph(header)
		p.set_horizontal_alignment('r')'''

		paragraph_i = doc_word.new_paragraph(header, None)
		paragraph_i.AddPicture(header, _path_i, 1800, 1200)
		if _path_d:
			paragraph_d = doc_word.new_paragraph(header, None)
			paragraph_d.AddPicture(header, _path_d, 1800, 1200)

			table = header.add_table(
				[
					[paragraph_i, None, paragraph_d],
					[None, num_pag]
				],
				column_width=[2000, 6000, 2000],
				horizontal_alignment=['l', 'l', 'r'],
				# borders={'all': {'sz': 4}}
			)

			table.get_row(1).get_cell(0).get_properties().set_grid_span(2)
		else:
			table = header.add_table(
				[
					[paragraph_i, None],
					[None, num_pag]
				],
				column_width=[2000, 8000],
				horizontal_alignment=['l', 'r'],
				# borders={'all': {'sz': 4}}
			)

		font_size = idioma['PTCI_FSIZE']
		heigth_line = font_size * 20 + 40
		for ln_data in idioma['PTCI_TITULO']:
			paragraph = table.get_row(1).get_cell(0).add_paragraph(merge_vars(ln_data, parte), font_format='b')
			paragraph.set_spacing({'after': '60', 'before': '60', 'line': heigth_line})
			paragraph.set_font_size(font_size + 2)
			for t in paragraph.elements:
				t.set_font(idioma['PTCI_FONT'])

		table.get_row(1).get_cell(0).set_font_format('b')
		table.get_row(1).get_cell(0).set_font_size(font_size + 4)
		table.get_row(1).get_cell(0).get_properties().set_table_cell_width(20000)
		table.get_row(1).get_cell(1).add_paragraph(
			merge_vars('Nº Certificado ' + parte['PT_NUMCERT'], parte),
			horizontal_alignment='r',
			font_size=font_size + 1,
			font_format='b'
		)
		ls = []
		for line in idioma['PTCI_HEAD']:
			ha = 'l'
			if not _path_d:
				ha = 'r'
			paragraph = table.get_row(0).get_cell(1).add_paragraph(line[0], horizontal_alignment=ha, font_format='b')
			paragraph.set_spacing({'after': '60', 'before': '60', 'line': heigth_line})
			paragraph.set_font_size(font_size)
			for t in paragraph.elements:
				t.set_font(idioma['PTCI_FONT'])

		pos = (-5013, 8984)
		size = (10946, 428)
		if horizontal:
			pos = (-2700, 6120)
			size = (6800, 650)
		reg_mercantil = shape.TextBox(
			header,
			text=idioma['PTCI_REGM'],
			size=size,
			rotation=270,
			r_position=[
				{'orientation': 'horizontal', 'position': pos[0], 'relative': 'page'},
				{'orientation': 'vertical', 'position': pos[1], 'relative': 'page'}
			],
			background_color='FFFFFF',
			horizontal_alignment='c',
			font_format='b',
			font_size=8
		)

		p = header.add_paragraph(reg_mercantil)
		p.set_spacing({'after': '0', 'before': '0', 'line': '10'})

	##

	def firma_sello(doc_word, body, _path_s, parte, idioma):
		ls = []
		font_size = idioma['PTCI_FSIZE']
		body.add_paragraph('')
		tec_firma = parte['PT_TCF']['IDX']
		deno_tec = parte['PT_TCF'].get('PE_DENO', '')
		deno_tec += ' ' + parte['PT_TCF'].get('PE_DNI', '')
		deno_tec += ' ' + parte['PT_TCF'].get('PE_NCOL', '')
		sello = doc_word.new_image(body, _path_s, 1500, 1000, anchor='inline', horizontal_alignment='c')
		d_sello = ''  # doc_word.new_image(body, _path_s, 1500, 1000, anchor='inline', horizontal_alignment='c')

		representante = parte['PT_REPR']
		table_sellos = doc_word.new_table(
			body,
			[
				[idioma['PTCI_TITSEL'], '', idioma['PTCI_TITCLI']],
				[sello, '', d_sello],
				[idioma['PTCI_PIESEL'] + deno_tec, '', idioma['PTCI_PIECLI'] + representante]
			],
			horizontal_alignment=['c', 'c', 'c'],
			column_width=[3000, 4000, 3000]
		)
		table_sellos.set_font_format({(0,): 'b', (2,): 'b'})
		table_sellos.get_row(0).set_font_size(font_size)
		table_sellos.get_row(2).set_font_size(font_size)

		table_bloc = body.add_table(
			[[table_sellos]],
			horizontal_alignment=['c'],
			column_width=[10000]
		)
		table_bloc.get_row(0).get_properties().cantSplit = True
		# TODO cargar firma técnico
		img_f = 1  # lee(cl,gpx,'imagenes-t',tec_firma)
		if img_f not in [1, '']:
			_path_f = path_temp + tec_firma
			open(_path_f, 'wb').write(img_f)

	##

	def Portada(body, idioma, parte, numero_norma, numero_capitulo, section_margins, FormateaTitulo=FormateaTitulo,
	            merge_vars=merge_vars,
	            FormateaTablaPrincipal=FormateaTablaPrincipal):

		# Texto inicial
		if idioma['PTCI_TXTINI']:
			#-body.add_rtf(merge_vars(idioma['PTCI_TXTINI'], parte))
			numero_norma += 1

		# Apartado 1
		titulo_elemento = '%d. %s' % (numero_norma, merge_vars(idioma['PTCI_TITDP'], parte))
		numero_norma += 1
		paragraph_1 = body.add_paragraph(titulo_elemento)
		FormateaTitulo(idioma, paragraph_1)
		grid_span = []
		lineas_1 = []
		e = []
		for i in range(len(idioma['PTCI_DP'])):
			etiqueta, variable = idioma['PTCI_DP'][i][:2]
			e.append([i, variable])
			if variable:
				lineas_1.append([merge_vars(etiqueta, parte), merge_vars(variable, parte)])
			else:
				grid_span.append(i)
				lineas_1.append([merge_vars(etiqueta, parte)])
		tabla_1 = body.add_table(lineas_1, column_width=['25%', '75%'], borders=borders)

		for i in grid_span:
			tabla_1.get_row(i).get_cell(0).get_properties().set_grid_span(2)
			tabla_1.get_row(i).get_cell(0).get_properties().set_table_cell_width('100%')
		FormateaTablaPrincipal(idioma, tabla_1, color_primario, color_secundario, False)

		# Apartado 2
		titulo_elemento = '%d. %s' % (numero_norma, merge_vars(idioma['PTCI_TITDI'], parte))
		numero_norma += 1
		paragraph_1 = body.add_paragraph(titulo_elemento)
		FormateaTitulo(idioma, paragraph_1)
		grid_span = []
		lineas_1 = []
		for i in range(len(idioma['PTCI_DI'])):
			etiqueta, variable = idioma['PTCI_DI'][i][:2]
			if variable:
				lineas_1.append([merge_vars(etiqueta, parte), merge_vars(variable, parte)])
			else:
				grid_span.append(i)
				lineas_1.append([merge_vars(etiqueta, parte)])
		tabla_1 = body.add_table(lineas_1, column_width=['25%', '75%'], borders=borders)
		for i in grid_span:
			tabla_1.get_row(i).get_cell(0).get_properties().set_grid_span(2)
			tabla_1.get_row(i).get_cell(0).get_properties().set_table_cell_width('100%')
		FormateaTablaPrincipal(idioma, tabla_1, color_primario, color_secundario, False)

		_section = body.add_section()
		_section.SetMargins(section_margins)
		# Apartado 3
		titulo_elementos = '%d. %s' % (numero_norma, merge_vars(idioma['PTCI_TITEI'], parte))
		FormateaTitulo(idioma, body.add_paragraph(titulo_elementos))
		numero_norma += 1

		if idioma['PTCI_TXTEI']:
			#-body.add_rtf(idioma['PTCI_TXTEI'])
			pass

		t = 1
		for ln in tipo_elementos:
			s = 1
			deno_grupo, familias = ln[:3]
			r = ' '
			if r > 9:
				r = ''
			FormateaLista(idioma, body.add_paragraph('%d. %s%s' % (t, r, deno_grupo)))
			for deno_familia in familias:
				FormateaLista(idioma, body.add_paragraph('    %d.%d. %s' % (t, s, deno_familia)))
				s += 1
			t += 1

		# Textos contrato
		for cd_text in parte['PT_CON']['CN_TXT']:
			body.add_section().SetMargins(section_margins)
			_text = lee_dc(lee_dc, gpx, 'fcartas', cd_text)

			_text['FA_TXT'] = merge_vars(_text['FA_TXT'], parte)
			FormateaTitulo(idioma, body.add_paragraph(str(numero_norma) + '. ' + _text['FA_DENO']))
			#print '88888888888888888'
			body.add_rtf(_text['FA_TXT'])
			numero_norma += 1

		lista_checks_c = checks.keys()
		lista_checks_c.sort()
		_nor = ''
		nn = n_capitulo = 1

		for ln in lista_checks_c:
			_norma, _capitulo = ln
			if _norma > 1:
				continue
			cd_normativa, nombre_normativa, nombre_capitulo, _texto_antes, _texto_despues = checks[ln]['DATA']
			if not _nor:
				body.add_section().SetMargins(section_margins)

			if not nombre_normativa:
				nombre_normativa = 'DATOS GENERALES'
			if not _nor or _nor != _norma:
				_nor = str(numero_norma) + '. ' + nombre_normativa
				nn = numero_norma
				numero_norma += 1
				n_capitulo = 1
				paragraph_elemento = body.add_paragraph(_nor)
				FormateaTitulo(idioma, paragraph_elemento)

			paragraph_elemento = body.add_paragraph('%d.%d. %s' % (nn, n_capitulo, nombre_capitulo))
			FormateaCapitulo(idioma, paragraph_elemento)

			datos_generales = copia_rg(checks[ln].get('BOLEANOS', []))
			otros_datos_generales = copia_rg(checks[ln].get('OTROS', []))
			del checks[ln]

			if _texto_antes:
				rg_txt = lee_dc(lee_dc, gpx, 'fcartas', _texto_antes)
				if rg_txt == 1:
					error(cl, "No existe el texto '%s'" % _texto_antes)
				txt_antes = rg_txt['FA_TXT']
				txt_antes = merge_vars(txt_antes, parte)
				#-body.add_rtf(txt_antes)
			if datos_generales:
				# #$>datos_generales
				titulo_datos_generales = idioma['PTCI_TBOO']
				for j in range(len(datos_generales)):
					datos_generales[j][0] = "%d.%d.%d" % (nn, n_capitulo, j + 1)
					if '\n' in datos_generales[j][1]:
						datos_generales[j][1] = datos_generales[j][1].split('\n')

				tabla_datos_generales = body.add_table(datos_generales,
				                                       titulo_datos_generales,
				                                       column_width=columnas_checks,
				                                       borders=borders_checks,
				                                       horizontal_alignment=['r', 'j', 'c', 'c', 'c']
				                                       )
				FormateaTablaPrincipal(idioma, tabla_datos_generales, color_primario, color_secundario)

			if otros_datos_generales:
				if datos_generales:
					body.add_paragraph('')

				for i in range(len(otros_datos_generales)):
					titulo, lineas = otros_datos_generales[i]

					for j in range(len(lineas)):
						lineas[j][0] = "%d.%d.%d" % (nn, n_capitulo, j + 1)
						if '\n' in lineas[j][1]:
							lineas[j][1] = lineas[j][1].split('\n')

					_pie = body.add_table(lineas, idioma['PTCI_TNOR'],
					                      horizontal_alignment=['r', 'l', 'l'],
					                      column_width=columnas_others,
					                      borders=borders_checks)
					FormateaTablaPrincipal(idioma, _pie, color_primario, color_secundario)
			n_capitulo += 1

			if _texto_despues:
				body.add_paragraph('')
				rg_txt = lee_dc(lee_dc, gpx, 'fcartas', _texto_despues)
				if rg_txt == 1:
					error(cl, "No existe el texto '%s'" % _texto_despues)
				txt_despues = rg_txt['FA_TXT']
				txt_despues = merge_vars(txt_despues, parte)
				#-body.add_rtf(txt_despues)
			else:
				body.add_paragraph('')

			if not dc_observaciones.get(ln, []):
				observaciones = ['']
			else:
				observaciones = [''] * len(dc_observaciones[ln])
			_tabla_obs = body.add_table(observaciones,
			                            [idioma['PTCI_OBS']],
			                            horizontal_alignment=['j'],
			                            column_width=['100%'],
			                            borders=borders_obs
			                            )
			for i in range(len(dc_observaciones.get(ln, []))):
				_tabla_obs.get_row(i + 1).get_cell(0).elements = []
				if dc_observaciones[ln][i].startswith('{\\rtf'):
					#-_tabla_obs.get_row(i + 1).get_cell(0).add_rtf(dc_observaciones[ln][i])
					pass
				else:
					_tabla_obs.get_row(i + 1).get_cell(0).add_paragraph(dc_observaciones[ln][i])
			FormateaTablaPrincipal(idioma, _tabla_obs, color_primario, color_secundario)

			if otros_datos_generales and datos_generales:
				firma_sello(doc_word, body, _path_s, parte, idioma)
		return numero_norma, numero_capitulo

	##

	idioma = lee_dc(lee_dc, gpx, 'partes_certificado_idioma', '000')
	path_temp = GS_INS + '/temp/'
	if not os.path.exists(path_temp):
		os.makedirs(path_temp)

	parametros = lee_dc(lee_dc, gpx, 'parametros', '0')

	# Se recuperan los colores para las cabeceras de la tabla
	color = lista(parametros['P_COLPRI'], ':', 3)
	rgb = (Int(color[0]), Int(color[1]), Int(color[2]))
	color_primario = '%02x%02x%02x' % rgb
	if not parametros['P_COLSEC']:
		parametros['P_COLSEC'] = '255:255:255'
	color = lista(parametros['P_COLSEC'], ':', 3)
	rgb = (Int(color[0]), Int(color[1]), Int(color[2]))
	color_secundario = '%02x%02x%02x' % rgb

	imd = idioma['PTCI_LOGOD']

	_path_d = ''
	if imd:
		img_d = lee(cl, gpx, 'imgs_v', imd)
		if img_d == 1:
			error(cl, "No existe el logo " + imd)
		_path_d = path_temp + imd

		open(_path_d, 'wb').write(img_d)

	imi = idioma['PTCI_LOGOI']

	img_i = lee(cl, gpx, 'imgs_v', imi)
	if img_i == 1:
		error(cl, "No existe el logo " + imi)
	_path_i = path_temp + imi
	open(_path_i, 'wb').write(img_i)

	im_sello = idioma['PTCI_IMGSEL']
	img_sello = lee(cl, gpx, 'imgs_v', im_sello)
	if img_i == 1:
		error(cl, "No existe el logo " + im_sello)
	_path_s = path_temp + im_sello
	open(_path_s, 'wb').write(img_sello)

	accion = args.get('accion', '')
	impresora = args.get('desti', '')

	borders_checks = {'all': {'sz': 4, 'color': color_primario, 'space': 0}}
	borders_obs = copia_rg(borders_checks)
	borders_obs['insideH'] = {'sz': 0, 'color': color_primario, 'space': 0}
	columnas_checks = [750, 7450, 600, 600, 600]
	columnas_others = [750, 7450, 1800]
	section_margins = {'left': 953, 'right': 953, 'top': 1000, 'header': 600, 'footer': 400}

	documentos_generados = []
	borders = {'all': {'sz': 4, 'color': color_primario, 'space': 0},
	           'insideV': {'value': 'nil'}}

	for cd_parte in partes:
		dc_parte, parte = LeeParte(cd_parte)
		parte.update(parametros)
		if parte['PT_NUMCERT'] and False:
			rgcertificado = lee_dc(lee_dc, gpx, 'partes_certificado', parte['PT_NUMCERT'])
			documentos_generados.append(rgcertificado['PTC_DOC'])
			continue

		serie_cer = parte['PT_SERC']['IDX']
		ser_cert_complete = Serie(serie_cer, parte['PT_FECR'])
		if ser_cert_complete not in u_libres.keys():
			u_libres[ser_cert_complete] = u_libre(gpx, 'partes_certificado', ser_cert_complete)
		idx_cert = u_libres[ser_cert_complete]
		parte['PT_NUMCERT'] = idx_cert
		r = GetAcciones(cd_parte, parte, idioma)
		checks, contratos_checks, tecnicos, dc_observaciones, tipo_elementos, periodicidad_revision = r
		if periodicidad_revision == 'M':
			parte['PT_PERIODICIDAD'] = 'TRIMESTRAL'
		else:
			parte['PT_PERIODICIDAD'] = 'ANUAL'

		# r = open('c:/users/jonathan/desktop/data_parte.txt', 'r').read()
		# dc_parte, parte, checks, tecnicos, observaciones = eval(r)
		checks_sin_normativa = []
		numero_norma = 1
		numero_capitulo = 1

		for clave in checks.keys():
			for i in range(len(checks[clave]['BOLEANOS'])):
				si, na, no = '', '', ''
				if checks[clave]['BOLEANOS'][i][4] > 0:
					si, na, no = '', '', 'X'
				elif checks[clave]['BOLEANOS'][i][3] > 0:
					si, na, no = '', 'X', ''
				elif checks[clave]['BOLEANOS'][i][2] > 0:
					si, na, no = 'X', '', ''
				checks[clave]['BOLEANOS'][i][2] = si
				checks[clave]['BOLEANOS'][i][3] = na
				checks[clave]['BOLEANOS'][i][4] = no

			if clave[0] == 0:
				try:
					checks_sin_normativa = copia_rg(checks[clave])
					del checks[clave]
				except:
					pass
				continue
		##

		# Se crea un documento en blanco
		f = cd_parte + '.docx'
		if True:
			f = cd_parte + Fecha('hms').replace(':', '') + '.docx'

		doc_word = document.Document(path_temp, f)
		doc_word.empty_document()
		# doc_word._debug = True

		# Body 1- Portada
		body = doc_word.get_body()
		section = body.get_active_section()
		# Modifico los márgenes de la página para hacerlo mas estrecho
		section.SetMargins(section_margins)
		# Header
		header = doc_word.get_DefaultHeader()
		Header(header, doc_word, idioma, _path_i, _path_d)
		# footer
		Footer(doc_word, idioma, parte, _path_s)

		if idioma['PTCI_TITULO1']:
			titulo = '%d. %s' % (numero_capitulo, merge_vars(idioma['PTCI_TITULO1'], parte))
			numero_capitulo += 1

		p_titulo = body.add_paragraph(titulo)
		FormateaTitulo(idioma, p_titulo)

		lista_checks = checks.keys()
		lista_checks.sort()
		# Se renumeran las preguntas
		for ln in lista_checks:
			if ln[0] == numero_norma and ln[1] == numero_capitulo:
				titulo_elemento = str(numero_capitulo) + '. ' + ln[3]
				dc_checks = copia_rg(checks[ln])

		numero_norma, numero_capitulo = Portada(body, idioma, parte, numero_norma, numero_capitulo, section_margins)

		# Body 2- Elementos

		change_header = False
		_nor = ''
		nn = 0
		cn = 0
		er = []
		for ln in lista_checks:
			# Cada grupo de elementos empezará en una nueva sección y en una nueva página
			n_norma, n_capitulo = ln
			if n_norma == 1:
				continue
			cd_check_, titulo_norma, titulo_elemento, texto_antes, texto_despues = checks[ln]['DATA']
			sect = body.add_section(margin_rigth=953, margin_left=953, margin_footer=400, orient='')
			if change_header:
				h = doc_word.new_header(sect)
				Header(h, doc_word, idioma, _path_i, _path_d)
				change_header = False

			if not cn or cn != n_norma:
				_nor = str(numero_norma) + '. ' + titulo_norma
				nn = numero_norma
				cn = n_norma
				numero_norma += 1
				n_capitulo = 1
				paragraph_elemento = body.add_paragraph(_nor)
				FormateaTitulo(idioma, paragraph_elemento)

			paragraph_1 = body.add_paragraph('%d.%d. %s' % (nn, n_capitulo, titulo_elemento))
			FormateaCapitulo(idioma, paragraph_1)

			if texto_antes:
				rg_txt = lee_dc(lee_dc, gpx, 'fcartas', texto_antes)
				if rg_txt == 1:
					error(cl, "No existe el texto '%s'" % texto_antes)
				txt_antes = rg_txt['FA_TXT']
				#-body.add_rtf(txt_antes)

			boolean_data = checks[ln].get('BOLEANOS', [])
			if boolean_data:
				for i in range(len(boolean_data)):
					boolean_data[i][0] = "%d.%d.%d" % (nn, numero_capitulo, i + 1)
					if '\n' in boolean_data[i][0]:
						boolean_data[i][0] = boolean_data[i][0].split('\n')
				# #$>boolean_data
				encabezado = body.add_table(boolean_data,
				                            idioma['PTCI_TBOO'],
				                            horizontal_alignment=['r', 'l', 'c', 'c', 'c'],
				                            column_width=columnas_checks,
				                            borders=borders_checks)
				FormateaTablaPrincipal(idioma, encabezado, color_primario, color_secundario)

			body.add_paragraph('')

			other_data = checks[ln].get('OTROS', [])
			if other_data:
				for i in range(len(other_data)):
					titulo, lineas = other_data[i]
					p_od = body.add_paragraph(titulo)
					FormateaElemento(idioma, p_od)
					p_od.get_properties().set_keep_next(True)

					for j in range(len(lineas)):
						lineas[j][0] = "%d.%d.%d" % (nn, numero_capitulo, j + 1)
						if '\n' in lineas[j][0]:
							lineas[j][0] = lineas[j][0].split('\n')
					pie = body.add_table(lineas,
					                     idioma['PTCI_TNOR'],
					                     horizontal_alignment=['r', 'l', 'l'],
					                     column_width=columnas_others,
					                     borders=borders_checks)
					FormateaTablaPrincipal(idioma, pie, color_primario, color_secundario)

			n_capitulo += 1

			body.add_paragraph('')
			numero_capitulo += 1

			if not dc_observaciones.get(ln, []):
				observaciones = ['']
			else:
				observaciones = [''] * len(dc_observaciones[ln])
			tabla_obs = body.add_table(
				observaciones,
				[idioma['PTCI_OBS']],
				horizontal_alignment=['j'],
				column_width=['100%'],
				borders=borders_obs
			)
			for i in range(len(dc_observaciones.get(ln, []))):

				# tabla_obs.get_row(i+1).get_cell(0).elements=[]
				if dc_observaciones[ln][i].startswith('{\\rtf'):
					pass
					#-tabla_obs.get_row(i + 1).get_cell(0).add_rtf(dc_observaciones[ln][i])
				else:
					tabla_obs.get_row(i + 1).get_cell(0).add_paragraph(dc_observaciones[ln][i])

			FormateaTablaPrincipal(idioma, tabla_obs, color_primario, color_secundario)

			if texto_despues:
				body.add_paragraph('')
				rg_txt = lee_dc(lee_dc, gpx, 'fcartas', texto_despues)
				if rg_txt == 1:
					error(cl, "No existe el texto '%s'" % texto_despues)
				txt_despues = rg_txt['FA_TXT']
				#-body.add_rtf(txt_despues)

			#firma_sello(doc_word, body, _path_s, parte, idioma)

			tabla_resumen = checks[ln].get('RESUMEN', {})
			if tabla_resumen.get('lineas', []):
				imprimir_horizontal = tabla_resumen['horizontal']
				orient = ''
				ml = 953
				if imprimir_horizontal:
					orient = 'landscape'
					ml = 1253
				body.add_section(margin_rigth=953, margin_left=ml, margin_footer=400, orient=orient)
				new_section = body.get_active_section()
				if imprimir_horizontal:
					h = doc_word.new_header(new_section)
					Header(h, doc_word, idioma, _path_i, _path_d, horizontal=True)
					change_header = True

				v = new_section.get_width() - new_section.get_margin_rigth() - new_section.get_margin_left()

				if tabla_resumen.get('texto', ''):
					#-body.add_rtf(tabla_resumen['texto'])
					pass
				titulo_tabla = tabla_resumen['titulos']
				anchos_tabla = tabla_resumen['anchos']
				for i in range(len(anchos_tabla)):
					anchos_tabla[i] = int(v * (Num(anchos_tabla[i].replace('%', '')) / 100.))

				h_align_tabla = tabla_resumen['h_align']
				lineas_tabla = tabla_resumen['lineas']

				tabla_res = body.add_table(lineas_tabla, titulo_tabla, column_width=anchos_tabla,
				                           borders=borders_checks)
				FormateaTablaPrincipal(idioma, tabla_res, color_primario, color_secundario)


		doc_word.set_variables({})
		doc_word.save()

		'''file_ = open(path_temp + f, 'rb').read()
		documentos_generados.append(file_)
		rg_certificado = rg_vacio(gpx[1], 'partes_certificado')
		rg_certificado < PTC_NPAR > = cd_parte
		rg_certificado < PTC_DOC > = file_
		serie_cer = parte['PT_SERC']['IDX']
		ser_cert_complete = Serie(serie_cer, parte['PT_FECR'])
		if ser_cert_complete not in u_libres.keys():
			u_libres[ser_cert_complete] = u_libre(gpx, 'partes_certificado', ser_cert_complete)
		idx_cert = u_libres[ser_cert_complete]
		rgparte = lee_dc(None, gpx, 'partes', cd_parte, rels='')
		rgparte['PT_NUMCERT'] = idx_cert
		# p_actu(cl, gpx, 'partes', cd_parte, rgparte, a_graba='', log='')
		u_libres[ser_cert_complete] = Busca_Prox(u_libres[ser_cert_complete])'''
		# p_actu(cl, gpx, 'partes_certificado', idx_cert, rg_certificado, a_graba='', log='')
		#del file_

		os.system('start ' + path_temp + f)
	return documentos_generados


if __name__ == '__main__':
	Abre_Aplicacion('sat')
	Abre_Empresa(gpx[0], gpx[1], gpx[2])
	CertificadoSAT(['0000004555'], {})
