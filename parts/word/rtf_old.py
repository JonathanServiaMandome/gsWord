#!/usr/bin/python
# -*- coding: utf-8 -*-

import re


class RtfCell:
	def __init__(self, parent, width='2000'):
		self.parent = parent
		self.rtf = self.parent.parent.parent

		self.width = round((int(width) * 100.) / self.rtf.size[0], 2)
		self.paragraphs = list()
		self.horizontal_alignment = 'left'
		self.vertical_alignment = 'None'
		self.adjust = 15

	def set_width(self, value):
		self.width = round((int(value) * 100.) / float(self.rtf.size[0]), 2)
		#print value, '*', 100., '/', self.rtf.size[0], '=', self.width

	def get_width(self):
		return self.width

	def set_horizontal_alignment(self, value):
		self.horizontal_alignment = value

	def get_horizontal_alignment(self):
		return self.horizontal_alignment

	def set_vertical_alignment(self, value):
		self.vertical_alignment = value

	def get_vertical_alignment(self):
		return self.vertical_alignment

	def get_value(self, mode='plain', word_object=None):
		value = str()

		if mode in ['plain', 'list']:
			for paragraph in self.paragraphs:
				txt = paragraph.get_value(mode)
				while len(txt) < self.adjust:
					txt = ' ' + txt
				value += txt
		elif mode in ['html']:
			value = list()
			for paragraph in self.paragraphs:
				v = [paragraph.get_value(mode, cell=True)]

				if v:
					#print value, self.width
					v.insert(0, '<td width="%s%s">' % (self.width, '%'))
					v.append('</td>')
				value.extend(v)

			value = ''.join(value)
		elif mode == 'word':
			cell = word_object.add_cell([], horizontal_alignment='l')
			# cell = Table.Cell(word_object, len(word_object.cells) + 1, [])
			for paragraph in self.paragraphs:
				paragraph.get_value(mode, word_object=cell)
		return value


class RtfRow:
	def __init__(self, parent):
		self.cells = list()
		self.horizontal_gap = '60'
		self.left = '0'
		self.parent = parent

	def set_horizontal_gap(self, value):
		self.horizontal_gap = value

	def get_horizontal_gap(self):
		return self.horizontal_gap

	def set_left(self, value):
		self.left = value

	def get_left(self):
		return self.left

	def get_value(self, mode='plain', word_object=None):
		value = str()
		if mode in ['plain', 'list']:
			for cell in self.cells:
				value += cell.get_value(mode)

		elif mode in ['html']:
			value = list()
			for cell in self.cells:
				value.append(cell.get_value(mode))
			if value:
				value.insert(0, '<tr>')
				value.append('</tr>')
			value = ''.join(value)
		elif mode in ['word']:
			row = word_object.add_row()
			for cell in self.cells:
				cell.get_value(mode, row)
		return value

	def add_cell(self):
		self.cells.append(RtfCell(self))


class RtfTable:
	def __init__(self, parent):
		self.name = 'RtfTable'
		self.rows = list()
		self.parent = parent

	def get_value(self, mode='plain', word_object=None):
		value = str()
		if mode in ['plain', 'list']:
			for row in self.rows:
				value += row.get_value(mode)
				value += '\n'

		elif mode in ['html']:
			value = list()
			for row in self.rows:
				value.append(row.get_value(mode))
			if value:
				value.insert(0, '<table>')
				value.append('</table>')
			value = ''.join(value)
		elif mode == 'word':
			value = self.parent.parent.add_table()

			for row in self.rows:
				row.get_value(mode, value)

		return value

	def add_row(self):
		self.rows.append(RtfRow(self))


class RtfParagraph:
	def __init__(self, parent, text_indent=0, id_paragraph=0):
		self.name = 'RtfParagraph'
		self.parent = parent
		self.texts = list()
		self.text_indent = text_indent
		self.type_list = ''
		self.number = 1
		self.id = id_paragraph
		self.horizontal_alignement = None

	def __str__(self):
		res = []
		for text in self.texts:
			res.append(text.get_value('plain').encode('utf-8', errors='ignore'))
		return '\n'.join(res)

	def set_horizontal_alignment(self, alignement):
		self.horizontal_alignement = alignement

	def get_value(self, mode='plain', cell=False, word_object=None):
		value = str()
		if mode in ['plain', 'list']:
			li = 0
			if self.text_indent:
				li = int(self.text_indent) // 90
			for text in self.texts:
				if li:
					for i in range(li - 1):
						value += u' '
					if self.type_list == 'q':
						value += u'· '
					else:
						value += u'%d. ' % self.number
				v = text.get_value(mode)
				if v is not None:
					value += v

		elif mode == 'html':
			if cell:
				for text in self.texts:
					value += text.get_value(mode, cell)
			else:
				value = u'<p>'
				if self.text_indent:
					value = u'<p style:"text-indent: %dpx">' % int(self.text_indent)
				for text in self.texts:
					value += text.get_value(mode)
				value = u'\t\t\t' + value + u'\t\t\t</p>\n'
		elif mode == 'word':
			paragraph = word_object.add_paragraph(_text=(), horizontal_alignment='l', font_format='', font_size=None)
			if self.horizontal_alignement is not None:
				paragraph.set_horizontal_alignment(self.horizontal_alignement)

			for text in self.texts:
				text.get_value(mode, word_object=paragraph)
		return value


class RtfText:
	def __init__(self, parent, text='', bold=False, italic=False, underline=False, font_size=1., font_color=1, font=0):
		self.text = text
		self.bold = bold
		self.italic = italic
		self.underline = underline
		self.font_size = font_size
		self.font_color = font_color
		self.font = font
		self.parent = parent
		self.rtf = self.parent.parent

	def __str__(self):
		txt = unicode(self.text.encode('utf-8'), errors='ignore')
		if len(txt) > 27:
			txt = txt[:27] + u'...'

		return u'b: %s i:%s ul:%s f:%s fs:%s cf:%s text: %s' % (
			self.bold, self.italic, self.underline, self.font, self.font_size, self.font_color, txt
		)

	def get_value(self, mode='plain', cell=False, word_object=None):
		# print self
		if mode in ['plain', 'list']:
			return self.text
		elif mode == 'word':
			font_format = ''
			if self.bold:
				font_format += 'b'
			elif self.italic:
				font_format += 'i'
			elif self.underline:
				font_format += 'u'
			txt = self.text.encode('cp1252')
			text = word_object.add_text(txt, font_format, int(self.font_size) // 2)

			colour = ''
			if self.font_color is not None:
				c = self.rtf.colors.get(int(self.font_color), ['0', '0', '0'])
				for _c in range(len(c)):
					c[_c] = int(c[_c])
				colour = '%02x%02x%02x' % tuple(c)

			if colour:
				#print colour
				text.get_properties().set_color(colour)
		elif mode == 'html':
			value = [self.text.replace('\n', '</br>')]
			if value == u'</br>':
				return u''.join(value)

			if self.bold:
				value.insert(0, u'<b>')
				value.append(u'</b>')
			if self.italic:
				value.insert(0, u'<i>')
				value.append(u'</i>')
			if self.underline:
				value.insert(0, u'<u>')
				value.append(u'</u>')
			try:
				colour = '%02x%02x%02x' % self.rtf.colors.get(int(self.font_color), ('0', '0', '0'))
			except:
				colour = ''

			if not cell:
				label = u'<label style="'
				try:
					fs = float(self.font_size) // 20
				except:
					fs = 1
				label += u'font-size: %.1fem; ' % fs
				if colour:
					label += u'color:%s; ' % colour
				try:
					label += u"font-family:'%s'" % self.rtf.fonts.get(int(self.font), [])[1]
				except:
					pass
				label += u'">'

				value.insert(0, label)

				value.append(u'</label>')
			return u''.join(value)


class RtfList:
	def __init__(self, parent):
		self.name = u'RtfList'
		self.parent = parent
		self.paragraphs = list()
		self.mark = 0

	def get_value(self, mode='plain', word_object=None):
		level = 0
		value = ''
		if mode in ['plain', 'list']:
			for paragraph in self.paragraphs:
				value += paragraph.get_value(mode) + '\n'
		elif mode in ['word']:
			pass
		else:
			for paragraph in self.paragraphs:
				if paragraph.text_indent < level and value:
					value += u'\t</ul>\n'
				if paragraph.text_indent > level or not value:
					value += u'\t<ul>\n'

				value += u'\t\t<li>\n'
				value += u'\t\t' + paragraph.get_value(mode)
				value += u'\t\t</li>\n'

				level = paragraph.text_indent
			if value:
				value += u'\t</ul>\n'
		return value


class Rtf:
	def __init__(self, text='', rtf_path='', parent=None):
		self.elements = list()
		self.rtf = text
		self.parent = parent

		self.path = rtf_path
		self.fonts = dict()
		self.colors = dict()
		self.default_font_size = 20
		self.default_foreground = 3
		self.default_background = 1
		self.cs = 1
		self.size = [11908, 16833]
		self.margins = [1800, 1800, 1440, 1440]
		if self.path:
			self.open_rtf(rtf_path)
		elif self.rtf:
			self.rtf_to_object()

	def open_rtf(self, _path):
		with open(_path, 'r') as f:
			self.rtf = f.read()
		self.rtf_to_object()

	def get_elements(self):
		return self.elements

	def get_value(self, mode='plain'):
		value = ''
		for element in self.elements:
			#print self.parent
			txt = element.get_value(mode, word_object=self.parent)
			if mode != 'word' and txt is not None:
				value += txt
		if mode == 'list':
			value = value.split('\n')
		return value

	def rtf_to_object(self):

		def ajustar(i):
			i = str(i)
			while len(i) < 10:
				i = '0' + i
			return i

		pattern = re.compile(r"\\([a-z]{1,32})(-?\d{1,10})?[ ]?|\\'([0-9a-f]{2})|\\([^a-z])|([{}])|[\r\n]+|(.)", re.I)
		# control words which specify a "destinations".
		destinations = frozenset((
			'aftncn', 'aftnsep', 'aftnsepc', 'annotation', 'atnauthor', 'atndate', 'atnicn', 'atnid', 'atnparent',
			'atnref', 'atntime', 'atrfend', 'atrfstart', 'author', 'background', 'bkmkend', 'bkmkstart', 'blipuid',
			'buptim', 'category', 'colorschememapping', 'colortbl', 'comment', 'company', 'creatim', 'datafield',
			'datastore', 'defchp', 'defpap', 'do', 'doccomm', 'docvar', 'dptxbxtext', 'ebcend', 'ebcstart',
			'factoidname', 'falt', 'fchars', 'ffdeftext', 'ffentrymcr', 'ffexitmcr', 'ffformat', 'ffhelptext', 'ffl',
			'ffname', 'ffstattext', 'field', 'file', 'filetbl', 'fldinst', 'fldrslt', 'fldtype', 'fname', 'fontemb',
			'fontfile', 'fonttbl', 'footer', 'footerf', 'footerl', 'footerr', 'footnote', 'formfield', 'ftncn',
			'ftnsep', 'ftnsepc', 'g', 'generator', 'gridtbl', 'header', 'headerf', 'headerl', 'headerr', 'hl', 'hlfr',
			'hlinkbase', 'hlloc', 'hlsrc', 'hsv', 'htmltag', 'info', 'keycode', 'keywords', 'latentstyles', 'lchars',
			'levelnumbers', 'leveltext', 'lfolevel', 'linkval', 'list', 'listlevel', 'listname', 'listoverride',
			'listoverridetable', 'listpicture', 'liststylename', 'listtable', 'listtext', 'lsdlockedexcept', 'macc',
			'maccPr', 'mailmerge', 'maln', 'malnScr', 'manager', 'margPr', 'mbar', 'mbarPr', 'mbaseJc', 'mbegChr',
			'mborderBox', 'mborderBoxPr', 'mbox', 'mboxPr', 'mchr', 'mcount', 'mctrlPr', 'md', 'mdeg', 'mdegHide',
			'mden', 'mdiff', 'mdPr', 'me', 'mendChr', 'meqArr', 'meqArrPr', 'mf', 'mfName', 'mfPr', 'mfunc', 'mfuncPr',
			'mgroupChr', 'mgroupChrPr', 'mgrow', 'mhideBot', 'mhideLeft', 'mhideRight', 'mhideTop', 'mhtmltag',
			'mlim', 'mlimloc', 'mlimlow', 'mlimlowPr', 'mlimupp', 'mlimuppPr', 'mm', 'mmaddfieldname', 'mmath',
			'mmathPict', 'mmathPr', 'mmaxdist', 'mmc', 'mmcJc', 'mmconnectstr', 'mmconnectstrdata', 'mmcPr', 'mmcs',
			'mmdatasource', 'mmheadersource', 'mmmailsubject', 'mmodso', 'mmodsofilter', 'mmodsofldmpdata',
			'mmodsomappedname', 'mmodsoname', 'mmodsorecipdata', 'mmodsosort', 'mmodsosrc', 'mmodsotable', 'mmodsoudl',
			'mmodsoudldata', 'mmodsouniquetag', 'mmPr', 'mmquery', 'mmr', 'mnary', 'mnaryPr', 'mnoBreak', 'mnum',
			'mobjDist', 'moMath', 'moMathPara', 'moMathParaPr', 'mopEmu', 'mphant', 'mphantPr', 'mplcHide', 'mpos',
			'mr', 'mrad', 'mradPr', 'mrPr', 'msepChr', 'mshow', 'mshp', 'msPre', 'msPrePr', 'msSub', 'msSubPr',
			'msSubSup', 'msSubSupPr', 'msSup', 'msSupPr', 'mstrikeBLTR', 'mstrikeH', 'mstrikeTLBR', 'mstrikeV', 'msub',
			'msubHide', 'msup', 'msupHide', 'mtransp', 'mtype', 'mvertJc', 'mvfmf', 'mvfml', 'mvtof', 'mvtol',
			'mzeroAsc', 'mzeroDesc', 'mzeroWid', 'nesttableprops', 'nextfile', 'nonesttables', 'objalias', 'objclass',
			'objdata', 'object', 'objname', 'objsect', 'objtime', 'oldcprops', 'oldpprops', 'oldsprops', 'oldtprops',
			'oleclsid', 'operator', 'panose', 'password', 'passwordhash', 'pgp', 'pgptbl', 'picprop', 'pict', 'pn',
			'pnseclvl', 'pntext', 'pntxta', 'pntxtb', 'printim', 'private', 'propname', 'protend', 'protstart',
			'protusertbl', 'pxe', 'result', 'revtbl', 'revtim', 'rsidtbl', 'rxe', 'shp', 'shpgrp', 'shpinst', 'shppict',
			'shprslt', 'shptxt', 'sn', 'sp', 'staticval', 'stylesheet', 'subject', 'sv', 'svb', 'tc', 'template',
			'themedata', 'title', 'txe', 'ud', 'upr', 'userprops', 'wgrffmtfilter', 'windowcaption', 'writereservation',
			'writereservhash', 'xe', 'xform', 'xmlattrname', 'xmlattrvalue', 'xmlclose', 'xmlname', 'xmlnstbl',
			'xmlopen'
		))
		# Translation of some special characters.
		specialchars = {
			'par': '\n', 'sect': '\n\n', 'page': '\n\n', 'line': '\n', 'tab': '\t', 'emdash': u'\u2014',
			'endash': u'\u2013', 'emspace': u'\u2003', 'enspace': u'\u2002', 'qmspace': u'\u2005',
			'bullet': u'\u2022',
			'lquote': u'\u2018', 'rquote': u'\u2019', 'ldblquote': u'\201C', 'rdblquote': u'\u201D',
		}

		stack = []
		ignorable = False  # Whether this group (and all inside it) are "ignorable".
		unicode_case_skip = 1  # Number of ASCII characters to skip after a unicode character.
		case_skip = 0  # Number of ASCII characters left to skip
		c = 0

		dc = {}
		new_bloc = False  # Indicara si empieza un nuevo bloque
		id_bloc = -1
		id_element = 0
		bloc = ''
		key_element = ''
		properties = False
		ini_picture = False
		image = ''
		n_font = 0
		n_color = 0

		brace_number = 0
		for match in pattern.finditer(self.rtf):
			c += 1
			label, arg, hexadecimal, char, brace, case = match.groups()

			if label == 'stylesheet':
				brace_number += 1

			if bloc == 'fonttbl':
				if label == 'f':
					n_font = int(arg)
					self.fonts[n_font] = self.fonts.get(n_font, ['', ''])

				elif not self.fonts.get(n_font, [''])[0] and label:
					self.fonts[n_font] = self.fonts.get(n_font, ['', ''])
					self.fonts[n_font][0] = label[1:]
				elif label is None and brace is None and case is not None:
					if case != ';':
						self.fonts[n_font] = self.fonts.get(n_font, ['', ''])
						self.fonts[n_font][1] += case

			elif bloc == 'colortbl':
				if label is None:
					continue
				if label == u'red':
					n_color += 1
				self.colors[n_color] = self.colors.get(n_color, ['0', '0', '0'])
				if label == u'red':
					self.colors[n_color][0] = arg
				elif label == u'green':
					self.colors[n_color][1] = arg
				elif label == u'blue':
					self.colors[n_color][2] = arg

			elif bloc == 'stylesheet':
				if label is None and brace is None:
					continue
				if label == 'fs':
					self.default_font_size = int(arg)
				elif label == 'cs':
					self.cs = int(arg)
				elif label == 'cf':
					self.default_foreground = int(arg)
				elif label == 'cb':
					self.default_background = int(arg)
			elif bloc == 'properties':
				if label == u'paperw':
					properties = True
				if not properties:
					continue

				if label == u'paperw':
					self.size[0] = int(arg)
				elif label == u'paperh':
					self.size[1] = int(arg)
				elif label == u'margl':
					self.margins[0] = int(arg)
				elif label == u'margr':
					self.margins[1] = int(arg)
				elif label == u'margt':
					self.margins[2] = int(arg)
				elif label == u'margb':
					self.margins[3] = int(arg)

			if ini_picture:
				if case:
					image += case

			if label == 'pict':
				ini_picture = True

			if brace:
				if brace == '}' and ini_picture:
					ini_picture = False

				if brace == '{' and bloc == 'stylesheet':
					brace_number += 1
				elif brace == '}' and bloc == 'stylesheet':
					brace_number -= 1

				case_skip = 0
				if brace == '{' and bloc not in ['body', 'fonttbl']:
					if bloc != 'stylesheet':
						new_bloc = True
						id_bloc += 1
						id_element = 0
					# Push state
					stack.append((unicode_case_skip, ignorable))

				elif brace == '}' and bloc not in ['body', 'fonttbl']:
					if ini_picture:
						ini_picture = False
					if bloc == 'stylesheet':
						if not brace_number:
							new_bloc = True
							id_bloc += 1
							id_element = 0
							bloc = 'properties'

					unicode_case_skip, ignorable = stack.pop()

			elif char:  # \x (not a letter)
				case_skip = 0
				if char == '~' and not ignorable:
					dc[ajustar(id_bloc), bloc][key_element].append(u'\xA0')
				elif char in '{}\\' and not ignorable:
					dc[ajustar(id_bloc), bloc][key_element].append(char)
				elif char == '*':
					ignorable = True

			elif label:  # Etiquetas
				if label in ['colortbl', 'stylesheet'] or (label == 'plain' and bloc == 'properties'):
					new_bloc = True
					id_bloc += 1
					id_element = 0

				if new_bloc:
					if label in ['rtf', 'fonttbl', 'colortbl', 'stylesheet']:
						bloc = label
					elif label == 'plain':
						bloc = 'body'
					dc[ajustar(id_bloc), bloc] = {}
					new_bloc = False

				key_element = (ajustar(id_element), label, arg)
				if key_element not in dc[ajustar(id_bloc), bloc].keys():
					dc[ajustar(id_bloc), bloc][key_element] = []
				id_element += 1

				case_skip = 0
				if label in destinations:
					ignorable = True
				elif ignorable:
					pass
				elif label in specialchars:
					dc[ajustar(id_bloc), bloc][key_element].append(specialchars[label])
				elif label == 'uc':
					unicode_case_skip = int(arg)
				elif label == 'u':
					c = int(arg)
					if c < 0:
						c += 0x10000
					if c > 127:
						dc[ajustar(id_bloc), bloc][key_element].append(unichr(c))
					else:
						dc[ajustar(id_bloc), bloc][key_element].append(unichr(c))

					case_skip = unicode_case_skip

			elif hexadecimal:  # \'xx
				if case_skip > 0:
					case_skip -= 1
				elif not ignorable:
					c = int(hexadecimal, 16)
					if c > 127:
						dc[ajustar(id_bloc), bloc][key_element].append(unichr(c))
					else:
						dc[ajustar(id_bloc), bloc][key_element].append(unichr(c))

			elif case:
				ignorable = False
				if case == '}' and ini_picture:
					ini_picture = False
					continue
				if ini_picture:
					continue
				if case_skip > 0:
					case_skip -= 1
				elif not ignorable:
					if (ajustar(id_bloc), bloc) in dc.keys():
						dc[ajustar(id_bloc), bloc][key_element].append(case)
		self.parse_dc(dc)

	def parse_dc(self, dc):
		d = {}

		def font_format(_type, _value, _is_bold, _is_italic, _is_underline):
			_find = 0
			# Negrita
			if _type == 'b':
				_find = 1
				if _value is None:
					_is_bold = True
				else:
					_is_bold = False

			# Cursiva
			elif _type == 'i':
				_find = 1
				if _value is None:
					_is_italic = True
				else:
					_is_italic = False

			# subrayado
			elif _type == 'ul':
				_find = 1
				if _value is None:
					_is_underline = True
				else:
					_is_underline = False

			# Si el tipo pasa a ser plano se reinician las negritas, cursiva y subrayado
			if _type == 'plain':
				_is_underline = False
				_is_italic = False
				_is_bold = False
			return [_is_bold, _is_italic, _is_underline, _find]

		def font_style(_type, _value, _font, _font_size, _font_color, _find):
			if _type == 'f':
				_find = 1
				_font = _value
			elif _type == 'fs':
				_find = 1
				_font_size = _value
			elif _type == 'cf':
				_find = 1
				_font_color = _value
			return [_font, _font_size, _font_color, _find]

		font = None
		font_size = None
		font_color = None

		is_bold = False
		is_italic = False
		is_underline = False

		ini_table = False
		ini_list = False

		open_lang = False

		n_par = 0
		li = ''
		type_list = ''
		number_list = 1

		keys_1 = dc.keys()
		keys_1.sort()
		for key in keys_1:
			_id, _bloc = key
			if _bloc != 'body':
				continue
			ke = str(key)
			d[ke] = {}
			keys_2 = dc[key].keys()
			keys_2.sort()
			for key2 in keys_2:
				_id_element, _type, _value = key2
				is_bold, is_italic, is_underline, find = font_format(_type, _value, is_bold, is_italic, is_underline)
				font, font_size, font_color, find = font_style(_type, _value, font, font_size, font_color, find)

				txt = ''.join(dc[key][key2])

				if _type in ['plain']:
					font = 1
					font_size = self.default_foreground
					font_color = self.default_font_size
					if _value is None and not txt:
						continue

				## Se comprueba si hay algún salto de página mezclado con el texto
				if '\n' in txt and txt != '\n':
					aux = txt.split('\n')
				else:
					aux = [txt]

				if _type in ['par']:
					if ini_table:
						ini_table = False
						self.elements.append(
							RtfParagraph(self)
						)
						self.elements[-1].texts.append(
							RtfText(self.elements[-1], '\n')
						)
					elif ini_list:
						if n_par == 1:
							self.elements.append(
								RtfParagraph(self)
							)
							self.elements[-1].texts.append(RtfText(self, '\n'))
							n_par = 0
							ini_list = False
						txt = txt.replace('\n', '')
						if txt:
							if self.elements[-1].name == 'RtfList':
								self.elements[-1].paragraphs.append(
									RtfParagraph(self.elements[-1])
								)
								self.elements[-1].paragraphs[-1].text_indent = li
								self.elements[-1].paragraphs[-1].type_list = type_list
								self.elements[-1].paragraphs[-1].number = number_list
								number_list += 1
								self.elements[-1].paragraphs[-1].texts.append(
									RtfText(self.elements[-1].paragraphs[-1], txt, is_bold, is_italic, is_underline,
									        font_size, font_color, font)
								)
							else:
								self.elements[-1].texts.append(
									RtfText(
										self.elements[-1], txt, is_bold, is_italic, is_underline, font_size, font_color,
										font)
								)
						continue
					elif self.elements and self.elements[-1].name == 'RtfParagraph':

						pass

				if _type in ['lang']:
					if open_lang:
						if self.elements[-1].texts:
							self.elements.append(
								RtfParagraph(self)
							)
							#print aux
							'''self.elements[-1].texts.append(
								RtfText(self.elements[-1], '\n')
							)'''
						open_lang = False
					else:
						open_lang = True

				elif _type == 'li':
					li = _value

				elif _type == 'pnstart':
					type_list = 'n'

				elif _type == 'pntext':
					type_list = 'q'

				elif _type == 'ls' and not ini_list:
					ini_list = True
					number_list = 1
					self.elements.append(
						RtfList(self)
					)

				elif _type in ['pard']:
					if ini_list:
						n_par += 1
						continue
					elif not ini_table:
						self.elements.append(
							RtfParagraph(self)
						)
					else:
						self.elements[-1].rows[-1].cells[-1].paragraphs.append(
							RtfParagraph(self)
						)

				elif _type == 'trowd':
					if not ini_table:
						ini_table = True
						self.elements.append(
							RtfTable(self)
						)
					self.elements[-1].add_row()

				elif _type == 'trgaph':
					self.elements[-1].rows[-1].set_horizontal_gap(_value)
					continue

				elif _type == 'trleft':
					self.elements[-1].rows[-1].set_left(_value)
					continue

				elif _type == 'clvertalt':
					self.elements[-1].rows[-1].add_cell()
					self.elements[-1].rows[-1].cells[-1].set_vertical_alignment(_value)
					continue

				elif _type == 'cellx':
					self.elements[-1].rows[-1].cells[-1].set_width(_value)
					continue

				elif _type in ['qc', 'ql', 'qr', 'qj']:
					n_par = 0
					val = None
					if _type in ['qc', 'ql', 'qr', 'qj']:
						val = _type[1:]

					if not self.elements:
						self.elements.append(
							RtfParagraph(self)
						)

					if self.elements[-1].name == 'RtfTable':
						self.elements[-1].rows[-1].cells[-1].set_horizontal_alignment(val)
					elif self.elements[-1].name == 'RtfList':
						if not txt:
							ini_list = False
							self.elements.append(
								RtfParagraph(self.elements[-1])
							)
							continue
						self.elements[-1].paragraphs.append(
							RtfParagraph(self.elements[-1])
						)
						self.elements[-1].paragraphs[-1].text_indent = li
						self.elements[-1].paragraphs[-1].type_list = type_list
						self.elements[-1].paragraphs[-1].number = number_list
						number_list += 1
						self.elements[-1].paragraphs[-1].texts.append(
							RtfText(
								self.elements[-1].paragraphs[-1], txt, is_bold, is_italic, is_underline, font_size,
								font_color, font
							)
						)
					elif self.elements[-1].name == 'RtfParagraph':
						if (val is not None and self.elements[-1].horizontal_alignment is not None) and \
								val != self.elements[-1].horizontal_alignment:
							self.elements.append(
								RtfParagraph(self)
							)

						self.elements[-1].texts.append(
							RtfText(self.elements[-1], txt, is_bold, is_italic, is_underline, font_size, font_color, font)
						)
						self.elements[-1].set_horizontal_alignment(val)
					else:
						raise ValueError(
							str((key2, txt, self.elements[-1].name))
						)
					continue

				elif _type in ['s', 'intbl', 'cell'] and ini_table:
					continue

				for a in range(len(aux)):
					ke2 = str((key2[0] + '-' + str(a), key2[1], key2[2]))
					d[ke][ke2] = aux[a]

					if aux[a]:
						if not self.elements:
							self.elements.append(RtfParagraph(self))

						if self.elements[-1].name == 'RtfParagraph':
							self.elements[-1].texts.append(
								RtfText(self.elements[-1], aux[a], is_bold, is_italic, is_underline, font_size,
										font_color, font)
							)
						if self.elements[-1].name == 'RtfTable':
							self.elements[-1].rows[-1].cells[-1].paragraphs[-1].texts.append(
								RtfText(self.elements[-1].rows[-1].cells[-1].paragraphs[-1], aux[a], is_bold, is_italic,
										is_underline, font_size, font_color, font)
							)

		return d




	def parse_dc_old(self, dc):
		d = {}

		def font_format(_type, _value, _is_bold, _is_italic, _is_underline):
			_find = 0
			# Negrita
			if _type == 'b':
				_find = 1
				if _value is None:
					_is_bold = True
				else:
					_is_bold = False

			# Cursiva
			elif _type == 'i':
				_find = 1
				if _value is None:
					_is_italic = True
				else:
					_is_italic = False

			# subrayado
			elif _type == 'ul':
				_find = 1
				if _value is None:
					_is_underline = True
				else:
					_is_underline = False

			# Si el tipo pasa a ser plano se reinician las negritas, cursiva y subrayado
			if _type == 'plain':
				_is_underline = False
				_is_italic = False
				_is_bold = False
			return [_is_bold, _is_italic, _is_underline, _find]

		def font_style(_type, _value, _font, _font_size, _font_color, _background, _find):
			if _type == 'f':
				_find = 1
				_font = _value
			elif _type == 'fs':
				_find = 1
				_font_size = _value
			elif _type == 'cf':
				_find = 1
				_font_color = _value
			elif _type == 'clcbpat':
				_find = 1
				_background = _value
			elif _type in ['plain', 'pard']:
				_find = 1
				# _background = '1'
				_font_color = self.default_foreground
				_font = 1
				_font_size = self.default_font_size
			return [_font, _font_size, _font_color, _background, _find]

		font = None
		font_size = None
		font_color = None
		background = '1'
		is_bold = False
		is_italic = False
		is_underline = False
		h_alignment = ''

		ini_table = False
		ini_list = False

		highlight = None
		n_par = 0
		li = ''
		type_list = ''
		number_list = 1
		key_border = ''

		keys_1 = dc.keys()
		keys_1.sort()
		for key in keys_1:
			_id, _bloc = key
			if _bloc != 'body':
				continue
			ke = str(key)
			d[ke] = {}
			keys_2 = dc[key].keys()
			keys_2.sort()
			for key2 in keys_2:
				_id_element, _type, _value = key2
				is_bold, is_italic, is_underline, find = font_format(_type, _value, is_bold, is_italic, is_underline)
				font, font_size, font_color, background, find = font_style(_type, _value, font, font_size, font_color,
																		   background, find)

				txt = ''.join(dc[key][key2])

				if _type in ['plain']:
					if _value is None and not txt:
						continue

				# Se comprueba si hay algún salto de página mezclado con el texto
				if '\n' in txt and txt != '\n':
					aux = txt.split('\n')
				else:
					aux = [txt]

				if _type in ['par']:
					if ini_list:
						if n_par == 1:
							self.elements.append(
								RtfParagraph(self, self)
							)
							self.elements[-1].texts.append(RtfText(self, self.elements[-1], '\n'))
							n_par = 0
							ini_list = False
						txt = txt.replace('\n', '')
						if txt:
							if self.elements[-1].name == 'RtfList':
								self.elements[-1].paragraphs.append(
									RtfParagraph(self, self.elements[-1])
								)
								self.elements[-1].paragraphs[-1].text_indent = li
								self.elements[-1].paragraphs[-1].type_list = type_list
								self.elements[-1].paragraphs[-1].number = number_list
								number_list += 1
								self.elements[-1].paragraphs[-1].texts.append(
									RtfText(self, self.elements[-1].paragraphs[-1], txt, is_bold, is_italic,
											is_underline,
											font_size, font_color, font)
								)
							else:
								self.elements[-1].texts.append(
									RtfText(self,
											self.elements[-1], txt, is_bold, is_italic, is_underline, font_size,
											font_color,
											font)
								)
						continue
					elif self.elements and self.elements[-1].name == 'RtfParagraph':
						self.elements.append(
							RtfParagraph(self, self)
						)

				elif _type == 'li':
					li = _value

				elif _type == 'pnstart':
					type_list = 'n'

				elif _type == 'pntext':
					type_list = 'q'

				elif _type == 'ls' and not ini_list:
					ini_list = True
					number_list = 1
					self.elements.append(
						RtfList(self)
					)

				elif _type in ['pard']:
					if ini_list:
						n_par += 1
						continue
					elif not ini_table:
						self.elements.append(
							RtfParagraph(self, self)
						)
					else:
						if not self.elements[-1].rows:
							self.elements[-1].add_row()
						if not self.elements[-1].rows[-1].cells:
							self.elements[-1].rows[-1].add_cell()
						# self.elements[-1].rows[-1].cells[-1].set_background_colour(background)
						self.elements[-1].rows[-1].cells[-1].paragraphs.append(
							RtfParagraph(self, self.elements[-1].rows[-1].cells[-1])
						)

				elif _type == 'trowd':
					if not ini_table:
						ini_table = True
						self.elements.append(
							RtfTable(self, self)
						)
					self.elements[-1].add_row()

				elif _type == 'row':
					self.elements[-1].add_row()

				elif _type == 'trgaph':
					self.elements[-1].rows[-1].set_horizontal_gap(_value)
					continue

				elif _type == 'trleft':
					self.elements[-1].rows[-1].set_left(_value)
					continue

				elif _type.startswith('clbrdr'):
					if _type == 'clbrdrt':
						key_border = 'top'
					elif _type == 'clbrdrb':
						key_border = 'bottom'
					elif _type == 'clbrdrl':
						key_border = 'left'
					elif _type == 'clbrdrr':
						key_border = 'right'

					if key_border:
						self.elements[-1].rows[-1].borders[-1][key_border] = 1
					continue

				elif _type == 'brdrw':
					self.elements[-1].rows[-1].borders[-1][key_border] = _value

				elif _type == 'highlight':
					highlight = _value

				elif _type == 'clvertalt':
					self.elements[-1].rows[-1].borders.append({})
					'''if not self.elements[-1].rows[-1].cells:
						self.elements[-1].rows[-1].add_cell()
					self.elements[-1].rows[-1].cells[-1].set_vertical_alignment(_value)
					self.elements[-1].rows[-1].cells[-1].set_background_colour(background)'''
					continue

				elif _type == 'cellx':
					self.elements[-1].rows[-1].colors.append(background)
					background = '1'
					if len(self.elements[-1].rows) < 2:
						self.elements[-1].add_width(_value)
					continue

				elif _type in ['qc', 'ql', 'qr', 'qj']:
					n_par = 0
					val = _type[1:]

					h_alignment = val
					if not self.elements:
						self.elements.append(
							RtfParagraph(self, self)
						)

					if self.elements[-1].name == 'RtfTable':
						if txt:
							self.elements[-1].rows[-1].cells[-1].paragraphs[-1].texts.append(
								RtfText(self, self.elements[-1].rows[-1].cells[-1].paragraphs[-1],
										txt, is_bold, is_italic, is_underline, font_size, font_color, font)
							)

						self.elements[-1].rows[-1].cells[-1].paragraphs[-1].set_horizontal_alignment(val)
					elif self.elements[-1].name == 'RtfList':
						if not txt:
							ini_list = False
							self.elements.append(
								RtfParagraph(self, self.elements[-1])
							)
							continue
						self.elements[-1].paragraphs.append(
							RtfParagraph(self, self.elements[-1])
						)
						self.elements[-1].paragraphs[-1].text_indent = li
						self.elements[-1].paragraphs[-1].type_list = type_list
						self.elements[-1].paragraphs[-1].number = number_list
						number_list += 1
						self.elements[-1].paragraphs[-1].texts.append(
							RtfText(self,
									self.elements[-1].paragraphs[-1], txt, is_bold, is_italic, is_underline, font_size,
									font_color, font
									)
						)
					elif self.elements[-1].name == 'RtfParagraph':
						if (val is not None and self.elements[-1].horizontal_alignment is not None) and \
								val != self.elements[-1].horizontal_alignment:
							self.elements.append(
								RtfParagraph(self, self)
							)

						self.elements[-1].texts.append(
							RtfText(self, self.elements[-1], txt, is_bold, is_italic, is_underline, font_size,
									font_color,
									font)
						)
						self.elements[-1].set_horizontal_alignment(val)
					else:
						raise ValueError(
							str((key2, txt, self.elements[-1].name))
						)
					continue

				elif _type in ['s', 'intbl', 'cell'] and ini_table:
					continue

				for a in range(len(aux)):
					ke2 = str((key2[0] + '-' + str(a), key2[1], key2[2]))
					d[ke][ke2] = aux[a]

					if aux[a]:
						if not self.elements:
							self.elements.append(RtfParagraph(self, self))

						if self.elements[-1].name == 'RtfParagraph':
							self.elements[-1].texts.append(
								RtfText(self, self.elements[-1], aux[a], is_bold, is_italic, is_underline, font_size,
										font_color, font)
							)
						if self.elements[-1].name == 'RtfTable':
							ha = self.elements[-1].rows[-1].cells[-1].get_horizontal_alignment()

							txt_cell = RtfText(self, self.elements[-1].rows[-1].cells[-1].paragraphs[-1], aux[a],
											   is_bold,
											   is_italic, is_underline, font_size, font_color, font)
							txt_cell.highlight = highlight
							highlight = None
							self.elements[-1].rows[-1].cells[-1].paragraphs[-1].horizontal_alignment = ha
							self.elements[-1].rows[-1].cells[-1].paragraphs[-1].texts.append(txt_cell)

		return d
