#!/usr/bin/python
# -*- coding: utf-8 -*-

import re


def get_color(rtf, font_color):
	if font_color is not None:
		c = rtf.colors.get(int(font_color), ['0', '0', '0'])
		for _c in range(len(c)):
			c[_c] = int(c[_c])
		return '%02x%02x%02x' % tuple(c)
	return str()


class RtfCell:
	def __init__(self, rtf, parent, width='2000'):
		self.parent = parent
		self.rtf = rtf
		self.name = 'RtfCell'
		self.width = round((int(width) * 100.) / self.rtf.size[0], 2)
		self.paragraphs = list()
		self.horizontal_alignment = 'left'
		self.vertical_alignment = 'None'
		self.adjust = 15
		self.indent = parent.indent + 1
		self.background = 1

	def set_width(self, value):
		self.width = round((int(value) * 100.) / float(self.rtf.size[0]), 2)

	def get_width(self):
		return self.width

	def set_background_colour(self, color):
		self.background = color

	def get_background_colour(self):
		return self.background

	def set_horizontal_alignment(self, value):
		self.horizontal_alignment = value

	def get_horizontal_alignment(self):
		return self.horizontal_alignment

	def set_vertical_alignment(self, value):
		self.vertical_alignment = value

	def get_vertical_alignment(self):
		return self.vertical_alignment

	def __str__(self):
		dc = dict()
		for key in self.__dict__.keys():
			if key in ['parent', 'rtf', 'paragraphs']:
				continue
			dc[key] = eval(repr(self.__dict__[key]))
		value = str()
		for p in self.paragraphs:
			value += '\t\t\t' + p.__str__()
		dc['n_paragraphs'] = len(self.paragraphs)
		return str(dc) + '\n' + value

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

			borders = {}
			if self.parent.borders:
				borders = self.parent.borders.pop(0)

			background = get_color(self.rtf, self.background)
			if len(self.parent.colors):
				background = get_color(self.rtf, self.parent.colors.pop(0))

			tx = ''
			for paragraph in self.paragraphs:
				v = [paragraph.get_value(mode)]

				if v:
					if background or borders:

						tx = ' style='
						if background:
							tx += '"background-color: #%s;' % background
						if borders:
							for key in borders:
								if not tx.endswith(' '):
									tx += ' '
								tx += 'border-%s:%spx solid #000000;' % (key, str(borders[key]))
						tx += '"'

				value.extend(v)

			value.insert(0, '%s<td%s>\n' % (('\t' * self.indent), tx))
			value.append('\n%s</td>\n' % ('\t' * self.indent))

			value = ''.join(value)
		elif mode == 'word':
			cell = word_object.add_cell([], horizontal_alignment='l')
			cell.get_properties().set_shading(get_color(self.rtf, self.parent.get_colors().pop(0)))
			# cell = Table.Cell(word_object, len(word_object.cells) + 1, [])
			for paragraph in self.paragraphs:
				paragraph.get_value(mode, word_object=cell)
		return value


class RtfRow:
	def __init__(self, rtf, parent):
		self.rtf = rtf
		self.parent = parent
		self.indent = parent.indent + 2
		self.widths = list()
		self.colors = list()
		self.borders = list()
		self.vertical_align = list()
		self.horizontal_align = list()
		self.cells = list()
		self.number_cells = 0

	def get_widths(self):
		if self.widths:
			return self.widths
		return self.parent.get_widths()

	def set_widths(self, widths):
		self.widths = widths

	def add_width(self, width):
		self.widths.append(width)

	def get_colors(self):
		return self.colors

	def set_colors(self, colors):
		self.colors = colors

	def add_colors(self, color):
		self.colors.append(color)

	def get_vertical_align(self):
		return self.vertical_align

	def set_vertical_align(self, vertical_align):
		self.vertical_align = vertical_align

	def add_vertical_align(self, vertical_align):
		self.vertical_align.append(vertical_align)

	def get_horizontal_align(self):
		return self.horizontal_align

	def set_horizontal_align(self, horizontal_align):
		self.horizontal_align = horizontal_align

	def add_horizontal_align(self, horizontal_align):
		self.horizontal_align.append(horizontal_align)

	def get_borders(self):
		return self.borders

	def set_borders(self, borders):
		self.borders = borders

	def add_border(self, key, value):
		self.borders[key] = value

	def add_cell(self):
		self.cells.append(RtfCell(self.rtf, self))

	def __str__(self):
		value = str()
		dc = dict()
		dc['n_cells'] = len(self.cells)
		dc['widths'] = self.get_widths()
		dc['colors'] = self.get_colors()
		dc['v_align'] = self.get_vertical_align()
		dc['borders'] = self.get_borders()

		for cell in self.cells:
			cell.adjust = 0
			value += '\t\t' + cell.__str__()

		return str(dc) + '\n' + value

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
				value.insert(0, '%s<tr>\n' % ('\t' * self.indent))
				value.append('%s</tr>\n' % ('\t' * self.indent))
			value = ''.join(value)

		elif mode in ['word']:
			va = []
			for ln in self.get_vertical_align():
				if ln == 't':
					va.append('top')
				elif ln == 'b':
					va.append('bottom')
				else:
					va.append('center')
			row = word_object.add_row()
			if not self.widths:
				self.widths = self.parent.get_widths()
			for k in range(len(self.cells)):
				cell = self.cells[k]
				cell.set_width(self.widths[k])
				cell.get_value(mode, row)
			row.set_vertical_alignment(va)

			ls = list()
			for b in self.get_borders():
				bo = {}
				for key in b.keys():
					bo[key] = {'sz': b[key]}
				ls.append(bo)

			row.set_cell_borders(ls)
		return value


class RtfTable:
	def __init__(self, rtf, parent):
		self.name = 'RtfTable'
		self.rows = list()
		self.parent = parent
		self.widths = list()
		self.indent = parent.indent + 1
		self.rtf = rtf
		self.horizontal_gap = '60'
		self.left = '0'

	def get_widths(self):
		return self.widths

	def set_widths(self, widths):
		self.widths = widths

	def get_rows(self):
		return self.rows

	def set_rows(self, rows):
		self.rows = rows

	def add_width(self, width):
		self.widths.append(width)

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
			for row in self.rows:
				value += row.get_value(mode)
				value += '\n'

		elif mode in ['html']:
			value = list()
			for row in self.rows:
				value.append(row.get_value(mode))
			if value:
				value.insert(0, '%s<tbody>\n' % ('\t' * (self.indent + 1)))
				value.insert(0, '%s</thead>\n' % ('\t' * (self.indent + 1)))
				value.insert(0, '%s<thead>\n' % ('\t' * (self.indent + 1)))
				value.insert(0, '%s</colgroup>\n' % ('\t' * (self.indent + 1)))
				w = 0
				total_with = 0
				for width in self.widths:
					_w = int(width) - w
					w_rtf = float(self.rtf.size[0]) - float(self.rtf.margins[0]) - float(self.rtf.margins[1])
					_w = int(round((_w * 100.) / w_rtf, 0))
					total_with += _w
					w = int(width)

					value.insert(0, '%s<col style="width: %s%s;">\n' % (('\t' * (self.indent + 2)), str(_w), '%'))
				value.insert(0, '%s<colgroup>\n' % ('\t' * (self.indent + 1)))
				value.insert(0, '%s<table style="width: %s%s; border-collapse: collapse">\n' % (
					'\t' * self.indent, str(total_with), '%'))

				value.append('%s</tbody>\n' % ('\t' * (self.indent + 1)))
				value.append('%s</table>\n' % ('\t' * self.indent))
			value = ''.join(value)

		elif mode == 'word':
			if not self.rows:
				return value

			cw = []
			for k in range(len(self.rows[0].get_widths())):
				if k == 0:
					cw.append(int(self.rows[0].get_widths()[k]))
				else:
					cw.append(int(self.rows[0].get_widths()[k]) - int(self.rows[0].get_widths()[k-1]))

			value = self.parent.parent.add_table()

			for row in self.rows:
				row.get_value(mode, value)

			value.set_columns_width(cw)

		return value

	def __str__(self):
		value = []
		for row in self.rows:
			value.append(row.__str__())
		return '\n'.join(value)

	def add_row(self):
		self.rows.append(RtfRow(self.rtf, self))


class RtfParagraph:
	def __init__(self, rtf, parent, h_alignment=None, text_indent=0, id_paragraph=0):
		self.name = 'RtfParagraph'
		self.parent = parent
		self.texts = list()
		self.text_indent = text_indent
		self.type_list = ''
		self.number = 1
		self.id = id_paragraph
		self.horizontal_alignment = h_alignment
		self.rtf = rtf
		self.indent = parent.indent + 1
		self.ha = {'l': 'left', 'r': 'right', 'c': 'center', 'j': 'justify'}

	def __str__(self):
		dc = dict()
		for key in self.__dict__.keys():
			if key in ['parent', 'rtf', 'texts']:
				continue
			dc[key] = eval(repr(self.__dict__[key]))
		value = str()
		for t in self.texts:
			value += '\t-\t\t\t\t' + t.__str__()+'\n'
			
		return self.name+'\n'+str(dc) + '-----\n' + value + '\n'

	def set_horizontal_alignment(self, alignment):
		self.horizontal_alignment = alignment

	def get_value(self, mode='plain', word_object=None):
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

			if hasattr(self.parent, 'name') and self.parent.name == 'RtfTables':
				for text in self.texts:
					value += text.get_value(mode, text_align=self.horizontal_alignment)
				if value and value in [u'</br>\n']:
					value = (u'\t' * self.indent) + value
			else:
				value = ''
				for text in self.texts:
					x = text.get_value(mode)
					if not x:
						continue
					value += x + '\n'

				if value and value in ['</br>\n']:
					value = ('\t' * self.indent) + value
				elif value:
					style = ''
					if self.horizontal_alignment:
						style = u' style="text-align: %s;"' % (
							self.ha.get(self.horizontal_alignment, self.horizontal_alignment))
					if self.text_indent:
						if not style:
							style = u' style='
						else:
							style += u' '
						style += u'%s<"text-indent: %dpx";' % (('\t' * self.indent), int(self.text_indent))

					value_ = u'%s<p%s>\n' % ('\t' * self.indent, style)
					value = value_ + value + u'%s</p>\n' % ('\t' * self.indent)

		elif mode == 'word':
			paragraph = word_object.add_paragraph(_text=(), horizontal_alignment='l', font_format='', font_size=None)
			if self.horizontal_alignment is not None:
				paragraph.set_horizontal_alignment(self.horizontal_alignment)

			for text in self.texts:
				text.get_value(mode, word_object=paragraph)

		return value


class RtfText:
	def __init__(self, rtf, parent, text='', bold=False, italic=False, underline=False, font_size=10., font_color=1,
					font=1, highlight=None):
		self.name = 'RtfText'
		self.text = text
		self.bold = bold
		self.italic = italic
		self.underline = underline
		self.font_size = font_size
		self.font_color = font_color
		self.font = font
		self.parent = parent
		self.rtf = rtf
		self.indent = parent.indent + 1
		self.highlight = highlight

	def __str__(self):
		txt = unicode(self.text.encode('utf-8'), errors='ignore')
		if len(txt) > 27:
			txt = txt[:27] + u'...'

		dc = dict()
		for key in self.__dict__.keys():
			if key in ['parent', 'rtf', 'text']:
				continue
			dc[key] = eval(repr(self.__dict__[key]))
		dc['text'] = txt
		return str(dc)

	def get_value(self, mode='plain', word_object=None):
		colour = get_color(self.rtf, self.font_color)

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
			if txt in ['', '\n']:
				return

			text = word_object.add_text(txt, font_format, int(self.font_size) // 2)
			text.set_font(self.rtf.fonts.get(int(self.font))[1])

			if colour:
				text.get_properties().set_color(colour)

		elif mode == 'html':
			value = [self.text.replace(u'\n', u'')]
			if value == ['']:
				return ''
			if value in [[u'</br>'], u'</br>']:
				return u'%s</br>' % ('\t' * self.indent)

			if self.bold:
				value.insert(0, u'<b>')
				value.append(u'</b>')
			if self.italic:
				value.insert(0, u'<i>')
				value.append(u'</i>')
			if self.underline:
				value.insert(0, u'<u>')
				value.append(u'</u>')

			label = u'<span style="'
			try:
				fs = round(float(self.font_size) / 20., 1)
			except Exception as e:
				raise ValueError(e.message)

			label += u'font-size: %.1fem; ' % fs
			if colour:
				label += u'color:#%s; ' % colour
			if self.highlight:
				label += u'background-color:#%s; ' % get_color(self.rtf, self.highlight)
			try:
				label += u"font-family:'%s'" % self.rtf.fonts.get(int(self.font), [])[1]
			except Exception as e:
				raise ValueError(e.message)
			if label and not label[-1]:
				label = label[:-1]
			label += u'">'

			value.insert(0, (u'\t' * self.indent) + label)

			value.append(u'</span>')

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
		self.name = 'rtf'
		self.elements = list()
		self.rtf = text
		self.parent = parent
		self.indent = 1

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
			txt = element.get_value(mode, word_object=self.parent)
			if mode != 'word' and txt is not None:
				value += txt
		if mode == 'list':
			value = value.split('\n')
		elif mode == 'html':
			_value = '<html>\n'
			_value += '%s<body>\n' % '\t' * self.indent
			value = _value + value
			value += '%s</body>\n' % '\t' * self.indent
			value += '<html>\n'
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

		def font_format(_type, _value, _is_bold, _is_italic, _is_underline, _h_alignment):
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

			elif _type in ['qc', 'ql', 'qr', 'qj']:
				_h_alignment = _type[1:]

			# Si el tipo pasa a ser plano se reinician las negritas, cursiva y subrayado
			elif _type == 'plain':
				_is_underline = False
				_is_italic = False
				_is_bold = False

			elif _type == 'pard':
				_h_alignment = 'l'

			return [_is_bold, _is_italic, _is_underline, _h_alignment, _find]

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

		def table(_type, _value, _ini_table, _key_border, _background, _is_row):
			if _type == 'trowd':
				if not _ini_table:
					self.elements.append(
						RtfTable(self, self)
					)
					_ini_table = True
					self.elements[-1].add_row()

			elif _type == 'trgaph':
				self.elements[-1].set_horizontal_gap(_value)

			elif _type == 'trleft':
				self.elements[-1].set_left(_value)

			elif _type.startswith('clbrdr'):
				if _type == 'clbrdrt':
					_key_border = 'top'
				elif _type == 'clbrdrb':
					_key_border = 'bottom'
				elif _type == 'clbrdrl':
					_key_border = 'left'
				elif _type == 'clbrdrr':
					_key_border = 'right'

				if _key_border:
					self.elements[-1].rows[-1].borders[-1][_key_border] = 1

			elif _type == 'brdrw':
				self.elements[-1].rows[-1].borders[-1][_key_border] = _value

			elif _type.startswith('clvertal'):

				_is_row = False
				self.elements[-1].rows[-1].borders.append({})
				self.elements[-1].rows[-1].add_vertical_align(_type[-1])
				_key_border = None
				_background = self.default_background

			elif _type == 'cellx':
				self.elements[-1].rows[-1].colors.append(_background)
				if len(self.elements[-1].rows) == 1:
					self.elements[-1].add_width(_value)
				self.elements[-1].rows[-1].number_cells += 1

			elif _type == 'row':
				self.elements[-1].add_row()
				_is_row = True

			elif _type == 'pard' and _is_row:
				_ini_table = False
				if not self.elements[-1].rows[-1].cells:
					del self.elements[-1].rows[-1]

			return [_key_border, _background, _ini_table, _is_row]

		font = 1
		font_size = self.default_font_size
		font_color = self.default_foreground
		background = self.default_background
		bold = False
		italic = False
		underline = False
		h_alignment = 'l'

		ini_table = False
		is_row = False
		ini_list = False
		ini_paragraph = False
		forzar_salto_linea = False

		highlight = None
		n_par = 0
		li = ''
		type_list = ''
		number_list = 1
		key_border = ''
		last_type = ''

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
				last_type = _type

				bold, italic, underline, h_alignment, find = font_format(_type, _value, bold, italic, underline,
																			h_alignment)
				font, font_size, font_color, background, find = font_style(_type, _value, font, font_size, font_color,
																			background, find)
				key_border, background, ini_table, is_row = table(_type, _value, ini_table, key_border, background,
																	is_row)

				txt = ''.join(dc[key][key2])

				if _type in ['plain']:
					if _value is None and not txt:
						continue

				# Se comprueba si hay algún salto de página mezclado con el texto
				if '\n' in txt and txt != '\n':
					aux = txt.split('\n')
				else:
					aux = [txt]

				if _type == 'li':
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

				elif _type in ['par']:
					if not ini_paragraph and aux == ['\n']:
						forzar_salto_linea = True
					else:
						ini_paragraph = False

				elif _type in ['pard']:
					if ini_list:
						n_par += 1
					elif ini_table:
						self.elements[-1].rows[-1].add_cell()
					continue

				elif _type == 'highlight':
					highlight = _value

				for a in range(len(aux)):
					ke2 = str((key2[0] + '-' + str(a), key2[1], key2[2]))
					d[ke][ke2] = aux[a]

					if aux[a].replace('\n', '') or forzar_salto_linea:

						if not self.elements or (not ini_paragraph and not ini_table):
							self.elements.append(RtfParagraph(self, self, h_alignment))
							ini_paragraph = True
						if self.elements[-1].name == 'RtfParagraph':
							self.elements[-1].texts.append(
								RtfText(self, self.elements[-1], aux[a], bold, italic, underline, font_size,
										font_color, font, highlight)
							)
						if self.elements[-1].name == 'RtfTable':

							if not self.elements[-1].rows[-1].cells[-1].paragraphs:
								self.elements[-1].rows[-1].cells[-1].paragraphs.append(
									RtfParagraph(self, self.elements[-1].rows[-1].cells[-1], h_alignment))

							txt_cell = RtfText(self, self.elements[-1].rows[-1].cells[-1].paragraphs[-1], aux[a],
												bold, italic, underline, font_size, font_color, font, highlight)

							highlight = None

							self.elements[-1].rows[-1].cells[-1].paragraphs[-1].texts.append(txt_cell)

				forzar_salto_linea = False

		return d
