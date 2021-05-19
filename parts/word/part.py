#!/usr/bin/python
# -*- coding: utf-8 -*-

import elements
import rtf


class Part:
	"""
	classdocs
	"""

	def __init__(self, parent, part_type, number=1, type_reference=''):
		self.tag = 'w:%s' % part_type
		self.type_reference = type_reference
		self.rId = 0
		if part_type is 'hdr':
			name_ = 'header'
			self.relRId = 1
		elif part_type is 'ftr':
			name_ = 'footer'
			self.relRId = 100
		else:
			raise ValueError("Parte %s no reconocida." % part_type)
		self.content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.%s+xml" % name_
		self.name = 'word/%s%d.xml' % (name_, number)

		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.get_separator()

		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if not self.separator:
			self.xml_header += '\n'

		self.indent = 1

		self.attributes = list()
		self.attributes.append('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"')
		self.attributes.append('xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"')
		self.attributes.append('xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"')
		self.attributes.append('xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"')
		self.attributes.append('xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"')
		self.attributes.append('xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"')
		self.attributes.append('xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"')
		self.attributes.append('xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"')
		self.attributes.append('xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"')
		self.attributes.append('xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"')
		self.attributes.append('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
		self.attributes.append('xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"')
		self.attributes.append('xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"')
		self.attributes.append('xmlns:o="urn:schemas-microsoft-com:office:office"')
		self.attributes.append('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
		self.attributes.append('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"')
		self.attributes.append('xmlns:v="urn:schemas-microsoft-com:vml"')
		self.attributes.append('xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"')
		self.attributes.append('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"')
		self.attributes.append('xmlns:w10="urn:schemas-microsoft-com:office:word"')
		self.attributes.append('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')
		self.attributes.append('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')
		self.attributes.append('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"')
		self.attributes.append('xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"') #
		self.attributes.append('xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"')
		self.attributes.append('xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"') #
		self.attributes.append('xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"') #
		self.attributes.append('xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"')
		self.attributes.append('xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"')
		self.attributes.append('xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"')
		self.attributes.append('xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"')
		self.attributes.append('xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
		# self.attributes.append('mc:Ignorable="w14 w15 w16se w16cid wp14"')
		self.attributes.append('mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"')

		self.elements = list()
		''' Id's '''
		self.rsidR = '00D4598A'
		self.rsidRDefault = '00D4598A'

	def get_TypeReference(self):
		return self.type_reference

	def SetRId(self, value):
		self.rId = value

	def get_RId(self):
		return self.rId

	def AddRelRId(self):
		self.relRId += 1

	def get_RelRId(self):
		return self.relRId

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def ContentType(self):
		return self.content_type

	def get_Tag(self):
		return self.tag

	def get_name(self):
		return self.name

	def get_separator(self):
		return self.separator

	def get_parent(self):
		return self.parent

	def get_xmlHeader(self):
		return self.xml_header

	def get_Elements(self):
		return self.elements

	def get_xml(self):
		value = list()

		value.append(self.get_xmlHeader())
		value.append('<%s %s>' % (self.get_Tag(), ' '.join(self.attributes)))

		''' Si no hay ningún elemento se añade un párrafo en blanco '''
		if not self.get_Elements():
			pass
			'''value.append('%s<w:p w:rsidR="%s" w:rsidRDefault="%s">' % (self.get_tab(), self.rsidR, self.rsidRDefault))
			value.append('%s<w:pPr>' % self.get_tab(2))
			value.append('%s<w:pStyle w:val="Encabezado"/>' % self.get_tab(3))
			value.append('%s</w:pPr>' % self.get_tab(2))
			value.append('%s</w:p>' % self.get_tab())'''
		else:
			for element in self.get_Elements():
				value.append(element.get_xml())

		# value.append('<w:pPr> <w:pStyle w:val = "Encabezado"/> </w:pPr>')
		value.append('</%s>' % self.get_Tag())
		value.append('')

		return self.get_separator().join(value)

	def add_paragraph(self, text='', horizontal_alignment='', font_format='', font_size=None):
		self.get_parent().idx += 1
		idx = self.get_parent().idx
		paragraph = elements.paragraph.Paragraph(self, idx, text, horizontal_alignment, font_format,
													font_size)
		self.elements.append(paragraph)
		return paragraph

	def add_table(self, data=(), titulos=(), column_width=(), horizontal_alignment=None, borders=None):
		self.get_parent().idx += 1
		idx = self.get_parent().idx
		table = elements.table.Table(self, idx, data, titulos, column_width, horizontal_alignment, borders)
		self.elements.append(table)
		return table

	def add_rtf(self, text):
		_rtf = rtf.Rtf(text=text, parent=self)
		_rtf.get_value('word')

class Notes:
	"""
	classdocs
	"""

	def __init__(self, parent, part_type):
		self.tag = 'w:%s' % part_type
		self.name = 'word/%s.xml' % part_type
		self.content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.%s+xml" % part_type
		self.rId = 0
		self.tab = parent.tab
		self.separator = parent.get_separator()
		self.indent = 1
		self.parent = parent

	def get_RId(self):
		return self.rId

	def SetRId(self, value):
		self.rId = value

	def ContentType(self):
		return self.content_type

	def get_Tag(self):
		return self.tag

	def get_name(self):
		return self.name

	def get_separator(self):
		return self.separator

	def get_parent(self):
		return self.parent

	def get_xml(self):
		txt = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
		txt += '<%s xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"' % self.get_Tag()
		txt += ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
		txt += ' xmlns:o="urn:schemas-microsoft-com:office:office"'
		txt += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
		txt += ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
		txt += ' xmlns:v="urn:schemas-microsoft-com:vml"'
		txt += ' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'
		txt += ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
		txt += ' xmlns:w10="urn:schemas-microsoft-com:office:word"'
		txt += ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
		txt += ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
		txt += ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'
		txt += ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
		txt += ' xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"'
		txt += ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"'
		txt += ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
		txt += ' mc:Ignorable="w14 w15 wp14">'

		value = list()
		value.append(txt)
		value.append('<%s w:type="separator" w:id="-1">' % self.get_Tag()[:-1])

		value.append('<w:p w:rsidR="000E0AE3" w:rsidRDefault="000E0AE3" w:rsidP="00D4598A">')
		value.append('<w:pPr>')
		value.append('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
		value.append('</w:pPr>')
		value.append('<w:r>')
		value.append('<w:separator/>')
		value.append('</w:r>')
		value.append('</w:p>')
		value.append('</w:endnote>')
		value.append('<w:endnote w:type="continuationSeparator" w:id="0">')
		value.append('<w:p w:rsidR="000E0AE3" w:rsidRDefault="000E0AE3" w:rsidP="00D4598A">')
		value.append('<w:pPr>')
		value.append('<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
		value.append('</w:pPr>')
		value.append('<w:r>')
		value.append('<w:continuationSeparator/>')
		value.append('</w:r>')
		value.append('</w:p>')
		value.append('</%s>' % self.get_Tag()[:-1])
		value.append('</%s>' % self.get_Tag())
		value.append('')
		return self.separator.join(value)
