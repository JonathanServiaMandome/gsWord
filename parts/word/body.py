#!/usr/bin/python
# -*- coding: utf-8 -*-

import elements
import rtf


class Body(object):
	"""
	classdocs
	"""

	class Section:
		def __init__(self, parent, width=11906, height=16838, margin_top=1417, margin_rigth=1701, margin_left=1701,
						margin_bottom=1417, margin_header=708, margin_footer=708, margin_gutter=0, cols_space=708,
						line_pitch=360, orient=''):
			self.tag = 'w:sectPr'
			self.rsidR = '000F5F75'
			self.rsidRDefault = '000F5F75'
			self.rsidP = '000F5F75'
			self.rsidSect = ''
			self.tab = parent.tab
			self.indent = parent.indent + 1
			self.separator = parent.separator
			self.header_references = list()
			self.footer_references = list()
			self.width = width
			self.height = height
			self.margin_top = margin_top
			self.margin_rigth = margin_rigth
			self.margin_bottom = margin_bottom
			self.margin_left = margin_left
			self.margin_header = margin_header
			self.margin_footer = margin_footer
			self.margin_gutter = margin_gutter
			self.cols_space = cols_space
			self.line_pitch = line_pitch
			self.orient = orient
			if self.orient:
				self.SetHorizontalOrient()
			self.first = False
			self.page_break = False
			self.elements = list()

		def get_xml(self):
			value = list()
			if self.first:
				value.append(
					'%s<w:p w:rsidR="%s" w:rsidRDefault="%s" w:rsidP="%s">' % (self.get_tab(-2), self.get__rsidR(),
																				self.get__rsidRDefault(),
																				self.get__rsidRDefault()))
				value.append('%s<w:pPr>' % (self.get_tab(-1)))
			rsidsection = ''
			if self.get__rsidSect():
				rsidsection = ' w:rsidR="%s"' % self.get__rsidSect()
			value.append('%s<w:sectPr w:rsidR="%s%s">' % (self.get_tab(), self.get__rsidR(), rsidsection))
			for reference in self.get_HeaderReferences():
				value.append(
					'%s<w:headerReference w:type="%s" r:id="rId%d"/>' % (
						self.get_tab(1), reference[0], reference[1]))
			for reference in self.get_FooterReferences():
				value.append(
					'%s<w:footerReference w:type="%s" r:id="rId%d"/>' % (
						self.get_tab(1), reference[0], reference[1]))
			orient = ''
			if self.get_Orient():
				orient = ' w:orient="%s"' % self.get_Orient()
			value.append(
				'%s<w:pgSz w:w="%d" w:h="%d"%s/>' % (self.get_tab(1), self.get_width(), self.get_height(), orient))
			value.append('%s<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="%d" '
							'w:footer="%d" w:gutter="%d"/>' % (self.get_tab(1), self.get_MarginTop(), self.get_margin_rigth(),
															self.get_MarginBottom(), self.get_margin_left(),
															self.get_MarginHeader(),
															self.get_MarginFooter(), self.get_MarginGutter()))
			value.append('%s<w:cols w:space="%d"/>' % (self.get_tab(1), self.get_ColsSpace()))
			value.append('%s<w:docGrid w:linePitch="%d"/>' % (self.get_tab(1), self.get_LinePitch()))
			value.append('%s</w:sectPr>' % (self.get_tab()))
			if self.first:
				value.append('%s</w:pPr>' % (self.get_tab(-1)))
				value.append('%s</w:p>' % (self.get_tab(-2)))

			return self.get_separator().join(value)

		def AddPageBreak(self):
			self.page_break = True

		def get_Orient(self):
			return self.orient

		def SetHorizontalOrient(self):
			width = self.get_width()
			self.SetWidth(self.get_height())
			self.set_height(width)
			self.orient = 'landscape'

		def SetVerticalOrient(self):
			self.orient = ''

		def get_Elements(self):
			return self.elements

		def get__rsidSect(self):
			return self.rsidSect

		def Set_rsidSect(self, value):
			self.rsidSect = value

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_separator(self):
			return self.separator

		def get__rsidR(self):
			return self.rsidR

		def Set_rsidR(self, value):
			self.rsidR = value

		def get__rsidRDefault(self):
			return self.rsidRDefault

		def Set_rsidRDefault(self, value):
			self.rsidRDefault = value

		def get__rsidP(self):
			return self.rsidP

		def Set_rsidP(self, value):
			self.rsidP = value

		def get_HeaderReferences(self):
			return self.header_references

		def add_header_reference(self, type_reference, part):
			self.header_references.append([type_reference, part])

		def get_FooterReferences(self):
			return self.footer_references

		def add_footer_reference(self, type_reference, part):
			self.footer_references.append([type_reference, part])

		def get_LinePitch(self):
			return self.line_pitch

		def SetLinePitch(self, value):
			self.line_pitch = value

		def get_ColsSpace(self):
			return self.cols_space

		def SetColsSpace(self, value):
			self.cols_space = value

		def get_margin_rigth(self):
			return self.margin_rigth

		def SetMarginRigth(self, value):
			self.margin_rigth = value

		def get_MarginBottom(self):
			return self.margin_bottom

		def SetMarginBottom(self, value):
			self.margin_bottom = value

		def get_margin_left(self):
			return self.margin_left

		def SetMarginLeft(self, value):
			self.margin_left = value

		def get_MarginTop(self):
			return self.margin_top

		def SetMarginTop(self, value):
			self.margin_top = value

		def get_MarginHeader(self):
			return self.margin_header

		def SetMarginHeader(self, value):
			self.margin_header = value

		def get_MarginFooter(self):
			return self.margin_footer

		def SetMarginFooter(self, value):
			self.margin_footer = value

		def get_MarginGutter(self):
			return self.margin_gutter

		def SetMarginGutter(self, value):
			self.margin_gutter = value

		def SetMargins(self, dc):
			if 'bottom' in dc.keys():
				self.SetMarginBottom(dc['bottom'])
			elif 'all' in dc.keys():
				self.SetMarginBottom(dc['all'])

			if 'right' in dc.keys():
				self.SetMarginRigth(dc['right'])
			elif 'all' in dc.keys():
				self.SetMarginRigth(dc['all'])

			if 'left' in dc.keys():
				self.SetMarginLeft(dc['left'])
			elif 'all' in dc.keys():
				self.SetMarginLeft(dc['all'])

			if 'top' in dc.keys():
				self.SetMarginTop(dc['top'])
			elif 'all' in dc.keys():
				self.SetMarginTop(dc['all'])

			if 'header' in dc.keys():
				self.SetMarginHeader(dc['header'])
			elif 'all' in dc.keys():
				self.SetMarginHeader(dc['all'])

			if 'footer' in dc.keys():
				self.SetMarginFooter(dc['footer'])
			elif 'all' in dc.keys():
				self.SetMarginFooter(dc['all'])

			if 'gutter' in dc.keys():
				self.SetMarginGutter(dc['gutter'])
			elif 'all' in dc.keys():
				self.SetMarginGutter(dc['all'])

		def get_width(self):
			return self.width

		def SetWidth(self, value):
			self.width = value

		def get_height(self):
			return self.height

		def set_height(self, value):
			self.height = value

		def get_WidthBody(self):
			w = self.get_width()

			w -= self.get_margin_rigth()
			w -= self.get_margin_left()
			return w

	def __init__(self, parent, sections=True):
		self.tag = 'w:body'
		self.name = 'word/document.xml'
		self.content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = 1
		self.parent = parent
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if self.separator is '':
			self.xml_header += '\n'

		self.attributes = list()
		self.attributes.append('xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"')
		self.attributes.append('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"')
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
		self.attributes.append('xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"')
		self.attributes.append('xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"')
		self.attributes.append('xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"')
		self.attributes.append('xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
		self.attributes.append('mc:Ignorable="w14 w15 wp14"')

		self.rsidR = '00253112'
		self.rsidRDefault = '00253112'
		self.rsidSect = None
		self.active_section = None
		self.sections = list()
		'''if sections:
			self.sections = [self.AddPrincipalSection()]
			self.active_section = self.sections[-1]'''

	def get_XmlHeader(self):
		return self.xml_header

	def ContentType(self):
		return self.content_type

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def set_tag(self):
		return self.tag

	def get_name(self):
		return self.name

	def get_separator(self):
		return self.separator

	def get_parent(self):
		return self.parent

	def get_Sections(self):
		return self.sections

	def get_active_section(self):
		return self.active_section

	def get_xml(self):
		value = list()

		value.append(self.get_XmlHeader())
		value.append('<w:document %s>' % ' '.join(self.attributes))

		value.append('%s<%s>' % (self.get_tab(), self.set_tag()))

		for section in self.get_Sections():
			for element in section.get_Elements():
				value.append(element.get_xml())
			value.append(section.get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.set_tag()))
		value.append('</w:document>')
		value.append('')

		return self.get_separator().join(value)

	def add_paragraph(self, text, horizontal_alignment='l', font_format='', font_size=None, nulo=False):
		self.get_parent().idx += 1
		paragraph = elements.paragraph.Paragraph(self, self.get_parent().idx, text, horizontal_alignment, font_format,
												 font_size, nulo=nulo)
		self.get_active_section().elements.append(paragraph)
		return paragraph

	def add_table(self, data=(), titulos=(), column_width=(), horizontal_alignment=None, borders=None):
		self.get_parent().idx += 1
		table = elements.table.Table(self, self.get_parent().idx, data, titulos, column_width, horizontal_alignment, borders)
		self.get_active_section().elements.append(table)
		return table

	def add_section(self, width=11906, height=16838, margin_top=1417, margin_rigth=1701, margin_left=1701,
					margin_bottom=1417, margin_header=708, margin_footer=708, margin_gutter=0, cols_space=708,
					line_pitch=360, orient=''):
		section = self.Section(self, width, height, margin_top, margin_rigth, margin_left, margin_bottom, margin_header,
								margin_footer, margin_gutter, cols_space, line_pitch, orient)
		# self.active_section.PageBreak(page_break)
		if self.sections:
			self.sections[-1].first = True
		self.sections.append(section)
		self.active_section = section
		return section

	def AddPrincipalSection(self, width=11906, height=16838, margin_top=1417, margin_rigth=1701, margin_left=1701,
							margin_bottom=1417, margin_header=708, margin_footer=708, margin_gutter=0, cols_space=708,
							line_pitch=360, orient=''):
		section = self.Section(self, width, height, margin_top, margin_rigth, margin_left, margin_bottom, margin_header,
								margin_footer, margin_gutter, cols_space, line_pitch, orient)
		parts = self.get_parent().get_parts()

		for part_name in parts:

			if part_name.startswith('header'):
				part = self.get_parent().get_part(part_name)
				if hasattr(part, 'get_TypeReference'):
					section.add_header_reference(part.get_TypeReference(), part)

			elif part_name.startswith('footer') and part_name is not 'footernotes':
				part = self.get_parent().get_part(part_name)
				if hasattr(part, 'get_TypeReference'):
					section.add_footer_reference(part.get_TypeReference(), part)

		return section

	def add_rtf(self, text):
		_rtf = rtf.Rtf(text=text, parent=self)
		_rtf.get_value('word')


'''def add_table(self, idx, dimensiones, bordes=None):
	self.tabla_activa=idx
	#import Components 
	self.elementos[idx]=Table(idx, dimensiones, bordes)


def AddRow(self, data, dimensiones=None, font_format='', alineamiento_hor='l', font_size=11, bg_color='', 
cabecera=False, bloque=False): #from Components import Row self.elementos[self.tabla_activa].rows.append(Row(
self.elementos[self.tabla_activa],data, dimensiones, font_format, alineamiento_hor, font_size, bg_color, cabecera, 
bloque)) if cabecera: self.elementos[self.tabla_activa].header_row=Row(self.elementos[self.tabla_activa],data, 
dimensiones, font_format, alineamiento_hor, font_size, bg_color, cabecera, bloque) 

def AddBreakPage(self, idx):		
	self.elementos[idx]=PageBreak()
'''
