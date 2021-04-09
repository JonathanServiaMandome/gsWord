#!/usr/bin/python
# -*- coding: utf-8 -*-

FONTS = {
	'Calibri': {
		'panose1': '020F0502020204030204',
		'charset': '00',
		'family': 'swiss',
		'pitch': 'variable',
		'usb0': 'E00002FF',
		'usb1': '4000ACFF',
		'usb2': '00000001',
		'usb3': '00000000',
		'csb0': '0000019F',
		'csb1': '00000000'
	},

	'Times New Roman': {
		'panose1': '02020603050405020304',
		'charset': '00',
		'family': 'roman',
		'pitch': 'variable',
		'usb0': 'E0002AFF',
		'usb1': 'C0007841',
		'usb2': '00000009',
		'usb3': '00000000',
		'csb0': '000001FF',
		'csb1': '00000000'
	},

	'Calibri Light': {
		'panose1': '020F0302020204030204',
		'charset': '00',
		'family': 'swiss',
		'pitch': 'variable',
		'usb0': 'A00002EF',
		'usb1': '4000207B',
		'usb2': '00000000',
		'usb3': '00000000',
		'csb0': '0000019F',
		'csb1': '00000000'
	},

	'Barcode EAN13': {
		'panose1': '04027200000000000000',
		'charset': '00',
		'family': 'decorative',
		'pitch': 'variable',
		'usb0': '00000003',
		'usb1': '00000000',
		'usb2': '00000000',
		'usb3': '00000000',
		'csb0': '00000881',
		'csb1': '00000000'
	}
}
DEFAULT_FONT = {
		'panose1': '04027200000000000000',
		'charset': '00',
		'family': 'decorative',
		'pitch': 'variable',
		'usb0': '00000003',
		'usb1': '00000000',
		'usb2': '00000000',
		'usb3': '00000000',
		'csb0': '00000881',
		'csb1': '00000000'
	}


class Font(object):
	def __init__(self, parent, font_name):
		self.tag = 'w:font'
		self.parent = parent
		self.font_name = font_name

		self.separator = self.parent.separator
		self.tab = self.parent.get_tab()
		self.indent = 1

	def GetFontName(self):
		return self.font_name

	def SetFontName(self, font_name):
		self.font_name = font_name

	def get_parent(self):
		return self.parent

	def GetTag(self):
		return self.tag

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ['%s<%s w:name="%s">' % (self.get_tab(), self.GetTag(), self.GetFontName())]

		font = FONTS.get(self.GetFontName(), DEFAULT_FONT)
		if font == 1:
			raise ValueError("La fuente %s no está soportada por la librería." % self.GetFontName())
		value.append('%s<w:panose1 w:val="%s"/>' % (self.get_tab(1), font['panose1']))
		value.append('%s<w:charset w:val="%s"/>' % (self.get_tab(1), font['charset']))
		value.append('%s<w:family w:val="%s"/>' % (self.get_tab(1), font['family']))
		value.append('%s<w:pitch w:val="%s"/>' % (self.get_tab(1), font['pitch']))
		value.append('%s<w:sig w:usb0="%s" w:usb1="%s" w:usb2="%s" w:usb3="%s" w:csb0="%s" w:csb1="%s"/>' % (
			self.get_tab(1), font['usb0'], font['usb1'], font['usb2'], font['usb3'], font['csb0'], font['csb1']))

		value.append('%s</%s>' % (self.get_tab(), self.GetTag()))
		return self.separator.join(value)
