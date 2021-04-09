#!/usr/bin/python
# -*- coding: utf-8 -*-

import font


class FontTable:
	def __init__(self, parent):
		self.name = 'word/fontTable.xml'
		self.tag = 'w:fonts'
		self.content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
		self.rId = 0
		self.parent = parent
		self.separator = self.parent.separator
		self.tab = self.parent.tab
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if not self.separator:
			self.xml_header += '\n'

		self.fonts = {'Calibri': font.Font(self.parent, 'Calibri'),
						'Times New Roman': font.Font(self.parent, 'Times New Roman'),
						'Calibri Light': font.Font(self.parent, 'Calibri Light')}

		self.xmlns = {
			'mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
			'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
			'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			'w14': "http://schemas.microsoft.com/office/word/2010/wordml",
			'w15': "http://schemas.microsoft.com/office/word/2012/wordml"}
		self.ignorable = 'w14 w15'

	def GetRId(self):
		return self.rId

	def SetRId(self, value):
		self.rId = value

	def ContentType(self):
		return self.content_type

	def get_xmlHeader(self):
		return self.xml_header

	def SetXMLHeader(self, value):
		self.xml_header = value

	def GetIgnorable(self):
		return self.ignorable

	def SetIgnorable(self, value):
		self.ignorable = value

	def get_xmlNS(self):
		return self.xmlns

	def get_xmlNSbyName(self, name):
		return self.xmlns.get(name, '')

	def SetXMLNS(self, dc):
		self.xmlns = dc

	def GetFonts(self):
		return self.fonts

	def GetFont(self, font_name):
		return self.fonts[font_name]

	def SetFonts(self, dc):
		self.fonts = dc

	def SetFont(self, font_):
		self.fonts[font_.GetFontName()] = font_

	def AddFont(self, font_):
		if font_.GetFontName() in self.GetFonts().keys():
			raise ValueError("La fuente %s ya está añadida al documento." % font_.GetFontName())
		self.fonts[font_.GetFontName()] = font_

	def AddFontbyName(self, name):
		_font = font.Font(self.parent, name)
		self.AddFont(_font)

	def get_parent(self):
		return self.parent

	def get_name(self):
		return self.name

	def GetTag(self):
		return self.tag

	def get_xml(self):
		value = [self.get_xmlHeader(),
					'<%s xmlns:mc="%s" xmlns:r="%s" xmlns:w="%s" xmlns:w14="%s" xmlns:w15="%s" mc:Ignorable="%s">' % (
					self.GetTag(), self.get_xmlNSbyName('mc'), self.get_xmlNSbyName('r'),
					self.get_xmlNSbyName('w'), self.get_xmlNSbyName('w14'),
					self.get_xmlNSbyName('w15'), self.GetIgnorable())]
		fonts = self.GetFonts()
		keys = fonts.keys()
		keys.sort()
		for _font in keys:
			value.append(self.GetFont(_font).get_xml())

		value.append('</%s>' % self.GetTag())
		value.append('')
		return self.separator.join(value)
