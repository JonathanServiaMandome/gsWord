#!/usr/bin/python
# -*- coding: utf-8 -*-


class WebSettings:
	def __init__(self, parent):
		self.name = 'word/webSettings.xml'
		self.tag = 'w:webSettings'
		self.content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
		self.rId = 3
		self.parent = parent
		self.separator = self.parent.separator
		self.tab = self.parent.tab

		self.optimizeForBrowser = ''
		self.allowPNG = ''
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if not self.separator:
			self.xml_header += '\n'

		self.xmlns = {
			'mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
			'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
			'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			'w14': "http://schemas.microsoft.com/office/word/2010/wordml",
			'w15': "http://schemas.microsoft.com/office/word/2012/wordml"}

		self.ignorable = 'w14 w15'

	def get_rid(self):
		return self.rId

	def set_rid(self, value):
		self.rId = value

	def ContentType(self):
		return self.content_type

	def get_AllowPNG(self):
		return self.allowPNG

	def SetAllowPNG(self, value):
		self.allowPNG = value

	def get_OptimizeForBrowser(self):
		return self.optimizeForBrowser

	def SetOptimizeForBrowser(self, value):
		self.optimizeForBrowser = value

	def get_xmlHeader(self):
		return self.xml_header

	def SetXMLHeader(self, value):
		self.xml_header = value

	def get_Ignorable(self):
		return self.ignorable

	def SetIgnorable(self, value):
		self.ignorable = value

	def get_xmlNS(self):
		return self.xmlns

	def get_xmlNSbyName(self, name):
		return self.xmlns.get(name, '')

	def SetXMLNS(self, dc):
		self.xmlns = dc

	def get_parent(self):
		return self.parent

	def set_tag(self):
		return self.tag

	def get_name(self):
		return self.name

	def get_xml(self):
		def Value_(name, value_, attr=''):
			t1, t2, t3 = ('', '', '/')
			if value_ != '':
				t1 = '>'
				t2 = value_
				t3 = '</%s' % name
			return '%s<%s%s%s%s%s>' % (self.tab, name, attr, t1, t2, t3)

		value = list()
		value.append(self.get_xmlHeader())
		value.append(
			'<%s xmlns:mc="%s" xmlns:r="%s" xmlns:w="%s" xmlns:w14="%s" xmlns:w15="%s" mc:Ignorable="%s">' % (
				self.set_tag(), self.get_xmlNSbyName('mc'), self.get_xmlNSbyName('r'), self.get_xmlNSbyName('w'),
				self.get_xmlNSbyName('w14'), self.get_xmlNSbyName('w15'), self.get_Ignorable()))
		value.append(Value_('w:optimizeForBrowser', self.get_OptimizeForBrowser()))
		value.append(Value_('w:allowPNG', self.get_AllowPNG()))
		value.append('</%s>' % (self.set_tag()))
		value.append('')
		return self.separator.join(value)
