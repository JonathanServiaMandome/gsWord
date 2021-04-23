#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Created on 15/03/2020

@author: Jonathan Servia Mandome
"""


class ContentType(object):
	def __init__(self, parent):
		self.name = '[Content_Types].xml'
		self.tag = 'Types'
		self.parent = parent
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		self.xmlns = 'http://schemas.openxmlformats.org/package/2006/content-types'
		self.default = {
			"rels": 'ContentType="application/vnd.openxmlformats-package.relationships+xml"',
			"xml": 'ContentType="application/xml"'}

		self.separator = self.parent.separator
		self.tab = self.parent.get_tab()

	def get_xmlHeader(self):
		return self.xml_header

	def SetXMLHeader(self, value):
		self.xml_header = value

	def get_xmlNS(self):
		return self.xmlns

	def SetXMLNS(self, value):
		self.xmlns = value

	def get_Default(self):
		return self.default

	def SetDefault(self, value):
		self.default = value

	def AddDefault(self, key, value):
		self.default[key] = value

	def get_parent(self):
		return self.parent

	def get_separator(self):
		return self.separator

	def get_tab(self):
		return self.tab

	def set_tag(self):
		return self.tag

	def get_name(self):
		return self.name

	def get_xml(self):
		value = [self.get_xmlHeader(), '<%s xmlns="%s">' % (self.set_tag(), self.get_xmlNS())]

		for target in self.get_Default().keys():
			value.append(
				self.tab + '<Default Extension="%s" %s/>' % (
					target, self.get_Default()[target]))

		for part_name in self.get_parent().get_parts():
			if part_name in ['rels', 'document_rels']:
				continue
			part = self.get_parent().get_part(part_name)
			if hasattr(part, 'ContentType'):
				value.append('%s<Override PartName="/%s" ContentType="%s"/>' % (
					self.get_tab(), part.get_name(), part.ContentType()))

		value.append('</%s>' % self.set_tag())
		value.append('')
		return self.separator.join(value)
