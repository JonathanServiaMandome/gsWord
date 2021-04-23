#!/usr/bin/python
# -*- coding: utf-8 -*-


class DocumentRels(object):
	def __init__(self, parent):
		self.parent = parent
		self.name = 'word/_rels/document.xml.rels'
		self.tag = 'Relationships'
		self.separator = self.parent.separator
		self.tab = self.parent.tab
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if self.separator is '':
			self.xml_header += '\n'
		self.xmlns = 'http://schemas.openxmlformats.org/package/2006/relationships'
		self.parts = dict()
		self.images = list()

	def get_images(self):
		return self.images

	def add_image(self, target, image):
		self.images.append([target, image])

	def get_parts(self):
		return self.parts

	def set_parts(self, dc):
		self.parts = dc

	def add_part(self, name, value):
		if name in self.parts.keys():
			raise ValueError("La parte %s ya está añadida al documento." % name)
		self.parts[name] = value

	def get_part(self, name):
		return self.parts[name]

	def set_part(self, name, dc):
		self.parts[name] = dc

	def get_xmlNS(self):
		return self.xmlns

	def SetXMLNS(self, value):
		self.xmlns = value

	def get_xmlHeader(self):
		return self.xml_header

	def SetXMLHeader(self, value):
		self.xml_header = value

	def get_parent(self):
		return self.parent

	def get_name(self):
		return self.name

	def set_tag(self):
		return self.tag

	def get_xml(self):
		value = [self.get_xmlHeader(), '<%s xmlns="%s">' % (self.set_tag(), self.get_xmlNS())]

		for target in self.get_parts():
			dc = self.get_part(target)
			value.append(
				self.tab + '<Relationship Id="%s" Type="%s" Target="%s"/>' % (
					dc.get('Id', ''), dc.get('Type', ''), target))

		value.append('</%s>' % self.set_tag())
		value.append('')
		return self.separator.join(value)
