#!/usr/bin/python
# -*- coding: utf-8 -*-


class Rels(object):
	def __init__(self, parent):
		self.parent = parent
		self.name = '_rels/.rels'
		self.tag = 'Relationships'
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		self.xmlns = 'http://schemas.openxmlformats.org/package/2006/relationships'
		self.separator = self.parent.separator
		self.tab = self.parent.get_tab()
		self.parts = {}

	def GetParts(self):
		return self.parts

	def SetParts(self, dc):
		self.parts = dc

	def AddPart(self, name, value):
		if name in self.parts.keys():
			raise ValueError("La parte %s ya está añadida al documento." % name)
		self.parts[name] = value

	def GetPart(self, name):
		return self.parts[name]

	def SetPart(self, name, dc):
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

	def GetTag(self):
		return self.tag

	def get_name(self):
		return self.name

	def DefaultValues(self):
		ids = {"docProps/app.xml": {"Id": "rId3",
									"Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"},
				"docProps/core.xml": {"Id": "rId2",
										"Type": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"},
				"word/document.xml": {"Id": "rId1",
										"Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"}}
		self.SetParts(ids)

	def get_xml(self):
		value = [self.get_xmlHeader(), '<%s xmlns="%s">' % (self.GetTag(), self.get_xmlNS())]

		for target in self.GetParts():
			dc = self.GetPart(target)
			value.append(
				self.tab + '<Relationship Id="%s" Type="%s" Target="%s"/>' % (
					dc.get('Id', ''), dc.get('Type', ''), target))

		value.append('</%s>' % self.GetTag())
		value.append('')
		return self.separator.join(value)
