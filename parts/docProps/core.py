#!/usr/bin/python
# -*- coding: utf-8 -*-

from datetime import datetime


class Core(object):
	def __init__(self, parent):
		self.parent = parent
		self.name = 'docProps/core.xml'
		self.tag = 'cp:coreProperties'
		self.content_type = "application/vnd.openxmlformats-package.core-properties+xml"
		self.xml_header = ''
		self.xmlns_cp = ''
		self.xmlns_dc = ''
		self.dcterms = ''
		self.dcmitype = ''
		self.keywords = ''
		self.xsi = ''
		self.title = ''
		self.subject = ''
		self.creator = ''
		self.description = ''
		self.lastModifiedBy = ''
		self.revision = '1'
		self.lastPrinted = ''
		self.dcterms_created = ''
		self.dcterms_modified = ''
		self.separator = self.parent.separator
		self.tab = self.parent.get_tab()

	def ContentType(self):
		return self.content_type

	def GetXmlnsCp(self):
		return self.xmlns_cp

	def SetXmlnsCp(self, value):
		self.xmlns_cp = value

	def GetXmlnsDc(self):
		return self.xmlns_dc

	def SetXmlnsDc(self, value):
		self.xmlns_dc = value

	def GetDcterms(self):
		return self.dcterms

	def SetDcterms(self, value):
		self.dcterms = value

	def GetDcmitype(self):
		return self.dcmitype

	def SetDcmitype(self, value):
		self.dcmitype = value

	def GetTitle(self):
		return self.title

	def SetTitle(self, value):
		self.title = value

	def GetSubject(self):
		return self.subject

	def SetSubject(self, value):
		self.subject = value

	def GetCreator(self):
		return self.creator

	def SetCreator(self, value):
		self.creator = value

	def GetKeywords(self):
		return self.keywords

	def SetKeywords(self, value):
		self.keywords = value

	def GetDescription(self):
		return self.description

	def SetDescription(self, value):
		self.description = value

	def GetLastModifiedBy(self):
		return self.lastModifiedBy

	def SetLastModifiedBy(self, value):
		self.lastModifiedBy = value

	def GetRevision(self):
		return self.revision

	def SetRevision(self, value):
		self.revision = value

	def GetLastPrinted(self):
		return self.lastPrinted

	def SetLastPrinted(self, value):
		self.lastPrinted = value

	def SetDctermsCreated(self, value):
		self.dcterms_created = value

	def GetDctermsCreated(self):
		return self.dcterms_created

	def GetDctermsModified(self):
		return self.dcterms_modified

	def SetDctermsModified(self, value):
		self.dcterms_modified = value

	def GetXsi(self):
		return self.xsi

	def SetXsi(self, value):
		self.xsi = value

	def get_xmlHeader(self):
		return self.xml_header

	def SetXMLHeader(self, value):
		self.xml_header = value

	def get_parent(self):
		return self.parent

	def get_name(self):
		return self.name

	def GetTag(self):
		return self.tag

	def DefaultValues(self):
		self.SetXMLHeader('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
		self.SetXmlnsCp("http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
		self.SetXmlnsDc("http://purl.org/dc/elements/1.1/")
		self.SetDcterms("http://purl.org/dc/terms/")
		self.SetDcmitype("http://purl.org/dc/dcmitype/")
		self.SetXsi("http://www.w3.org/2001/XMLSchema-instance")
		self.SetCreator('Cocodin Technology')
		self.SetLastModifiedBy('Cocodin Technology')
		self.SetRevision('1')

	def get_xml(self):
		def Value_(tipo_, value_, attr=''):
			t1, t2, t3 = ('', '', '/')
			if value_ != '':
				t1 = '>'
				t2 = value_
				t3 = '</%s' % tipo_
			return '%s<%s%s%s%s%s>' % (self.tab, tipo_, attr, t1, t2, t3)

		value = list()
		value.append(self.get_xmlHeader())
		value.append('<%s xmlns:cp="%s" xmlns:dc="%s" xmlns:dcterms="%s" xmlns:dcmitype="%s" xmlns:xsi="%s">' % (
			self.GetTag(), self.GetXmlnsCp(), self.GetXmlnsDc(), self.GetDcterms(), self.GetDcmitype(), self.GetXsi()))

		value.append(Value_('dc:title', self.GetTitle()))
		value.append(Value_('dc:subject', self.GetSubject()))
		value.append(Value_('dc:creator', self.GetCreator()))
		value.append(Value_('cp:keywords', self.GetKeywords()))
		value.append(Value_('dc:description', self.GetDescription()))
		value.append(Value_('cp:lastModifiedBy', self.GetLastModifiedBy()))
		value.append(Value_('cp:revision', self.GetRevision()))
		if self.GetLastPrinted():
			value.append(Value_('cp:lastPrinted', self.GetLastPrinted()))
		now = datetime.now()
		value.append(Value_('dcterms:created', now.strftime("%Y-%m-%dT%H:%M:%SZ"), ' xsi:type="dcterms:W3CDTF"'))
		value.append(Value_('dcterms:modified', now.strftime("%Y-%m-%dT%H:%M:%SZ"), ' xsi:type="dcterms:W3CDTF"'))

		value.append('</%s>' % self.GetTag())
		value.append('')
		return self.separator.join(value)
