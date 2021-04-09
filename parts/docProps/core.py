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

	def get_XmlnsCp(self):
		return self.xmlns_cp

	def SetXmlnsCp(self, value):
		self.xmlns_cp = value

	def get_XmlnsDc(self):
		return self.xmlns_dc

	def SetXmlnsDc(self, value):
		self.xmlns_dc = value

	def get_Dcterms(self):
		return self.dcterms

	def SetDcterms(self, value):
		self.dcterms = value

	def get_Dcmitype(self):
		return self.dcmitype

	def SetDcmitype(self, value):
		self.dcmitype = value

	def get_Title(self):
		return self.title

	def SetTitle(self, value):
		self.title = value

	def get_Subject(self):
		return self.subject

	def SetSubject(self, value):
		self.subject = value

	def get_Creator(self):
		return self.creator

	def SetCreator(self, value):
		self.creator = value

	def get_Keywords(self):
		return self.keywords

	def SetKeywords(self, value):
		self.keywords = value

	def get_Description(self):
		return self.description

	def SetDescription(self, value):
		self.description = value

	def get_LastModifiedBy(self):
		return self.lastModifiedBy

	def SetLastModifiedBy(self, value):
		self.lastModifiedBy = value

	def get_Revision(self):
		return self.revision

	def SetRevision(self, value):
		self.revision = value

	def get_LastPrinted(self):
		return self.lastPrinted

	def SetLastPrinted(self, value):
		self.lastPrinted = value

	def SetDctermsCreated(self, value):
		self.dcterms_created = value

	def get_DctermsCreated(self):
		return self.dcterms_created

	def get_DctermsModified(self):
		return self.dcterms_modified

	def SetDctermsModified(self, value):
		self.dcterms_modified = value

	def get_Xsi(self):
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

	def get_Tag(self):
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
			self.get_Tag(), self.get_XmlnsCp(), self.get_XmlnsDc(), self.get_Dcterms(), self.get_Dcmitype(), self.get_Xsi()))

		value.append(Value_('dc:title', self.get_Title()))
		value.append(Value_('dc:subject', self.get_Subject()))
		value.append(Value_('dc:creator', self.get_Creator()))
		value.append(Value_('cp:keywords', self.get_Keywords()))
		value.append(Value_('dc:description', self.get_Description()))
		value.append(Value_('cp:lastModifiedBy', self.get_LastModifiedBy()))
		value.append(Value_('cp:revision', self.get_Revision()))
		if self.get_LastPrinted():
			value.append(Value_('cp:lastPrinted', self.get_LastPrinted()))
		now = datetime.now()
		value.append(Value_('dcterms:created', now.strftime("%Y-%m-%dT%H:%M:%SZ"), ' xsi:type="dcterms:W3CDTF"'))
		value.append(Value_('dcterms:modified', now.strftime("%Y-%m-%dT%H:%M:%SZ"), ' xsi:type="dcterms:W3CDTF"'))

		value.append('</%s>' % self.get_Tag())
		value.append('')
		return self.separator.join(value)
