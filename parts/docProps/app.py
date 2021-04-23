#!/usr/bin/python
# -*- coding: utf-8 -*-


class App(object):
	def __init__(self, parent):
		self.xml_header = ''
		self.parent = parent
		self.name = 'docProps/app.xml'
		self.tag = 'Properties'
		self.content_type = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
		self.xmlns = ''
		self.xmlnsvt = ''
		self.template = ''
		self.total_time = 0
		self.words = 0
		self.pages = 0
		self.characters = 0
		self.application = ''
		self.doc_security = 0
		self.paragraphs = 0
		self.lines = 0
		self.app_version = ''
		self.hyperlinks_changed = ''
		self.shared_doc = ''
		self.links_up_to_date = ''
		self.company = ''
		self.scale_crop = ''
		self.characters_with_spaces = 0
		self.separator = self.parent.separator
		self.tab = self.parent.get_tab()

	def ContentType(self):
		return self.content_type

	def get_ScaleCrop(self):
		return self.scale_crop

	def SetScaleCrop(self, value):
		self.scale_crop = value

	def get_CharactersWithSpaces(self):
		return self.characters_with_spaces

	def SetCharactersWithSpaces(self, value):
		self.characters_with_spaces = value

	def get_Company(self):
		return self.company

	def SetCompany(self, value):
		self.company = value

	def get_LinksUpToDate(self):
		return self.links_up_to_date

	def SetLinksUpToDate(self, value):
		self.links_up_to_date = value

	def get_SharedDoc(self):
		return self.shared_doc

	def SetSharedDoc(self, value):
		self.shared_doc = value

	def get_HyperlinksChanged(self):
		return self.hyperlinks_changed

	def SetHyperlinksChanged(self, value):
		self.hyperlinks_changed = value

	def get_AppVersion(self):
		return self.app_version

	def SetAppVersion(self, value):
		self.app_version = value

	def get_Lines(self):
		return self.lines

	def SetLines(self, value):
		self.lines = value

	def get_paragraphs(self):
		return self.paragraphs

	def set_paragraphs(self, value):
		self.paragraphs = value

	def get_DocSecurity(self):
		return self.doc_security

	def SetDocSecurity(self, value):
		self.doc_security = value

	def get_Application(self):
		return self.application

	def SetApplication(self, value):
		self.application = value

	def get_Characters(self):
		return self.characters

	def SetCharacters(self, value):
		self.characters = value

	def get_TotalTime(self):
		return self.total_time

	def SetTotalTime(self, value):
		self.total_time = value

	def get_Pages(self):
		return self.pages

	def SetPages(self, value):
		self.pages = value

	def get_Words(self):
		return self.words

	def SetWords(self, value):
		self.words = value

	def get_xmlNS(self):
		return self.xmlns

	def SetXMLNS(self, value):
		self.xmlns = value

	def get_Template(self):
		return self.template

	def SetTemplate(self, value):
		self.template = value

	def get_xmlNSVT(self):
		return self.xmlnsvt

	def SetXMLNSVT(self, value):
		self.xmlnsvt = value

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

	def DefaultValues(self):
		self.SetXMLHeader('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
		self.SetXMLNS("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
		self.SetXMLNSVT("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
		self.SetTemplate('Normal.dotm')
		self.SetTotalTime(0)
		self.SetWords(0)
		self.SetPages(1)
		self.SetCharacters(0)
		self.SetCharactersWithSpaces(0)
		self.SetDocSecurity(0)
		self.SetApplication('Microsoft Office Word')
		self.SetScaleCrop('false')
		self.SetLinksUpToDate('false')
		self.SetSharedDoc('false')
		self.SetHyperlinksChanged('false')
		self.SetAppVersion('15.0000')

	def get_xml(self):
		value = list()
		value.append(self.get_xmlHeader())
		value.append('<%s xmlns="%s" xmlns:vt="%s">' % (self.set_tag(), self.get_xmlNS(), self.get_xmlNSVT()))

		value.append('%s<Template>%s</Template>' % (self.tab, self.get_Template()))
		value.append('%s<TotalTime>%d</TotalTime>' % (self.tab, self.get_TotalTime()))
		value.append('%s<Pages>%d</Pages>' % (self.tab, self.get_Pages()))
		value.append('%s<Words>%d</Words>' % (self.tab, self.get_Words()))
		value.append('%s<Characters>%d</Characters>' % (self.tab, self.get_Characters()))
		value.append('%s<Application>%s</Application>' % (self.tab, self.get_Application()))
		value.append('%s<DocSecurity>%d</DocSecurity>' % (self.tab, self.get_DocSecurity()))
		value.append('%s<Lines>%d</Lines>' % (self.tab, self.get_Lines()))
		value.append('%s<Paragraphs>%d</Paragraphs>' % (self.tab, self.get_paragraphs()))
		value.append('%s<ScaleCrop>%s</ScaleCrop>' % (self.tab, self.get_ScaleCrop()))

		if self.get_Company():
			value.append('%s<Company>%s</Company>' % (self.tab, self.get_Company()))
		else:
			value.append('%s<Company/>' % self.tab)
		value.append('%s<LinksUpToDate>%s</LinksUpToDate>' % (self.tab, self.get_LinksUpToDate()))
		value.append('%s<CharactersWithSpaces>%d</CharactersWithSpaces>' % (self.tab, self.get_CharactersWithSpaces()))
		value.append('%s<SharedDoc>%s</SharedDoc>' % (self.tab, self.get_SharedDoc()))
		value.append('%s<HyperlinksChanged>%s</HyperlinksChanged>' % (self.tab, self.get_HyperlinksChanged()))
		value.append('%s<AppVersion>%s</AppVersion>' % (self.tab, self.get_AppVersion()))

		value.append('</%s>' % self.set_tag())
		value.append('')
		return self.separator.join(value)
