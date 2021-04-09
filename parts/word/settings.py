#!/usr/bin/python
# -*- coding: utf-8 -*-


class Settings:
	def __init__(self, parent):
		self.name = 'word/settings.xml'
		self.tag = 'settings'
		self.content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
		self.rId = 2
		self.parent = parent
		self.separator = self.parent.separator
		self.tab = self.parent.get_tab()
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if not self.separator:
			self.xml_header += '\n'
		self.xmlns = {
			'mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
			'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
			'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
			'w14': "http://schemas.microsoft.com/office/word/2010/wordml",
			'w15': "http://schemas.microsoft.com/office/word/2012/wordml"
		}
		self.ignorable = 'w14 w15'
		self.compat = {}
		self.rsids = []
		self.zoom = ''
		self.hyphenationZone = ''
		self.characterSpacingControl = ''
		self.mathFont = ''
		self.brkBin = ''
		self.brkBinSub = ''
		self.smallFrac = ''
		self.dispDef = ''
		self.lMargin = ''
		self.rMargin = ''
		self.defJc = ''
		self.wrapIndent = ''
		self.intLim = ''
		self.naryLim = ''
		self.themeFontLang = ''
		self.clrSchemeMapping = {}
		self.shapedefaults = {}
		self.shapelayout = ''
		self.idmap = {}
		self.decimalSymbol = ''
		self.listSeparator = ''
		self.chartTrackingRefBased = ''
		self.docId = ''

	def GetRId(self):
		return self.rId

	def SetRId(self, value):
		self.rId = value

	def ContentType(self):
		return self.content_type

	def GetZoom(self):
		return self.zoom

	def SetZoom(self, value):
		self.zoom = value

	def Get_rsids(self):
		return self.rsids

	def Set_rsids(self, value):
		self.rsids = value

	def AddRsid(self, value):
		self.rsids.append(value)

	def GetCharacterSpacingControl(self):
		return self.characterSpacingControl

	def SetCharacterSpacingControl(self, value):
		self.characterSpacingControl = value

	def GetHyphenationZone(self):
		return self.hyphenationZone

	def SetHyphenationZone(self, value):
		self.hyphenationZone = value

	def get_xmlHeader(self):
		return self.xml_header

	def SetXMLHeader(self, value):
		self.xml_header = value

	def GetIgnorable(self):
		return self.ignorable

	def SetIgnorable(self, value):
		self.ignorable = value

	def GetCompat(self):
		return self.compat

	def GetCompatbyName(self, name):
		return self.compat.get(name, '')

	def SetCompat(self, dc):
		self.compat = dc

	def get_xmlNS(self):
		return self.xmlns

	def get_xmlNSbyName(self, name):
		return self.xmlns.get(name, '')

	def SetXMLNS(self, dc):
		self.xmlns = dc

	def get_parent(self):
		return self.parent

	def GetTag(self):
		return self.tag

	def GetTag(self):
		return self.tag

	def get_name(self):
		return self.name

	def get_xml_(self):
		"""def Value_(tipo_, value_, attr=''):
			t1, t2, t3 = ('', '', '/')
			if value_ != '':
				t1 = '>'
				t2 = value_
				t3 = '</%s' % tipo_
			return '%s<w:%s%s%s%s%s>' % (self.tab, tipo_, attr, t1, t2, t3)"""

		value = list()
		value.append(self.get_xmlHeader())
		value.append(
			'<w:%s xmlns:mc="%s" xmlns:r="%s" xmlns:w="%s" xmlns:w14="%s" xmlns:w15="%s" mc:Ignorable="%s" >' % (
				self.GetTag(), self.get_xmlNSbyName('mc'), self.get_xmlNSbyName('r'), self.get_xmlNSbyName('w'),
				self.get_xmlNSbyName('w14'), self.get_xmlNSbyName('w15'), self.GetIgnorable()))
		# value.append(Value_('optimizeForBrowser', self.GetOptimizeForBrowser()))
		# value.append(Value_('allowPNG', self.GetAllowPNG()))
		value.append('/<w:%s>' % (self.GetTag()))
		return self.separator.join(value)

	def get_xml(self):
		return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main" mc:Ignorable="w14 w15">
	<w:zoom w:percent="110"/>
	<w:defaultTabStop w:val="708"/>
	<w:hyphenationZone w:val="425"/>
	<w:characterSpacingControl w:val="doNotCompress"/>
	<w:compat>
		<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
		<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
		<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
		<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
		<w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
	</w:compat>
	<w:rsids>
		<w:rsidRoot w:val="00455F4E"/>
		<w:rsid w:val="000F5F75"/>
		<w:rsid w:val="00583971"/>
		<w:rsid w:val="00575450"/>
	</w:rsids>
	<m:mathPr>
		<m:mathFont m:val="Cambria Math"/>
		<m:brkBin m:val="before"/>
		<m:brkBinSub m:val="--"/>
		<m:smallFrac m:val="0"/>
		<m:dispDef/>
		<m:lMargin m:val="0"/>
		<m:rMargin m:val="0"/>
		<m:defJc m:val="centerGroup"/>
		<m:wrapIndent m:val="1440"/>
		<m:intLim m:val="subSup"/>
		<m:naryLim m:val="undOvr"/>
	</m:mathPr>
	<w:themeFontLang w:val="es-ES"/>
	<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
	<w:shapeDefaults>
		<o:shapedefaults v:ext="edit" spidmax="1026"/>
		<o:shapelayout v:ext="edit">
			<o:idmap v:ext="edit" data="1"/>
		</o:shapelayout>
	</w:shapeDefaults>
	<w:decimalSymbol w:val=","/>
	<w:listSeparator w:val=";"/>
	<w15:chartTrackingRefBased/>
	<w15:docId w15:val="{0F016CD3-B598-420B-B9E3-A402EB2B080B}"/>
</w:settings>
"""
