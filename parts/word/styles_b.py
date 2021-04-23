#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
Created on 15/03/2020

@author: Jonathan Servia Mandome
"""


class Style:
	def __init__(self, parent):
		self.name = 'word/styles.xml'
		self.tag = 'style'
		self.content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
		self.rId = 1
		self.parent = parent
		self.separator = self.parent.separator
		self.tab = self.parent.tab
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if not self.separator:
			self.xml_header += '\n'

	def get_rid(self):
		return self.rId

	def set_rid(self, value):
		self.rId = value

	def ContentType(self):
		return self.content_type

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

	# noinspection PyMethodMayBeStatic
	def get_xml(self):
		txt = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15">
	<w:docDefaults>
		<w:rPrDefault>
			<w:rPr>
				<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
				<w:sz w:val="18"/>
				<w:szCs w:val="19"/>
				<w:lang w:val="es-ES" w:eastAsia="en-US" w:bidi="ar-SA"/>
			</w:rPr>
		</w:rPrDefault>
		<w:pPrDefault>
			<w:pPr>
				<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
			</w:pPr>
		</w:pPrDefault>
	</w:docDefaults>
	<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="371">
		<w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
		<w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
		<w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 4" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 5" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 6" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 7" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 8" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="heading 9" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="index 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index 9" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 1" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 2" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 3" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 4" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 5" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 6" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 7" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 8" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toc 9" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Normal Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="footnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="annotation text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="header" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="footer" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="index heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="caption" w:semiHidden="1" w:uiPriority="35" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="table of figures" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="envelope address" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="envelope return" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="footnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="annotation reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="line number" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="page number" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="endnote reference" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="endnote text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="table of authorities" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="macro" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="toa heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Bullet 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Number 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/>
		<w:lsdException w:name="Closing" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Default Paragraph Font" w:semiHidden="1" w:uiPriority="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="List Continue 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Message Header" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Subtitle" w:uiPriority="11" w:qFormat="1"/>
		<w:lsdException w:name="Salutation" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Date" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text First Indent" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text First Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Note Heading" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text Indent 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Body Text Indent 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Block Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Hyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="FollowedHyperlink" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Strong" w:uiPriority="22" w:qFormat="1"/>
		<w:lsdException w:name="Emphasis" w:uiPriority="20" w:qFormat="1"/>
		<w:lsdException w:name="Document Map" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Plain Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="E-mail Signature" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Top of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Bottom of Form" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Normal (Web)" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Acronym" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Address" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Cite" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Code" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Definition" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Keyboard" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Preformatted" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Sample" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Typewriter" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="HTML Variable" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Normal Table" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="annotation subject" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="No List" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Outline List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Outline List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Outline List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Simple 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Simple 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Simple 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Classic 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Colorful 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Colorful 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Colorful 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Columns 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 4" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 5" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 6" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 7" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table List 8" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table 3D effects 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table 3D effects 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table 3D effects 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Contemporary" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Elegant" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Professional" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Subtle 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Subtle 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Web 1" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Web 2" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Web 3" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Balloon Text" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Table Grid" w:uiPriority="39"/>
		<w:lsdException w:name="Table Theme" w:semiHidden="1" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="Placeholder Text" w:semiHidden="1"/>
		<w:lsdException w:name="No Spacing" w:uiPriority="1" w:qFormat="1"/>
		<w:lsdException w:name="Light Shading" w:uiPriority="60"/>
		<w:lsdException w:name="Light List" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 1" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 1" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 1" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 1" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 1" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 1" w:uiPriority="65"/>
		<w:lsdException w:name="Revision" w:semiHidden="1"/>
		<w:lsdException w:name="List Paragraph" w:uiPriority="34" w:qFormat="1"/>
		<w:lsdException w:name="Quote" w:uiPriority="29" w:qFormat="1"/>
		<w:lsdException w:name="Intense Quote" w:uiPriority="30" w:qFormat="1"/>
		<w:lsdException w:name="Medium List 2 Accent 1" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 1" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 1" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 1" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 1" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 1" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 1" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 1" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 2" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 2" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 2" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 2" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 2" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 2" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 2" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 2" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 2" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 2" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 2" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 2" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 2" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 2" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 3" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 3" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 3" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 3" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 3" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 3" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 3" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 3" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 3" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 3" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 3" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 3" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 3" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 3" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 4" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 4" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 4" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 4" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 4" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 4" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 4" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 4" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 4" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 4" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 4" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 4" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 4" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 4" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 5" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 5" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 5" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 5" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 5" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 5" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 5" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 5" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 5" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 5" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 5" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 5" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 5" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 5" w:uiPriority="73"/>
		<w:lsdException w:name="Light Shading Accent 6" w:uiPriority="60"/>
		<w:lsdException w:name="Light List Accent 6" w:uiPriority="61"/>
		<w:lsdException w:name="Light Grid Accent 6" w:uiPriority="62"/>
		<w:lsdException w:name="Medium Shading 1 Accent 6" w:uiPriority="63"/>
		<w:lsdException w:name="Medium Shading 2 Accent 6" w:uiPriority="64"/>
		<w:lsdException w:name="Medium List 1 Accent 6" w:uiPriority="65"/>
		<w:lsdException w:name="Medium List 2 Accent 6" w:uiPriority="66"/>
		<w:lsdException w:name="Medium Grid 1 Accent 6" w:uiPriority="67"/>
		<w:lsdException w:name="Medium Grid 2 Accent 6" w:uiPriority="68"/>
		<w:lsdException w:name="Medium Grid 3 Accent 6" w:uiPriority="69"/>
		<w:lsdException w:name="Dark List Accent 6" w:uiPriority="70"/>
		<w:lsdException w:name="Colorful Shading Accent 6" w:uiPriority="71"/>
		<w:lsdException w:name="Colorful List Accent 6" w:uiPriority="72"/>
		<w:lsdException w:name="Colorful Grid Accent 6" w:uiPriority="73"/>
		<w:lsdException w:name="Subtle Emphasis" w:uiPriority="19" w:qFormat="1"/>
		<w:lsdException w:name="Intense Emphasis" w:uiPriority="21" w:qFormat="1"/>
		<w:lsdException w:name="Subtle Reference" w:uiPriority="31" w:qFormat="1"/>
		<w:lsdException w:name="Intense Reference" w:uiPriority="32" w:qFormat="1"/>
		<w:lsdException w:name="Book Title" w:uiPriority="33" w:qFormat="1"/>
		<w:lsdException w:name="Bibliography" w:semiHidden="1" w:uiPriority="37" w:unhideWhenUsed="1"/>
		<w:lsdException w:name="TOC Heading" w:semiHidden="1" w:uiPriority="39" w:unhideWhenUsed="1" w:qFormat="1"/>
		<w:lsdException w:name="Plain Table 1" w:uiPriority="41"/>
		<w:lsdException w:name="Plain Table 2" w:uiPriority="42"/>
		<w:lsdException w:name="Plain Table 3" w:uiPriority="43"/>
		<w:lsdException w:name="Plain Table 4" w:uiPriority="44"/>
		<w:lsdException w:name="Plain Table 5" w:uiPriority="45"/>
		<w:lsdException w:name="Grid Table Light" w:uiPriority="40"/>
		<w:lsdException w:name="Grid Table 1 Light" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 1" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 1" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 1" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 1" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 1" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 1" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 1" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 2" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 2" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 2" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 2" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 2" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 2" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 3" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 3" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 3" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 3" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 3" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 3" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 3" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 4" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 4" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 4" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 4" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 4" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 4" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 4" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 5" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 5" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 5" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 5" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 5" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 5" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 5" w:uiPriority="52"/>
		<w:lsdException w:name="Grid Table 1 Light Accent 6" w:uiPriority="46"/>
		<w:lsdException w:name="Grid Table 2 Accent 6" w:uiPriority="47"/>
		<w:lsdException w:name="Grid Table 3 Accent 6" w:uiPriority="48"/>
		<w:lsdException w:name="Grid Table 4 Accent 6" w:uiPriority="49"/>
		<w:lsdException w:name="Grid Table 5 Dark Accent 6" w:uiPriority="50"/>
		<w:lsdException w:name="Grid Table 6 Colorful Accent 6" w:uiPriority="51"/>
		<w:lsdException w:name="Grid Table 7 Colorful Accent 6" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 1" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 1" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 1" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 1" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 1" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 1" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 1" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 2" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 2" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 2" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 2" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 2" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 2" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 2" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 3" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 3" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 3" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 3" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 3" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 3" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 3" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 4" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 4" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 4" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 4" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 4" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 4" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 4" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 5" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 5" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 5" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 5" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 5" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 5" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 5" w:uiPriority="52"/>
		<w:lsdException w:name="List Table 1 Light Accent 6" w:uiPriority="46"/>
		<w:lsdException w:name="List Table 2 Accent 6" w:uiPriority="47"/>
		<w:lsdException w:name="List Table 3 Accent 6" w:uiPriority="48"/>
		<w:lsdException w:name="List Table 4 Accent 6" w:uiPriority="49"/>
		<w:lsdException w:name="List Table 5 Dark Accent 6" w:uiPriority="50"/>
		<w:lsdException w:name="List Table 6 Colorful Accent 6" w:uiPriority="51"/>
		<w:lsdException w:name="List Table 7 Colorful Accent 6" w:uiPriority="52"/>
	</w:latentStyles>
	<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
		<w:name w:val="Normal"/>
		<w:qFormat/>
	</w:style>
	<w:style w:type="character" w:default="1" w:styleId="Fuentedeprrafopredeter">
		<w:name w:val="Default Paragraph Font"/>
		<w:uiPriority w:val="1"/>
		<w:semiHidden/>
		<w:unhideWhenUsed/>
	</w:style>
	<w:style w:type="table" w:default="1" w:styleId="Tablanormal">
		<w:name w:val="Normal Table"/>
		<w:uiPriority w:val="99"/>
		<w:semiHidden/>
		<w:unhideWhenUsed/>
		<w:tblPr>
			<w:tblInd w:w="0" w:type="dxa"/>
			<w:tblCellMar>
				<w:top w:w="0" w:type="dxa"/>
				<w:left w:w="0" w:type="dxa"/>
				<w:bottom w:w="0" w:type="dxa"/>
				<w:right w:w="0" w:type="dxa"/>
			</w:tblCellMar>
		</w:tblPr>
	</w:style>
	<w:style w:type="numbering" w:default="1" w:styleId="Sinlista">
		<w:name w:val="No List"/>
		<w:uiPriority w:val="99"/>
		<w:semiHidden/>
		<w:unhideWhenUsed/>
	</w:style>
</w:styles>
"""
		return txt

	def get_xml_(self):
		value = list()
		value.append(self.get_xmlHeader())
		"""value.append(
			'<w:%s xmlns:mc="%s" xmlns:r="%s" xmlns:w="%s" xmlns:w14="%s" xmlns:w15="%s" mc:Ignorable="%s" >' % (
				self.set_tag(), self.get_xmlNSbyName('mc'), self.get_xmlNSbyName('r'),
				self.get_xmlNSbyName('w'), self.get_xmlNSbyName('w14'),
				self.get_xmlNSbyName('w15'), self.get_Ignorable()))"""

		value.append('</%s>' % self.set_tag())
		return self.separator.join(value)
