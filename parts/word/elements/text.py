#!/usr/bin/python
# -*- coding: utf-8 -*-


class Properties(object):
	class Color(object):
		def __init__(self, parent, color='', theme_color='', theme_shade='', theme_tint=''):
			self.name = 'color'
			self.val = color
			self.themeColor = theme_color
			self.themeShade = theme_shade
			self.themeTint = theme_tint
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def setValue(self, value):
			self.val = value

		def setThemeColor(self, value):
			self.themeColor = value

		def setThemeShade(self, value):
			self.themeShade = value

		def setThemeTint(self, value):
			self.themeTint = value

		def get_Value(self):
			return self.val

		def get_name(self):
			return self.name

		def get_ThemeColor(self):
			return self.themeColor

		def get_ThemeShade(self):
			return self.themeShade

		def get_ThemeTint(self):
			return self.themeTint

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = ''
			if self.val:
				value += ' w:val="%s"' % self.val
			if self.themeColor:
				value += ' w:themeColor="%s"' % self.themeColor
			if self.themeShade:
				value += ' w:themeShade="%s"' % self.themeShade
			if self.themeTint:
				value += ' w:themeTint="%s"' % self.themeTint
			if value:
				value = '%s<w:%s %s />' % (self.get_tab(), self.name, value)
			return value

	class U(object):
		def __init__(self, parent, value, color=''):
			self.name = 'u'
			self.val = value
			self.color = color
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def SetValue(self, value):
			self.val = value

		def set_color(self, value):
			self.color = value

		def get_name(self):
			return self.name

		def get_color(self):
			return self.color

		def get_Value(self):
			return self.val

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = '<w:%s w:val="%s"' % (self.name, self.val)
			if self.color:
				value += ' w:color="%s"' % self.color
			value += '/>'
			return '%s%s' % (self.get_tab(), value)

	def __init__(self, parent, font_format='', font_size=None, tabulation=False):
		self.name = 'rPr'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.tabulation = tabulation

		self.bold = 'b' in font_format
		self.italic = 'i' in font_format
		self.underline = None
		if 'u' in font_format:
			self.underline = self.U(self, 'single')
		self.boldCs = False
		self.italicCs = False
		self.underlineCs = False

		'''Specifies that the content should be displayed with an underline: <w:u w:val="double"/>.
		The most common attributes are below (the theme-related attributed are omitted):

			color - specifies the color for the underlining. Values are given as hex values (in RRGGBB format). 
			No #, unlike hex values in HTML/CSS. E.g., color="FFFF00". A value of auto is also permitted and will 
			allow the consuming word processor to determine the color based on the context.
			val - Specifies the pattern to be used to create the underline. Possible values are:
				dash - a dashed line
				dashDotDotHeavy - a series of thick dash, dot, dot characters
				dashDotHeavy - a series of thick dash, dot characters
				dashedHeavy - a series of thick dashes
				dashLong - a series of long dashed characters
				dashLongHeavy - a series of thick, long, dashed characters
				dotDash - a series of dash, dot characters
				dotDotDash - a series of dash, dot, dot characters
				dotted - a series of dot characters
				dottedHeavy - a series of thick dot characters
				double - two lines
				none - no underline
				single - a single line
				thick - a single think line
				wave - a single wavy line
				wavyDouble - a pair of wavy lines
				wavyHeavy - a single thick wavy line
				words - a single line beneath all non-space characters'''

		'''Specifies that any lowercase characters are to be displayed as their uppercase equivalents. It cannot 
		appear with smallCaps in the same text run: <w:caps w:val="true" />. This is a toggle property. '''
		self.caps = False

		'''Specifies the color to be used to display text: <w:color w:val="FFFF00" /> Possible attributes are 
		themeColor, themeShade, themeTint, and val. The val attribute specifies the color as a hex value in RRGGBB 
		format, or auto may be specified to allow the consuming software to determine the color. '''
		self.color = None

		'''Specifies that the contents are to be displayed with two horizontal lines through each character: 
		<w:dstrike w:val="true"/>. It cannot appear with strike in the same text run. This is a toggle property. '''
		self.dstrike = False

		'''Specifies that the content should be displayed as if it were embossed, making text appear as if it is 
		raised off of the page: <w:emboss w:val="true" />. This is a toggle property. '''
		self.emboss = False

		'''Specifies that the content should be displayed as it it were imprinted (or engraved) into the page: 
		<w:imprint w:val="true"/>. This may not be present with either emboss or outline. This is a toggle property. '''
		self.imprint = False

		'''Specifies that the content should be displayed as if it had an outline. A one-pixel border is drawn 
		around the inside and outside borders of each character. <outline w:val="true"/>. This is a toggle property.'''
		self.outline = False

		'''Specifies the style ID of the character style to be used to format the contents of the run. Texto libre'''
		self.rStyle = ''

		'''Specifies that the content should be displayed as if each character has a shadow: <w:shadow w:val="true"/>. 
		For left-to-right text, the shadow is beneath the text and to its right. shadow may not be present with 
		either emboss or imprint. This is a toggle property.'''
		self.shadow = False

		'''Specifies that any lowercase characters are to be displayed as their uppercase equivalents in a font 
		size two points smaller than the specified font size: <w:smallCaps w:val="true"/>.
		It cannot appear with caps in the same text run. This is a toggle property. '''
		self.smallCaps = False

		'''Specifies that the contents are to be displayed with a horizontal line through the center of the line:
			<w:strike w:val="true"/>. It cannot appear with dstrike in the same text run. This is a toggle property. '''
		self.strike = False

		'''Specifies the font size in half points: <w:sz w:val="28"/>. Note that szCs is used for complex script 
		fonts such as Arabic.
		The single attribute val specifies a measurement in half-points (1/144 of an inch). '''
		self.sz = ''
		self.szCs = ''

		'''Specifies that the content is to be hidden from display at display time. <vanish/>. 
		This is a toggle property. '''
		self.vanish = False

		'''Subscript and superscript. <vertAlign w:val="superscript"/>.
		The single attribute is val. Permitted values are:			
			baseline - regular vertical positioning
			subscript - lowers the text below the baseline and changes it to a small size
			superscript - raises the text above the baseline and changes it to a smaller size'''
		self.vertAlign = ''

		self.rFonts = ''

		if font_size is not None:
			self.set_font_size(font_size)

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def set_font_format(self, font_format):
		self.bold = 'b' in font_format
		self.italic = 'i' in font_format
		if 'u' in font_format:
			self.underline = self.U(self, 'single')

	def set_font_format_cs(self, font_format):
		self.boldCs = 'b' in font_format
		self.italicCs = 'i' in font_format
		if 'u' in font_format:
			self.underlineCs = self.U(self, 'single')

	def get_RFonts(self):
		return self.rFonts

	def set_font(self, name):
		self.rFonts = name

	def get_separator(self):
		return self.separator

	def bold(self, value=True):
		self.bold = value

	def bold_cs(self, value=True):
		self.boldCs = value

	def italic(self, value=True):
		self.italic = value

	def italic_cs(self, value=True):
		self.italicCs = value

	def underline(self, value='single', color=''):
		self.underline = self.U(self, value, color)

	def underline_cs(self, value='single', color=''):
		self.underlineCs = self.U(self, value, color)

	def set_caps(self, value):
		self.caps = value

	def set_color(self, value):
		self.color = value

	def set_dstrike(self, value):
		self.dstrike = value

	def set_emboss(self, value):
		self.emboss = value

	def set_imprint(self, value):
		self.imprint = value

	def set_outline(self, value):
		self.outline = value

	def set_rstyle(self, value):
		self.rStyle = value

	def set_shadow(self, value):
		self.shadow = value

	def set_small_caps(self, value):
		self.smallCaps = value

	def set_strike(self, value):
		self.strike = value

	def set_font_size(self, value):
		value = str(int(value * 2))
		self.sz = value

	def set_font_size_cs(self, value):
		value = str(int(value * 2))
		self.szCs = value

	def set_vanish(self, value):
		self.vanish = value

	def set_vertical_align(self, value):
		self.vertAlign = value

	def get_caps(self):
		return self.caps

	def get_color(self):
		return self.color

	def get_dstrike(self):
		return self.dstrike

	def get_emboss(self):
		return self.emboss

	def get_imprint(self):
		return self.imprint

	def get_outline(self):
		return self.outline

	def get_rstyle(self):
		return self.rStyle

	def get_shadow(self):
		return self.shadow

	def get_small_caps(self):
		return self.smallCaps

	def get_strike(self):
		return self.strike

	def get_font_size(self):
		return self.sz

	def get_font_size_cs(self):
		return self.szCs

	def get_vanish(self):
		return self.vanish

	def get_vertical_align(self):
		return self.vertAlign

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		value.append('%s<w:%s>' % (self.get_tab(), self.name))
		if self.rFonts:
			value.append('%s<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (
				self.get_tab(1), self.rFonts, self.rFonts, self.rFonts))
		if self.bold:
			value.append('%s<w:b/>' % (self.get_tab(1)))
		if self.boldCs:
			value.append('%s<w:bCs/>' % (self.get_tab(1)))
		if self.italic:
			value.append('%s<w:i/>' % (self.get_tab(1)))
		if self.underline is not None:
			value.append(self.underline.get_xml())
		if self.caps:
			value.append('%s<w:caps w:val="true" />' % (self.get_tab(1)))
		if self.color is not None:
			value.append('%s<w:color w:val="%s" />' % (self.get_tab(1), self.get_color()))
		if self.dstrike:
			value.append('%s<w:dstrike w:val="true" />' % (self.get_tab(1)))
		if self.emboss:
			value.append('%s<w:emboss w:val="true" />' % (self.get_tab(1)))
		if self.imprint:
			value.append('%s<w:imprint w:val="true" />' % (self.get_tab(1)))
		if self.outline:
			value.append('%s<w:outline w:val="true" />' % (self.get_tab(1)))
		if self.rStyle:
			value.append('%s<w:rStyle w:val="%s" />' % (self.get_tab(1), self.rStyle))
		if self.shadow:
			value.append('%s<w:shadow w:val="true" />' % (self.get_tab(1)))
		if self.smallCaps:
			value.append('%s<w:smallCaps w:val="true" />' % (self.get_tab(1)))
		if self.strike:
			value.append('%s<w:strike w:val="true" />' % (self.get_tab(1)))
		if self.sz:
			value.append('%s<w:sz w:val="%s" />' % (self.get_tab(1), self.sz))
		if self.szCs:
			value.append('%s<w:szCs w:val="%s" />' % (self.get_tab(1), self.szCs))
		if self.vanish:
			value.append('%s<w:vanish/>' % (self.get_tab(1)))
		if self.vertAlign:
			value.append('%s<w:vertAlign w:val="%s" />' % (self.get_tab(1), self.vertAlign))

		# Si no tiene ninguna propiedad este elemento no es necesario
		if len(value) == 1:
			value = []
		else:
			value.append('%s</w:%s>' % (self.get_tab(), self.name))

			if self.tabulation:
				value.append('%s<w:tab/>' % (self.get_tab()))
		if value:
			return self.get_separator().join(value)
		else:
			return ''


class Text(object):

	def __init__(self, parent, text='', font_format='', font_size=None, tabulation=False, space="preserve"):
		self.name = 'r'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.rsidRPr = ''

		if type(text) in [str, unicode]:
			text = [text]
		self.text = []
		for txt in text:
			if txt:
				self.text.append(self.T(self, txt, tabulation, space))

		self.properties = Properties(self, font_format, font_size, tabulation)

		'''Para la númeración de páginas'''
		self.field_char = ''
		'''
		PAGE: Número de página actual
		NUMPAGE: Total de páginas
		'''
		self.instrText = ''

	def set_field_char(self, value):
		self.field_char = value

	def get_field_char(self):
		return self.field_char

	def set_instr_text(self, value):
		self.instrText = value

	def get_instr_text(self):
		return self.instrText

	def add_text(self, text):
		self.text.append(self.T(self, text))

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def set_font(self, name):
		self.get_properties().set_font(name)

	def get_separator(self):
		return self.separator

	def set_rsid_rpr(self, value):
		self.rsidRPr = value

	def get_rsid_rpr(self):
		return self.rsidRPr

	def get_properties(self):
		return self.properties

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()

		if self.rsidRPr == '':
			value.append('%s<w:%s>' % (self.get_tab(), self.name))
		else:
			value.append('%s<w:%s w:rsidRPr="%s">' % (self.get_tab(), self.name, self.rsidRPr))

		pr = self.properties.get_xml()
		if pr:
			value.append(pr)
		if self.get_field_char():
			value.append('%s<w:fldChar w:fldCharType="%s"/>' % (self.get_tab(1), self.get_field_char()))
		elif self.get_instr_text():
			value.append('%s<w:instrText>%s</w:instrText>' % (self.get_tab(1), self.get_instr_text()))
		else:
			for txt in self.text:
				tx = txt.get_xml()
				if not tx:
					continue
				value.append(tx)
		value.append('%s</w:%s>' % (self.get_tab(), self.name))

		return self.get_separator().join(value)

	class Tabs(object):
		def __init__(self, parent, tabs):
			self.name = 'tabs'
			self.tabs = tabs
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def SetTabs(self, value):
			self.tabs = value

		def get_tabs(self):
			return self.tabs

		def get_name(self):
			return self.name

		def get_separator(self):
			return self.separator

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = list()
			value.append('%s<w:%s>' % (self.get_tab(), self.name))
			for tab in self.tabs:
				value.append(tab.get_xml())
			value.append('%s</w:%s>' % (self.get_tab(), self.name))
			return self.get_separator().join(value)

	class Tab(object):
		def __init__(self, parent, value=None, posicion=None):
			self.name = 'tab'
			self.val = value
			self.pos = posicion
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def SetValue(self, value):
			self.val = value

		def SetPosition(self, value):
			self.pos = value

		def get_name(self):
			return self.name

		def get_Value(self):
			return self.val

		def get_Position(self):
			return self.pos

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			if self.val is None and self.pos is None:
				value = '%s<%s />' % (self.get_tab(), self.name)
			else:
				value = '%s<w:%s w:val="%s" w:pos="%s" />' % (self.get_tab(), self.name, self.val, self.pos)
			return value

	class Br(object):
		def __init__(self, parent, type_='', clear=''):
			self.name = 'br'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

			'''Specifies the type of break, which determines the location for the following text. Possible values are:
				column - Text restarts on the next column. If there are no columns or the current position is the last 
				column, then the restart location is the next page.
				page - Text restarts on the next page.
				textWrapping - Text restarts on the next line. This is the default value.'''
			self.type = type_

			'''Specifies the location where text restarts when the value of the type attribute is textWrapping.
			Note: This property only affects the restart location when the current run is being displayed on a line 
			which does not span the full text extents due to a floating object. Possible values are:			
				all - Text restarts on the next full line. This is often used for captions for objects.
				left - Text restarts text in the next text region left to right.
				none - Text restarts on the next line regardless of any floating objects. This is the typical break.
				right - Text restarts text in the next text region right to left.'''
			self.clear = clear

		def set_type(self, value):
			self.type = value

		def SetClear(self, value):
			self.clear = value

		def get_name(self):
			return self.name

		def get_type(self):
			return self.type

		def get_Clear(self):
			return self.clear

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = '%s<%s' % (self.get_tab(), self.name)
			if self.type:
				value += ' w:type="%s"' % self.type
			if self.clear:
				value += ' w:clear="%s"' % self.clear
			value += ' />'
			return value

	class Cr(object):
		def __init__(self, parent):
			self.name = 'cr'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def get_name(self):
			return self.name

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = '%s<%s />' % (self.get_tab(), self.name)
			return value

	class Drawing(object):
		def __init__(self, parent):
			self.name = 'drawing'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1
			self.val = ''

		def get_name(self):
			return self.name

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			return self.val

	class NoBreakHyphen(object):
		def __init__(self, parent):
			self.name = 'noBreakHyphen'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def get_name(self):
			return self.name

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = '%s<%s />' % (self.get_tab(), self.name)
			return value

	class SoftHyphen(object):
		def __init__(self, parent):
			self.name = 'softHyphen'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def get_name(self):
			return self.name

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = '%s<%s />' % (self.get_tab(), self.name)
			return value

	class Sym(object):
		def __init__(self, parent, font, char):
			self.name = 'sym'
			self.font = font
			self.char = char
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

		def SetFont(self, value):
			self.font = value

		def SetChar(self, value):
			self.char = value

		def get_Font(self):
			return self.font

		def get_Char(self):
			return self.char

		def get_name(self):
			return self.name

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = '%s<%s w:font="%s w:char="%s" />' % (self.get_tab(), self.name, self.font, self.char)
			return value

	class T(object):
		def __init__(self, parent, text, tabulation=False, space='preserve'):
			self.name = 't'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1

			if type(text) is not unicode:
				text = text.decode('iso-8859-1').encode('utf8').replace('%EURO%', '€').replace('&', '&amp;').replace(
					'<', '&lt;').replace('>', '&gt;')

			self.text = text

			'''True: tabulación, no incluye el texo, False:Texto'''
			self.tabulation = tabulation

			'''Specifies literal text which will be displayed in the run. Most all text within a document is contained 
			within t elements, except text within a field code.
			There is one possible attribute (xml:space) within the XML namespace that can be used to specify how space 
			within the t element is to be handled. 
			Possible values are preserve and default.
			If whitespace within a run needs to be preserved, it is important that this attribute be set to preserve. 
			See the XML 1.0 specification at � 2.10.'''

			self.space = space

		def SetText(self, text):
			self.text = text.decode('latin-1').encode('utf-8').replace('%EURO%', '€').replace('&', '&amp;').replace(
				'<', '&lt;').replace('>', '&gt;')

		def get_Text(self):
			return self.text

		def get_name(self):
			return self.name

		def IsTabulation(self):
			return self.tabulation

		def SetSpace(self, value):
			self.space = value

		def get_Space(self):
			return self.space

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			space = ''
			spaces = self.text.startswith(' ') or self.text.endswith(' ')
			if self.space and spaces:
				space = ' xml:space="%s"' % self.space
			value = '%s<w:%s%s>%s</w:%s>' % (self.get_tab(), self.name, space, self.text, self.name)

			if self.tabulation:
				value = '%s<w:tab/>' % self.get_tab()
			return value

