#!/usr/bin/python
# -*- coding: utf-8 -*-

import element
import text
import pictures


class ImagePart:
	def __init__(self, name):
		self.name = name
		self.rId = 0

	def set_rid(self, value):
		self.rId = value

	def get_rid(self):
		return self.rId

	def get_name(self):
		return self.name


class Properties(object):
	class FormatList(object):
		def __init__(self, parent, level='1', _type='1'):
			self.name = 'numPr'
			self.parent = parent
			self.tab = self.parent.tab
			self.indent = self.parent.indent + 1
			self.level = level
			self.type = _type

		def get_tab(self, number=0):
			return self.tab * (number + self.indent)

		def get_parent(self):
			return self.parent

		def get_name(self):
			return self.name

		def get_Level(self):
			return self.level

		def get_type(self):
			return self.type

		def SetLevel(self, level):
			self.level = level

		def set_type(self, _type):
			self.type = _type

		def get_xml(self):
			value = list()
			value.append('%s<w:%s>' % (self.get_tab(), self.name))
			value.append('%s <w:ilvl w:val="%s"/>' % (self.get_tab(1), self.get_Level()))
			value.append('%s <w:numId w:val="%s"/>' % (self.get_tab(1), self.get_type()))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))

			return self.parent.separator.join(value)

	class Spacing(object):
		def __init__(self, parent, dc):
			self.name = 'spacing'
			self.before = dc.get('before', '')
			self.after = dc.get('after', '')
			self.line = dc.get('line', '')
			self.lineRule = dc.get('lineRule', '')
			self.before_autospacing = dc.get('beforeAutospacing', '')
			self.after_autospacing = dc.get('afterAutospacing', '')
			self.parent = parent
			self.tab = self.parent.tab
			self.indent = self.parent.indent + 1

		def get_tab(self, number=0):
			return self.tab * (number + self.indent)

		def set_after(self, value):
			self.after = value

		def set_before(self, value):
			self.before = value

		def set_line(self, value):
			self.line = value

		def set_line_rule(self, value):
			self.lineRule = value

		def set_before_autospacing(self, value):
			self.before_autospacing = value

		def set_after_autospacing(self, value):
			self.after_autospacing = value

		def get_after(self):
			return self.after

		def get_before(self):
			return self.before

		def get_line(self):
			return self.line

		def get_line_rule(self):
			return self.lineRule

		def get_before_autospacing(self):
			return self.before_autospacing

		def get_after_autospacing(self):
			return self.after_autospacing

		def get_parent(self):
			return self.parent

		def get_xml(self):
			value = ''
			if self.before:
				value += ' w:before="%s"' % self.before
			if self.after:
				value += ' w:after="%s"' % self.after
			if self.line:
				value += ' w:line="%s"' % self.line
			if self.lineRule:
				value += ' w:lineRule="%s"' % self.lineRule
			if self.before_autospacing:
				value += ' w:beforeAutospacing="%s"' % self.before_autospacing
			if self.after_autospacing:
				value += ' w:afterAutospacing="%s"' % self.after_autospacing
			if value:
				value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
			return value

	class Indentation(object):
		def __init__(self, parent, dc):
			self.name = 'ind'
			'''Specifies the indentation to be placed at the left (for paragraphs going left to right).'''
			self.left = dc.get('left', '')
			'''Specifies the indentation to be placed at the right (for paragraphs going left to right). '''
			self.right = dc.get('right', '')
			'''Specifies indentation to be removed from the first line. This attribute and firstLine are mutually 
			exclusive. This attribute controls when both are specified. '''
			self.hanging = dc.get('hanging', '')
			'''Specifies additional indentation to be applied to the first line. This attribute and hanging are 
			mutually exclusive. This attribute is ignored if hanging is specified. '''
			self.firstLine = dc.get('firstLine', '')

			self.parent = parent
			self.tab = self.parent.tab
			self.indent = self.parent.indent + 1

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def SetLeft(self, value):
			self.left = value

		def SetRight(self, value):
			self.right = value

		def SetHanging(self, value):
			self.hanging = value

		def SetFirstLine(self, value):
			self.firstLine = value

		def get_Left(self):
			return self.left

		def get_Right(self):
			return self.right

		def get_Hanging(self):
			return self.hanging

		def get_FirstLine(self):
			return self.firstLine

		def get_parent(self):
			return self.parent

		def get_xml(self):
			value = ''

			if self.left:
				value += ' w:left="%s"' % self.left
			if self.right:
				value += ' w:right="%s"' % self.right
			if self.hanging:
				value += ' w:hanging="%s"' % self.hanging
			if self.firstLine:
				value += ' w:firstLine="%s"' % self.firstLine

			if value:
				value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
			return value

	class OutlineLvl(object):
		def __init__(self, parent, value):
			self.name = 'outlineLvl'
			self.val = value  # 0-9
			self.parent = parent
			self.tab = self.parent.tab
			self.indent = self.parent.indent + 1

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def SetValue(self, value):
			self.val = value

		def get_Value(self):
			return self.val

		def get_parent(self):
			return self.parent

		def get_xml(self):
			value = '%s<w:%s w:val="%s"/>' % (self.get_tab(), self.name, self.val)
			return value

	class Borders(object):

		class Element(object):
			def __init__(self, parent, name, value, sz='', space='', shadow=False, color=''):
				self.name = name
				'''Specifies the style of the border. Paragraph borders can be only line borders. 
				(Page borders can also be art borders.) Possible values are:
				single - a single line
				dashDotStroked - a line with a series of alternating thin and thick strokes
				dashed - a dashed line
				dashSmallGap - a dashed line with small gaps
				dotDash - a line with alternating dots and dashes
				dotDotDash - a line with a repeating dot - dot - dash sequence
				dotted - a dotted line
				double - a double line
				doubleWave - a double wavy line
				inset - an inset set of lines
				nil - no border
				none - no border
				outset - an outset set of lines
				thick - a single line
				thickThinLargeGap - a thick line contained within a thin line with a large-sized intermediate gap
				thickThinMediumGap - a thick line contained within a thin line with a medium-sized intermediate gap
				thickThinSmallGap - a thick line contained within a thin line with a small intermediate gap
				thinThickLargeGap - a thin line contained within a thick line with a large-sized intermediate gap
				thinThickMediumGap - a thick line contained within a thin line with a medium-sized intermediate gap
				thinThickSmallGap - a thick line contained within a thin line with a small intermediate gap
				thinThickThinLargeGap - a thin-thick-thin line with a large gap
				thinThickThinMediumGap - a thin-thick-thin line with a medium gap
				thinThickThinSmallGap - a thin-thick-thin line with a small gap
				threeDEmboss - a three-staged gradient line, getting darker towards the paragraph
				threeDEngrave - a three-staged gradient like, getting darker away from the paragraph
				triple - a triple line
				wave - a wavy line'''

				self.val = value
				'''Specifies the width of the border. Paragraph borders are line borders (see the val attribute 
				below), and so the width is specified in eighths of a point, with a minimum value of two (1/4 of a 
				point) and a maximum value of 96 (twelve points).
				(Page borders can alternatively be art borders, with the width is given in points and a minimum of 
				1 and a maximum of 31.)'''
				self.sz = sz
				''''Specifies the spacing offset. Values are specified in points'''
				self.space = space
				'''Specifies whether the border should be modified to create the appearance of a shadow. For right 
				and bottom borders, this is done by duplicating the border below and right of the normal location.
				For the right and top borders, this is done by moving the border down and to the right of the 
				original location. Permitted values are true and false.'''
				self.shadow = shadow
				'''Specifies the color of the border. Values are given as hex values (in RRGGBB format). No #, 
				unlike hex values in HTML/CSS. E.g., color="FFFF00". A value of auto is also permitted and will 
				allow the consuming word processor to determine the color.'''
				self.color = color

				self.parent = parent
				self.tab = self.parent.tab
				self.indent = self.parent.indent + 1

			def get_tab(self, number=0):
				return self.tab * (self.indent + number)

			def set_color(self, value):
				self.color = value

			def SetShadow(self, value):
				self.shadow = value

			def SetSpace(self, value):
				self.space = value

			def SetSize(self, value):
				self.sz = value

			def SetValue(self, value):
				self.val = value

			def get_color(self):
				return self.color

			def get_Shadow(self):
				return self.shadow

			def get_Space(self):
				return self.space

			def get_Size(self):
				return self.sz

			def get_Value(self):
				return self.val

			def get_parent(self):
				return self.parent

			def get_xml(self):
				value = ''

				if self.val:
					value += ' w:val="%s"' % self.val
				if self.sz:
					value += ' w:sz="%s"' % self.sz
				if self.space:
					value += ' w:space="%s"' % self.space
				if self.shadow:
					value += ' w:shadow="true"'
				if self.color:
					value += ' w:color="%s"' % self.color

				if value:
					value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
				return value

		def __init__(self, parent, dc):
			self.name = 'pBdr'

			self.top = None
			if 'top' in dc.keys():
				self.top = self.Element(self, 'top', dc['top'].get('value', ''), dc['top'].get('sz', ''),
				                        dc['top'].get('space', ''), dc['top'].get('shadow', ''),
				                        dc['top'].get('color', ''))
			elif 'all' in dc.keys():
				self.top = self.Element(self, 'top', dc['all'].get('value', ''), dc['all'].get('sz', ''),
				                        dc['all'].get('space', ''), dc['all'].get('shadow', ''),
				                        dc['all'].get('color', ''))

			self.bottom = None
			if 'bottom' in dc.keys():
				self.bottom = self.Element(self, 'bottom', dc['bottom'].get('value', ''),
				                           dc['bottom'].get('sz', ''),
				                           dc['bottom'].get('space', ''), dc['bottom'].get('shadow', ''),
				                           dc['bottom'].get('color', ''))
			elif 'all' in dc.keys():
				self.bottom = self.Element(self, 'bottom', dc['all'].get('value', ''), dc['all'].get('sz', ''),
				                           dc['all'].get('space', ''), dc['all'].get('shadow', ''),
				                           dc['all'].get('color', ''))

			self.left = None
			if 'left' in dc.keys():
				self.left = self.Element(self, 'left', dc['left'].get('value', ''), dc['left'].get('sz', ''),
				                         dc['left'].get('space', ''), dc['left'].get('shadow', ''),
				                         dc['left'].get('color', ''))
			elif 'all' in dc.keys():
				self.left = self.Element(self, 'left', dc['all'].get('value', ''), dc['all'].get('sz', ''),
				                         dc['all'].get('space', ''), dc['all'].get('shadow', ''),
				                         dc['all'].get('color', ''))

			self.right = None
			if 'right' in dc.keys():
				self.right = self.Element(self, 'right', dc['right'].get('value', ''), dc['right'].get('sz', ''),
				                          dc['right'].get('space', ''), dc['right'].get('shadow', ''),
				                          dc['right'].get('color', ''))
			elif 'all' in dc.keys():
				self.right = self.Element(self, 'right', dc['all'].get('value', ''), dc['all'].get('sz', ''),
				                          dc['all'].get('space', ''), dc['all'].get('shadow', ''),
				                          dc['all'].get('color', ''))

			self.between = None
			if 'between' in dc.keys():
				self.between = self.Element(self, 'between', dc['between'].get('value', ''),
				                            dc['between'].get('sz', ''),
				                            dc['between'].get('space', ''), dc['between'].get('shadow', ''),
				                            dc['between'].get('color', ''))

			self.parent = parent
			self.tab = self.parent.tab
			self.separator = self.parent.separator
			self.indent = self.parent.indent + 1

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_separator(self):
			return self.separator

		def SetTop(self, value, sz='', space='', shadow=False, color=''):
			self.top = self.Element(self, 'top', value, sz, space, shadow, color)

		def SetBottom(self, value, sz='', space='', shadow=False, color=''):
			self.bottom = self.Element(self, 'bottom', value, sz, space, shadow, color)

		def SetLeft(self, value, sz='', space='', shadow=False, color=''):
			self.left = self.Element(self, 'left', value, sz, space, shadow, color)

		def SetRight(self, value, sz='', space='', shadow=False, color=''):
			self.right = self.Element(self, 'right', value, sz, space, shadow, color)

		def SetBetween(self, value, sz='', space='', shadow=False, color=''):
			self.between = self.Element(self, 'between', value, sz, space, shadow, color)

		def get_Top(self):
			return self.top

		def get_Bottom(self):
			return self.bottom

		def get_Left(self):
			return self.left

		def get_Right(self):
			return self.right

		def get_Between(self):
			return self.between

		def get_parent(self):
			return self.parent

		def get_xml(self):
			value = []

			if self.top is not None:
				value.append(self.top.get_xml())
			if self.left is not None:
				value.append(self.left.get_xml())
			if self.bottom is not None:
				value.append(self.bottom.get_xml())
			if self.right is not None:
				value.append(self.right.get_xml())
			if self.between is not None:
				value.append(self.between.get_xml())

			if value:
				value.insert(0, '%s<w:%s>' % (self.get_tab(), self.name))
				value.append('%s</w:%s>' % (self.get_tab() * self.indent, self.name))

			return self.get_separator().join(value)

	class Tab(object):
		def __init__(self, parent, value, posicion, leader=''):
			self.name = 'tab'
			'''Specifies the style of the tab. Possible values are:
				bar - a bar tab does not result in a tab stop in the parent paragraph but instead a vertical bar is 
				drawn at the location.
				center - all text folowing and preceding the next tab shall be centered around the tab.
				clear - the current tab stop is cleared and shall be removed.
				decimal - all following text is aligned around the first decimal character in the following text.
				end - all following and preceding text is aligned against the trailing edge.
				num - a list tab or the tab between the numbering and the contents. This is for backward 
				compatibility and should be avoided in favor of paragraph indentation.
				start - all following and preceding text is aligned against the leading edge.'''
			self.val = value

			'''Specifies the position of the tab stop. Values are in twips (1440 twips = one inch). Negative values 
			are permitted and move the page margin the specified amount. '''
			self.pos = posicion

			'''Specifies the character used to fill in the space created by a tab. The character is repeated as 
			required to fill the tab space. Possible values are:
				dot - a dot
				heavy - a heavy solid line or underscore
				hyphen - a hyphen or dash
				middleDot - a centered dot
				none - no leader character
				underscore - an underscore'''
			self.leader = leader

			self.parent = parent
			self.tab = self.parent.tab
			self.indent = self.parent.indent

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def SetLeader(self, value):
			self.leader = value

		def SetPosition(self, value):
			self.pos = value

		def SetValue(self, value):
			self.val = value

		def get_Value(self):
			return self.val

		def get_Leader(self):
			return self.leader

		def get_Position(self):
			return self.pos

		def get_parent(self):
			return self.parent

		def get_xml(self):
			value = '%s<w:tab w:val="%s" w:pos="%s"' % (self.get_tab(), self.val, self.pos)
			if self.leader:
				value += ' w:leader="%s"' % self.leader
			value += '/>'

			return value

	def __init__(self, parent, horizontal_alignment='j'):
		self.name = 'pPr'
		self.parent = parent
		self.separator = self.parent.separator
		self.tab = self.parent.tab
		self.indent = self.parent.indent + 1

		'''Defines the paragraph as a text frame, which is a free-standing paragraph similar to a text box'''
		self.framePr = None  # Sin hacer

		'''Defines the indentation for the paragraph'''
		self.ind = None  #

		'''Specifies the paragraph alignment.'''
		self.jc = element.HorizontalAlignment(self, horizontal_alignment)

		'''Specifies that all lines of the paragraph are to be kept on a single page when possible. It is an empty 
		element: <w:keeplines/>.'''
		self.keepLines = False

		'''Specifies that the paragraph (or at least part of it) should be rendered on the same page as the next 
		paragraph when possible. It is an empty element: <w:keepNext/>. 
		If multiple paragraphs are to be kept together but they exceed a page, then the set of paragraphs begin on a
		new page and page breaks are used thereafter as needed. '''
		self.keepNext = False

		'''Specifies the outline level associated with the paragraph. It is used to build the table of contents and 
		does not affect the appearance of the text. The single attribute val can have a value of from 0 to 9,'''
		'''where 9 indicates that no outline level applies to the paragraph. So <w:outlineLvl w:val="0"/> indicates
		that the paragraph is an outline level 1.'''
		self.outlineLvl = None

		'''Specifies borders for the paragraph'''
		self.pBdr = None

		'''Specifies the style ID of a paragraph style (<w:pStyle w:val="TestParagraphStyle"/>)'''
		self.pStyle = ''

		'''Para hacer listas'''
		self.numPr = None

		'''Specifies the run properties for the paragraph glyph, which is used to represent the physical location 
		of the paragraph mark. When the mark is formatted, a rPr appears within pPr. The text is then formatted 
		accordingly, except for possible direct text formatting. '''
		''' Text() '''
		self.rPr = None
		# Falta

		self.rFont = None

		'''Specifies the properties for a section. For all sections except the last section, the sectPr element is 
		stored as a child element of the last paragraph in the section. For the last section, the sectPr is stored 
		as a child element of the body element. '''
		self.sectPr = None
		# Falta

		'''Specifies shading for the paragraph.'''
		self.shd = None

		'''Specifies between paragraphs and between lines of a paragraph'''
		self.spacing = None

		'''Specifies custom tabs.'''
		'''List object Tab() anadir <w:tabs>+lista tabs+</w:tabs>'''
		self.tabs = []

		'''Specifies the alignment of characters on each line when characters are of varying size.'''
		'''Posibles Valores: (<w:textAlignment w:val="top"/>)
			auto
			baseline
			bottom
			center
			top'''
		self.textAlignment = ''

	def get_FormatList(self):
		return self.numPr

	def SetFormatList(self, level='1', _type='1'):
		self.numPr = self.FormatList(self, level, _type)
		self.set_pstyle('Prrafodelista')

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_separator(self):
		return self.separator

	def set_spacing(self, _dc):
		self.spacing = self.Spacing(self, _dc)

	def set_keep_lines(self, value):
		self.keepLines = value

	def set_keep_next(self, value):
		self.keepNext = value

	def set_horizontal_alignment(self, alineamiento):
		if alineamiento:
			self.jc = element.HorizontalAlignment(self, alineamiento)
		else:
			self.jc = None

	def SetIndentation(self, _dc):
		self.ind = self.Indentation(self, _dc)

	def SetOutlineLvl(self, value):
		self.outlineLvl = self.OutlineLvl(self, value)

	def SetBorder(self, _dc):
		self.pBdr = self.Borders(self, _dc)

	def set_pstyle(self, value):
		self.pStyle = value

	def set_shading(self, value, fill='', color=''):
		self.shd = element.Shading(value, fill, color)

	def SetTabs(self, tabs):
		for tab in tabs:
			if len(tab) == 3:
				value, posicion, leader = tab
			else:
				value, posicion = tab
				leader = ''
			self.tabs.append(self.Tab(value, posicion, leader))

	def SetTextAlignment(self, value):
		self.textAlignment = value

	def get_spacing(self):
		return self.spacing

	def get_KeepLines(self):
		return self.keepLines

	def get_KeepNext(self):
		return self.keepNext

	def get_horizontal_alignment(self):
		return self.jc

	def get_Indentation(self):
		return self.ind

	def get_OutlineLvl(self):
		return self.outlineLvl

	def get_Border(self):
		return self.pBdr

	def get_PStyle(self):
		return self.pStyle

	def get_shading(self):
		return self.shd

	def get_tabs(self):
		return self.tabs

	def get_TextAlignment(self):
		return self.textAlignment

	def get_name(self):
		return self.name

	def get_parent(self):
		return self.parent

	def get_xml(self):
		value = []
		if self.keepLines and not self.keepNext:
			value.append('%s<w:keeplines/>' % (self.get_tab(1)))
		if self.keepNext:
			value.append('%s<w:keepNext/>' % (self.get_tab(1)))
		if self.pBdr is not None:
			value.append(self.pBdr.get_xml())
		if self.spacing is not None:
			x = self.spacing.get_xml()
			if x:
				value.append(x)
		if self.ind is not None:
			value.append(self.ind.get_xml())
		if self.outlineLvl is not None:
			value.append(self.outlineLvl.get_xml())
		if self.framePr is not None:
			value.append(self.framePr.get_xml())
		if self.pStyle:
			value.append('%s<w:pStyle w:val="%s"/>' % (self.get_tab(1), self.pStyle))

		if self.numPr is not None:
			value.append(self.numPr.get_xml())

		if self.jc is not None:
			value.append(self.jc.get_xml())
		if self.textAlignment:
			value.append('%s<w:textAlignment w:val="%s"/>' % (self.get_tab(1), self.textAlignment))
		if self.rPr:
			value.append(self.rPr.get_xml())
		if self.sectPr:
			value.append(self.sectPr.get_xml())
		if self.shd is not None:
			value.append(self.shd.get_xml())
		if self.tabs:
			value.append('%s<w:tabs>' % (self.get_tab(1)))
			for tab in self.tabs:
				value.append(self.tab + tab.get_xml())
			value.append('%s</w:tabs>' % (self.get_tab(1)))

		if value:
			value.insert(0, '%s<w:%s>' % (self.get_tab(), self.name))
			# value.append('%s<w:lang w:val="es-ES_tradnl"/>' % self.get_tab(1))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))
		return self.separator.join(value)


class Paragraph(object):

	def __init__(self, parent, idx, _text=(), horizontal_alignment='j', font_format='', font_size=None, nulo=False):
		self.id = idx
		self.name = 'p'
		self.parent = parent
		if self.parent is not None:
			self.separator = self.parent.separator
			self.tab = self.parent.tab
			self.indent = parent.indent + 1
		else:
			self.separator = '\n'
			self.tab = '\t'
			self.indent = 2

		self.nulo = nulo

		'''Atributos'''
		self.rsidR = '00C27534'
		self.rsidRDefault = '000F5F75'
		self.rsidP = '00DC4651'
		self.textId = '77777777'
		self.paraId = '0577C6EC'
		self.rsidRPr = ''
		self.bookmarkStart = None  # Falta
		self.bookmarkEnd = None  # Falta
		self.fldSimple = None  # Falta
		self.hyperlink = None  # Falta

		self.properties = Properties(self, horizontal_alignment)

		self.elements = list()

		if getattr(_text, 'name', '') == 'r':
			if font_size is not None:
				_text.get_properties().set_font_size(font_size)
			self.elements.append(_text)
		elif getattr(_text, 'name', '') in ['drawing', 'shape']:
			self.elements.append(_text)
		# self.get_properties().set_horizontal_alignment(None)
		elif _text is not None:
			if type(_text) == str:
				_text = [_text]
			for k in range(len(_text)):
				txt = _text[k]
				if type(font_format) == str:
					ff = font_format
				else:
					ff = ''
					if k < len(font_format):
						ff = font_format[k]

				if getattr(txt, 'name', '') == 'r':
					if font_size is not None:
						txt.get_properties().set_font_size(font_size)
					self.elements.append(txt)
				else:
					self.elements.append(text.Text(self, txt, font_format=ff, font_size=font_size))

	def AddPicture(self, parent, path, width, height, anchor='inline'):
		is_body = parent.tag == 'w:body'

		document = self
		while getattr(document, 'tag', '') != 'document':
			document = document.parent
		name = path.split('/')[-1]
		extension = name.split('.')[-1]
		document.get_content_types().AddDefault(extension, 'ContentType="image/%s"' % extension)
		img = open(path, 'rb').read()

		if is_body:
			rid = document.idx
			document.idx += 1
			target = 'media/image%d.%s' % (rid, extension)
			document.add_part(target, ImagePart(target))
			document.add_image(target, img, rid)
		else:
			name_part = parent.get_name().split('/')[-1]

			rid = parent.get_RelRId()
			rid = document.idx
			document.idx += 1
			print rid, name_part
			parent.AddRelRId()

			target = 'media/image%d.%s' % (rid, extension)

			if name_part not in document.get_parts().keys():
				document.add_part_rel(name_part)
			rel = document.get_part(name_part)
			rel.add_part(
				target,
				{"Id": "rId%d" % rid, "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"})
			rel.add_image(target, img)

		pict = pictures.Picture(self, rid, path, width, height, anchor=anchor)
		self.elements.append(pict)
		return self

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def SetParent(self, parent):
		self.parent = parent
		self.separator = self.parent.separator
		self.tab = self.parent.tab
		self.indent = parent.indent + 1
		self.get_properties().tab = self.tab
		self.get_properties().separator = self.separator

	def SetFormatList(self, level='1', _type='1'):
		self.get_properties().SetFormatList(level, _type)

	def get_separator(self):
		return self.separator

	def SetIndent(self, value):
		self.indent = value

	def get_Indent(self):
		return self.indent

	def Set_rsidP(self, value):
		self.rsidP = value

	def Set_rsidR(self, value):
		self.rsidR = value

	def Set_rsidRPr(self, value):
		self.rsidRPr = value

	def Set_rsidRDefault(self, value):
		self.rsidRDefault = value

	def get_parent(self):
		return self.parent

	def get_name(self):
		return self.name

	def get_id(self):
		return self.id

	def get_rsidP(self):
		return self.rsidP

	def get_rsidR(self):
		return self.rsidR

	def get_rsidRPr(self):
		return self.rsidRPr

	def get_rsidRDefault(self):
		return self.rsidRDefault

	def get_texts(self):
		return self.elements

	def get_Text(self, index):
		return self.elements[index]

	def add_text(self, txt, font_format='', font_size=10):
		self.elements.append(text.Text(self, txt, font_format=font_format, font_size=font_size))
		return self.elements[-1]

	def get_properties(self):
		return self.properties

	def set_horizontal_alignment(self, alignment):
		self.get_properties().set_horizontal_alignment(alignment)

	def get_horizontal_alignment(self):
		return self.get_properties().get_horizontal_alignment()

	def set_spacing(self, dc):
		self.get_properties().set_spacing(dc)

	def add_element(self, _element):
		self.elements.append(_element)

	def IsNulo(self):
		return self.nulo

	def set_font_format(self, font_format):
		for k in range(len(self.get_texts())):
			element = self.get_Text(k)
			ff = ''
			if type(font_format) is str:
				ff = font_format
			elif type(font_format) is list:
				if len(font_format) > k:
					ff = font_format[k]

			if getattr(element, 'name', '') == 'r':
				if 'b' in ff:
					element.get_properties().bold()
				if 'i' in ff:
					element.get_properties().italic()
				if 'u' in ff:
					element.get_properties().underline()

	def set_font_size(self, size):
		for text in self.get_texts():
			if hasattr(text, 'get_properties'):
				if hasattr(text.get_properties(), 'set_font_size'):
					text.get_properties().set_font_size(size)

	def get_xml(self):
		value = ['%s<w:%s' % (self.get_tab(), self.name)]

		if not self.elements:
			return value[0] + '/>'
		if self.rsidR is not '':
			value[0] += ' w:rsidR="%s"' % self.rsidR

		if self.rsidRPr is not '':
			value[0] += ' w:rsidRPr="%s"' % self.rsidRPr

		if self.rsidRDefault is not '':
			value[0] += ' w:rsidRDefault="%s"' % self.rsidRDefault

		if self.rsidP is not '':
			value[0] += ' w:rsidP="%s"' % self.rsidP
		value[0] += '>'

		value.append(self.get_properties().get_xml())

		'''
		self.fldSimple=None#Falta
		self.hyperlink=None#Falta
		'''
		if self.IsNulo():
			value.append('%s<w:r>' % (self.get_tab(1)))
			value.append('%s<w:lastRenderedPageBreak/>' % (self.get_tab(2)))
			value.append('%s<w:br w:type="page"/>' % (self.get_tab(2)))
			value.append('%s</w:r>' % (self.get_tab(1)))

		for txt in self.elements:
			# if isinstance(txt, pictures.Picture):
			if getattr(txt, 'name', '') == 'drawing':
				value.append('%s<w:r>' % (self.get_tab(1)))
				value.append('%s<w:rPr>' % (self.get_tab(2)))
				value.append('%s<w:noProof/>' % (self.get_tab(3)))
				value.append('%s</w:rPr>' % (self.get_tab(2)))
				value.append(txt.get_xml())
				value.append('%s</w:r>' % (self.get_tab(1)))

			elif getattr(txt, 'name', '') == 'shape':
				value.append('%s<w:r>' % (self.get_tab(1)))
				value.append('%s<w:rPr>' % (self.get_tab(2)))
				value.append('%s<w:noProof/>' % (self.get_tab(3)))
				value.append('%s</w:rPr>' % (self.get_tab(2)))
				value.append(txt.get_xml())
				value.append('%s</w:r>' % (self.get_tab(1)))

			else:
				value.append(txt.get_xml())

		value.append('%s</w:%s>' % (self.get_tab(), self.name))

		return self.separator.join(value)
