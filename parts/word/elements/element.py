#!/usr/bin/python
# -*- coding: utf-8 -*-


class HorizontalAlignment:

	def __init__(self, parent, value='left'):

		"""value: Specifies the value for the paragraph justification. Possible values are:
		start - justification to the left margin for left-to-right documents.
		end - justification to the right margin for left-to-right documents.
		center - center the text.
		left
		right
		both - justify text between both margins equally, but inter-character spacing is not affected.
		distribute - justify text between both margins equally, and both inter-word and inter-character spacing
		are affected."""

		self.name = 'jc'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent +1

		self.val = value

	def SetValue(self, value):
		self.val = value

	def get_Value(self):
		return self.val

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):

		if self.val in ['', 'l']:
			self.val = 'left'
		if self.val == 'r':
			self.val = 'right'
		if self.val == 'c':
			self.val = 'center'
		if self.val == 'j':
			self.val = 'both'
		if self.val == 'd':
			self.val = 'distribute'
		value = '%s<w:%s w:val="%s"/>' % (self.get_tab(), self.name, self.val)
		return value


class Shading:

	def __init__(self, parent, color='', mode='auto', value='clear'):
		"""Class: Shading
Atributtes:
	color:
		Specifies the color to be used for the background. Values are given as hex values (i.e., in RRGGBB format).
		No # is included, unlike hex values in HTML/CSS.
		E.g., fill="FFFF00". A value of auto is possible, enabling the consuming software to determine the value.
		Specifies the color to be used for any foreground pattern specified with the val attribute.
	mode:
		Values are given as hex values (in RRGGBB format).
		No #, unlike hex values in HTML/CSS. E.g., fill="FFFF00". A value of auto is possible, enabling the consuming
		software to determine the value.
	value:
		Specifies the pattern to be used to lay the pattern color over the background color. For example, w:val="pct10"
		indicates that the border style is a 10 percent foreground fill mask.
		Possible values are: clear (no pattern), pct10, pct12, pct15 . . ., diagCross, diagStripe, horzCross, horzStripe,
		nil, thinDiagCross, solid, etc. See ECMA-376, 3rd, � 17.18.78 for a complete listing.
"""
		self.name = 'shd'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.fill = color
		self.color = mode
		self.val = value

	def SetFill(self, value):
		self.fill = value

	def set_color(self, value):
		self.color = value

	def SetValue(self, value):
		self.val = value

	def get_Fill(self):
		return self.fill

	def get_color(self):
		return self.color

	def get_Value(self):
		return self.val

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''
		if self.val:
			value += ' w:val="%s"' % self.val
		if self.color:
			value += ' w:color="%s"' % self.color
		if self.fill:
			value += ' w:fill="%s"' % self.fill

		if value:
			value = '%s<w:%s %s/>' % (self.get_tab(), self.name, value)

		return value


class CellWidth:

	def __init__(self, parent, w, type_='dxa'):
		self.name = 'tcW'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.w = w
		self.type = type_

	def SetWidth(self, value):
		self.w = value

	def get_width(self):
		return self.w

	def set_type(self, value):
		self.type = value

	def get_type(self):
		return self.type

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''
		if self.w:
			value += ' w:w="%s"' % self.w

		if self.type:
			value += ' w:type="%s"' % self.type

		if value:
			value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
		return value


class Element:
	def __init__(self, parent, name, w, type_='dxa'):
		self.name = name
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.w = w
		'''Specifies the units of the width (w) property. Possible values are:
			dxa - Specifies that the value is in twentieths of a point (1/1440 of an inch).
			nil - Specifies a value of zero'''
		self.type = type_

	def SetWidth(self, value):
		self.w = value

	def get_width(self):
		return self.w

	def set_type(self, value):
		self.type = value

	def get_type(self):
		return self.type

	def get_name(self):
		return self.name

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''

		value += ' w:w="%s"' % self.w
		if self.type:
			value += ' w:type="%s"' % self.type

		if value:
			value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
		return value


class TableElement:
	def __init__(self, parent, name, value='single', sz='8', space='', shadow=False, color=''):
		self.name = name
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		'''Specifies the style of the border. Paragraph borders can be only line borders. (Page borders can also be 
		art borders.) Possible values are: single - a single line dashDotStroked - a line with a series of alternating 
		thin and thick strokes dashed - a dashed line dashSmallGap - a dashed line with small gaps dotDash - a line 
		with alternating dots and dashes dotDotDash - a line with a repeating dot - dot - dash sequence dotted - a 
		dotted line double - a double line doubleWave - a double wavy line inset - an inset set of lines nil - no 
		border none - no border outset - an outset set of lines thick - a single line thickThinLargeGap - a thick line 
		contained within a thin line with a large-sized intermediate gap thickThinMediumGap - a thick line contained 
		within a thin line with a medium-sized intermediate gap thickThinSmallGap - a thick line contained within a 
		thin line with a small intermediate gap thinThickLargeGap - a thin line contained within a thick line with a 
		large-sized intermediate gap thinThickMediumGap - a thick line contained within a thin line with a 
		medium-sized intermediate gap thinThickSmallGap - a thick line contained within a thin line with a small 
		intermediate gap thinThickThinLargeGap - a thin-thick-thin line with a large gap thinThickThinMediumGap - a 
		thin-thick-thin line with a medium gap thinThickThinSmallGap - a thin-thick-thin line with a small gap 
		threeDEmboss - a three-staged gradient line, getting darker towards the paragraph threeDEngrave - a 
		three-staged gradient like, getting darker away from the paragraph triple - a triple line wave - a wavy line 
		'''

		self.val = value
		'''Specifies the width of the border. Paragraph borders are line borders (see the val attribute below), and so
			the width is specified in eighths of a point, with a minimum value of two (1/4 of a point) and a maximum value
			of 96 (twelve points).
		(Page borders can alternatively be art borders, with the width is given in points and a minimum of 1 and a maximum 
		of 31.)'''
		self.sz = str(sz)
		''''Specifies the spacing offset. Values are specified in points'''
		self.space = space
		'''Specifies whether the border should be modified to create the appearance of a shadow. For right and bottom 
		borders, this is done by duplicating the border below and right of the normal location.
		For the right and top borders, this is done by moving the border down and to the right of the original location.'''
		self.shadow = shadow
		'''Specifies the color of the border. Values are given as hex values (in RRGGBB format). No #, unlike hex 
		values in HTML/CSS. E.g., color="FFFF00". A value of auto is also permitted and will allow the consuming word 
		processor to determine the color. '''
		self.color = color

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

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''

		if self.val:
			value += ' w:val="%s"' % self.val
		if self.sz:
			value += ' w:sz="%s"' % self.sz
		if self.space:
			value += ' w:space="%s"' % self.space
		if self.color:
			value += ' w:color="%s"' % self.color
		if self.shadow:
			value += ' w:shadow="true"'

		if value:
			value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
		return value


class TableBorders(object):

	def __init__(self, parent, dc):
		self.name = 'tblBorders'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.top = None
		if 'top' in dc.keys():
			self.top = TableElement(self, 'top', dc['top'].get('value', 'single'), dc['top'].get('sz', '4'),
					dc['top'].get('space', ''), dc['top'].get('shadow', ''), dc['top'].get('color', ''))
		elif 'all' in dc.keys():
			self.top = TableElement(self, 'top', dc['all'].get('value', 'single'), dc['all'].get('sz', '4'),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''), dc['all'].get('color', ''))
		self.bottom = None
		if 'bottom' in dc.keys():
			self.bottom = TableElement(self, 'bottom', dc['bottom'].get('value', 'single'), dc['bottom'].get('sz', '4'),
					dc['bottom'].get('space', ''), dc['bottom'].get('shadow', ''),
					dc['bottom'].get('color', ''))
		elif 'all' in dc.keys():
			self.bottom = TableElement(self, 'bottom', dc['all'].get('value', 'single'), dc['all'].get('sz', '4'),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

		self.start = None
		if 'start' in dc.keys():
			self.start = TableElement(self, 'start', dc['start'].get('value', 'single'), dc['start'].get('sz', '4'),
					dc['start'].get('space', ''), dc['start'].get('shadow', ''),
					dc['start'].get('color', ''))
		elif 'all' in dc.keys():
			self.start = TableElement(self, 'start', dc['all'].get('value', 'single'), dc['all'].get('sz', '4'),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

		self.end = None
		if 'end' in dc.keys():
			self.end = TableElement(self, 'end', dc['end'].get('value', 'single'), dc['end'].get('sz', '4'),
					dc['end'].get('space', ''), dc['end'].get('shadow', ''), dc['end'].get('color', ''))
		elif 'all' in dc.keys():
			self.end = TableElement(self, 'end', dc['all'].get('value', 'single'), dc['all'].get('sz', '4'),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''), dc['all'].get('color', ''))

		self.insideH = None
		if 'insideH' in dc.keys():
			self.insideH = TableElement(self, 'insideH', dc['insideH'].get('value', 'single'),
					dc['insideH'].get('sz', '4'),
					dc['insideH'].get('space', ''), dc['insideH'].get('shadow', ''),
					dc['insideH'].get('color', ''))
		elif 'all' in dc.keys():
			self.insideH = TableElement(self, 'insideH', dc['all'].get('value', 'single'), dc['all'].get('sz', '4'),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

		self.insideV = None
		if 'insideV' in dc.keys():
			self.insideV = TableElement(self, 'insideV', dc['insideV'].get('value', 'single'),
					dc['insideV'].get('sz', '4'),
					dc['insideV'].get('space', ''), dc['insideV'].get('shadow', ''),
					dc['insideV'].get('color', ''))
		elif 'all' in dc.keys():
			self.insideV = TableElement(self, 'insideV', dc['all'].get('value', 'single'), dc['all'].get('sz', '4'),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

	def SetTop(self, value, sz='', space='', shadow=False, color=''):
		self.top = TableElement(self, 'top', value, sz, space, shadow, color)

	def SetBottom(self, value, sz='', space='', shadow=False, color=''):
		self.bottom = TableElement(self, 'bottom', value, sz, space, shadow, color)

	def SetStart(self, value, sz='', space='', shadow=False, color=''):
		self.start = TableElement(self, 'start', value, sz, space, shadow, color)

	def SetEnd(self, value, sz='', space='', shadow=False, color=''):
		self.end = TableElement(self, 'end', value, sz, space, shadow, color)

	def SetInsideH(self, value, sz='', space='', shadow=False, color=''):
		self.insideH = TableElement(self, 'insideH', value, sz, space, shadow, color)

	def SetInsideV(self, value, sz='', space='', shadow=False, color=''):
		self.insideV = TableElement(self, 'insideV', value, sz, space, shadow, color)

	def get_Top(self):
		return self.top

	def get_Bottom(self):
		return self.bottom

	def get_Start(self):
		return self.start

	def get_End(self):
		return self.end

	def get_InsideH(self):
		return self.insideH

	def get_InsideV(self):
		return self.insideV

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = []

		if self.top is not None:
			value.append(self.top.get_xml())
		if self.bottom is not None:
			value.append(self.bottom.get_xml())
		if self.start is not None:
			value.append(self.start.get_xml())
		if self.end is not None:
			value.append(self.end.get_xml())
		if self.insideH is not None:
			value.append(self.insideH.get_xml())
		if self.insideV is not None:
			value.append(self.insideV.get_xml())

		if value:
			value.insert(0, '%s<w:%s>' % (self.get_tab(), self.name))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))

		return self.separator.join(value)


class CellMargin:

	def __init__(self, parent, dc):
		self.name = 'tblCellMar'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.top = None
		if 'top' in dc.keys():
			self.top = Element(self, 'top', dc['top'].get('w', '0'), dc['top'].get('type', 'dxa'))
		elif 'all' in dc.keys():
			self.top = Element(self, 'top', dc['all'].get('w', '0'), dc['all'].get('type', 'dxa'))

		self.bottom = None
		if 'bottom' in dc.keys():
			self.bottom = Element(self, 'bottom', dc['bottom'].get('w', '0'), dc['bottom'].get('type', 'dxa'))
		elif 'all' in dc.keys():
			self.bottom = Element(self, 'bottom', dc['all'].get('w', '0'), dc['all'].get('type', 'dxa'))

		self.start = None
		if 'start' in dc.keys():
			self.start = Element(self, 'start', dc['start'].get('w', '0'), dc['start'].get('type', 'dxa'))
		elif 'left' in dc.keys():
			self.start = Element(self, 'left', dc['left'].get('w', '0'), dc['left'].get('type', 'dxa'))
		elif 'all' in dc.keys():
			self.start = Element(self, 'left', dc['all'].get('w', '0'), dc['all'].get('type', 'dxa'))

		self.end = None
		if 'end' in dc.keys():
			self.end = Element(self, 'end', dc['end'].get('w', '0'), dc['end'].get('type', 'dxa'))
		elif 'right' in dc.keys():
			self.end = Element(self, 'right', dc['right'].get('w', '0'), dc['right'].get('type', 'dxa'))
		elif 'all' in dc.keys():
			self.end = Element(self, 'right', dc['all'].get('w', '0'), dc['all'].get('type', 'dxa'))

	def SetTop(self, x, type_='dxa'):
		self.top = Element(self, 'top', x, type_)

	def SetBottom(self, x, type_='dxa'):
		self.bottom = Element(self, 'bottom', x, type_)

	def SetStart(self, x, type_='dxa'):
		self.start = Element(self, 'start', x, type_)

	def SetEnd(self, x, type_='dxa'):
		self.end = Element(self, 'end', x, type_)

	def get_Top(self):
		return self.top

	def get_Bottom(self):
		return self.bottom

	def get_Start(self):
		return self.start

	def get_End(self):
		return self.end

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = []

		if self.top is not None:
			value.append(self.top.get_xml())
		if self.bottom is not None:
			value.append(self.bottom.get_xml())
		if self.start is not None:
			value.append(self.start.get_xml())
		if self.end is not None:
			value.append(self.end.get_xml())

		if value:
			value.insert(0, '%s<w:%s>' % (self.get_tab(), self.name))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))

		return self.separator.join(value)


class CellSpacing:

	def __init__(self, parent, w, type_='dxa'):
		self.name = 'tblCellSpacing'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.tblCellSpacing = Element(self, 'tblCellSpacing', w, type_)

	def SetTblCellSpacing(self, w, type_='dxa'):
		self.tblCellSpacing = Element(self, 'tblCellSpacing', w, type_)

	def get_TblCellSpacing(self):
		return self.tblCellSpacing

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''
		if self.tblCellSpacing is not None:
			value = self.tblCellSpacing.get_xml()

		return value


class TableIndentation:

	def __init__(self, parent, w, type_='dxa'):
		self.name = 'tblInd'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.w = w
		self.type = type_

	def SetWidth(self, value):
		self.w = value

	def set_type(self, value):
		self.type = value

	def get_width(self):
		return self.w

	def get_type(self):
		return self.type

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''

		if self.w:
			value += ' w:w="%s"' % self.w
		if self.type:
			value += ' w:type="%s"' % self.type

		if value:
			value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
		return value


class TableLayout:

	def __init__(self, parent, value='fixed'):
		self.name = 'tblLayout'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		'''Specifies the method of laying out the contents of the table. Possible values are:				
			fixed - Uses the preferred widths on the table items to generate the final sizing of the table. The width of
				the table is not changed regardless of the contents of the cells.
			autofit - Uses the preferred widths on the table items to generat the sizing of the table, but then uses the
				contents of each cell ot determine the final column widths.'''
		self.value = value

	def SetValue(self, value):
		self.value = value

	def get_Value(self):
		return self.value

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''

		if self.value:
			value += ' w:type="%s"' % self.value

		if value:
			value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
		return value


class TableLook:

	def __init__(self, parent, value='04A0', first_row='1', last_row='0', first_column='1', last_column='0', nohband='0',
					novband='1'):
		self.name = 'tblLook'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.value = value
		self.firstRow = first_row
		self.lastRow = last_row
		self.firstColumn = first_column
		self.lastColumn = last_column
		self.noHBand = nohband
		self.noVBand = novband

	def SetValue(self, value):
		self.value = value

	def get_Value(self):
		return self.value

	def SetFirstColumn(self, value):
		self.firstColumn = value

	def get_FirstColumn(self):
		return self.firstColumn

	def SetLastColumn(self, value):
		self.lastColumn = value

	def get_LastColumn(self):
		return self.lastColumn

	def SetFirstRow(self, value):
		self.firstRow = value

	def get_FirstRow(self):
		return self.firstRow

	def SetLastRow(self, value):
		self.lastRow = value

	def get_LastRow(self):
		return self.lastRow

	def SetNoHBand(self, value):
		self.noHBand = value

	def get_NoHBand(self):
		return self.noHBand

	def SetNoVBand(self, value):
		self.noVBand = value

	def get_NoVBand(self):
		return self.noVBand

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''

		if self.value:
			value += ' w:val="%s"' % self.value
		if self.firstRow:
			value += ' w:firstRow="%s"' % self.firstRow
		if self.lastRow:
			value += ' w:lastRow="%s"' % self.lastRow
		if self.firstColumn:
			value += ' w:firstColumn="%s"' % self.firstColumn
		if self.lastColumn:
			value += ' w:lastColumn="%s"' % self.lastColumn
		if self.noHBand:
			value += ' w:noHBand="%s"' % self.noHBand
		if self.noVBand:
			value += ' w:noVBand="%s"' % self.noVBand
		if value:
			value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
		return value


class FloatingTable:

	def __init__(self, parent):
		self.name = 'tblPr '
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		'''Specifies the horizontal anchor or the base object from which the horizontal positioning in the tblpX or 
		tblpXSpec attribute should be determined. Possible values are:
			margin - relative to the vertical edge of the text margin before any text runs (left edge for left-to-right 
			paragraphs)
			page - relative to the vertical edge of the page before any text runs (left edge for left-to-right paragraphs)
			text - relative to the vertical edge of the text margin for the column in which the anchor paragraph is located
		If omitted, the value is assumed to be page.'''
		self.horzAnchor = None

		'''Specifies the vertical anchor or the base object from which the vertical positioning in the tblpY attribute 
		should be determined. Possible values are:
			margin - relative to the horizontal edge of the text margin before any text runs (top edge for top-to-bottom 
			paragraphs)
			page - relative to the horizontal edge of the page before any text runs (top edge for top-to-bottom paragraphs)
			text - relative to the horizontal edge of the text margin for the column in which the anchor paragraph is located
		If omitted, the value is assumed to be page.'''
		self.vertAnchor = None

		'''Specifies an absolute horizontal position for the table, relative to the horzAnchor anchor. The value is in 
		twentieths of a point. Note that the value can be negative, 
		in which case the table is positioned before the anchor object in the direction of horizontal text flow. 
		If tblpXSpec is also specified, then the tblpX attribute is ignored. If the attribute is omitted, the value is 
			to be zero. '''
		self.tblpX = None
		self.tblpY = None

		'''Specifies a relative horizontal position for the table, relative to the horzAnchor attribute. 
		This will supersede the tblpX attribute. Possible values are:
			center - the table should be horizontally centered with respect to the anchor
			inside - the table should be inside of the anchor
			left - the table should be left aligned with respect to the anchor
			outside - the table should be outside of the anchor
			right - the table should be right aligned with respect to the anchor'''
		self.tblpXSpec = None

		'''Specifies an absolute vertical position for the table, relative to the vertAnchor anchor. The value is in 
		twentieths of a point. Note that the value can be negative, in which case the table is positioned before the 
		anchor object in the direction of vertical text flow. 
		If tblpYSpec is also specified, then the tblpX attribute is ignored. If the attribute is omitted, the value is 
		assumed to be zero. '''
		self.tblpYSpec = None

		'''Specifies a relative vertical position for the table, relative to the vertAnchor attribute. 
		This will supersede the tblpY attribute. Possible values are:
			center - the table should be vertically centered with respect to the anchor
			inside - the table should be vertically aligned to the edge of the anchor and inside the anchor
			bottom - the table should be vertically aligned to the bottom edge of the anchor
			outside - the table should be vertically aligned to the edge of the anchor and outside the anchor
			inline - the table should be vertically aligned in line with the surrounding text (so as to not allow any 
			text wrapping around it)
			top - the table should be vertically aligned to the top edge of the anchor'''
		self.tblpXSpec = None

		'''Specifies the minimun distance to be maintained between the table and the top of text in the paragraph below 
		the table. The value is in twentieths of a point. If omitted, the value is assumed to be zero. '''
		self.bottomFromText = None

		'''Specifies the minimun distance to be maintained between the table and the bottom edge of text in the 
		paragraph above the table. The value is in twentieths of a point. If omitted, the value is assumed to be zero. '''
		self.topFromText = None

		'''Specifies the minimun distance to be maintained between the table and the edge of text in the paragraph to 
		the left of the table. The value is in twentieths of a point. If omitted, the value is assumed to be zero.'''
		self.leftFromText = None

		'''Specifies the minimun distance to be maintained between the table and the edge of text in the paragraph to 
		the right of the table. The value is in twentieths of a point. If omitted, the value is assumed to be zero'''
		self.rightFromText = None

	def SetBottomFromText(self, value):
		self.bottomFromText = value

	def get_BottomFromText(self):
		return self.bottomFromText

	def SetTopFromText(self, value):
		self.topFromText = value

	def get_TopFromText(self):
		return self.topFromText

	def SetLeftFromText(self, value):
		self.leftFromText = value

	def get_LeftFromText(self):
		return self.leftFromText

	def SetRightFromText(self, value):
		self.rightFromText = value

	def get_RightFromText(self):
		return self.rightFromText

	def SetAnclajeHorizontal(self, value):
		self.horzAnchor = value

	def get_AnclajeVertical(self):
		return self.horzAnchor

	def SetAnclajeVertical(self, value):
		self.vertAnchor = value

	def get_AnclajeHorizontal(self):
		return self.vertAnchor

	def SetPositionX(self, value):
		self.tblpX = value

	def get_PositionX(self):
		return self.tblpX

	def SetRelativePositionX(self, value):
		self.tblpXSpec = value

	def get_RelativePositionX(self):
		return self.tblpXSpec

	def SetPositionY(self, value):
		self.tblpY = value

	def get_PositionY(self):
		return self.tblpY

	def SetRelativePositionY(self, value):
		self.tblpYSpec = value

	def get_RelativePositionY(self):
		return self.tblpYSpec

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = ''
		if self.horzAnchor is not None:
			value += ' w:horzAnchor="%s"' % self.horzAnchor

		if self.vertAnchor is not None:
			value += ' w:vertAnchor="%s"' % self.vertAnchor

		if self.tblpX is not None:
			value += ' w:tblpX="%s"' % self.tblpX

		if self.tblpXSpec is not None:
			value += ' w:tblpXSpec="%s"' % self.tblpXSpec

		if self.tblpYSpec is not None:
			value += ' w:tblpYSpec="%s"' % self.tblpYSpec

		if self.tblpXSpec is not None:
			value += ' w:tblpXSpec="%s"' % self.tblpXSpec

		if self.bottomFromText is not None:
			value += ' w:bottomFromText="%s"' % self.bottomFromText

		if self.topFromText is not None:
			value += ' w:topFromText="%s"' % self.topFromText

		if self.leftFromText is not None:
			value += ' w:leftFromText="%s"' % self.leftFromText

		if self.rightFromText is not None:
			value += ' w:rightFromText="%s"' % self.rightFromText

		if value:
			value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)
		return value


class CellBorders:
	class CellElement:
		def __init__(self, parent, name, value, sz='', space='', shadow=False, color=''):
			self.name = name
			self.parent = parent
			self.indent = parent.indent + 1
			self.tab = parent.tab
			self.separator = parent.separator

			'''Specifies the style of the border. Paragraph borders can be only line borders. (Page borders can also be 
			art borders.) Possible values are:
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
			'''Specifies the width of the border. Paragraph borders are line borders (see the val attribute below), and 
			so the width is specified in eighths of a point, with a minimum value of two (1/4 of a point) and a maximum 
			value of 96 (twelve points).
			(Page borders can alternatively be art borders, with the width is given in points and a minimum of 1 and a 
			maximum of 31.)'''
			self.sz = sz
			''''Specifies the spacing offset. Values are specified in points'''
			self.space = space
			'''Specifies whether the border should be modified to create the appearance of a shadow. For right and 
			bottom borders, this is done by duplicating the border below and right of the normal location.
			For the right and top borders, this is done by moving the border down and to the right of the original 
			location. Permitted values are true and false.'''
			self.shadow = shadow
			'''Specifies the color of the border. Values are given as hex values (in RRGGBB format). No #, unlike hex 
			values in HTML/CSS. E.g., color="FFFF00". A value of auto is also permitted and will allow the consuming 
			word processor to determine the color.'''
			self.color = color

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

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

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
		self.name = 'tcBorders'
		self.parent = parent
		self.indent = parent.indent + 1
		self.tab = parent.tab
		self.separator = parent.separator

		self.top = None
		if 'top' in dc.keys():
			self.top = self.CellElement(self, 'top', dc['top'].get('value', 'single'), dc['top'].get('sz', ''),
					dc['top'].get('space', ''), dc['top'].get('shadow', ''),
					dc['top'].get('color', ''))
		elif 'all' in dc.keys():
			self.top = self.CellElement(self, 'top', dc['all'].get('value', 'single'), dc['all'].get('sz', ''),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

		self.bottom = None
		if 'bottom' in dc.keys():
			self.bottom = self.CellElement(self, 'bottom', dc['bottom'].get('value', 'single'),
					dc['bottom'].get('sz', ''),
					dc['bottom'].get('space', ''), dc['bottom'].get('shadow', ''),
					dc['bottom'].get('color', ''))
		elif 'all' in dc.keys():
			self.bottom = self.CellElement(self, 'bottom', dc['all'].get('value', 'single'), dc['all'].get('sz', ''),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

		self.start = None
		if 'start' in dc.keys():
			self.start = self.CellElement(self, 'start', dc['start'].get('value', 'single'), dc['start'].get('sz', ''),
					dc['start'].get('space', ''), dc['start'].get('shadow', ''),
					dc['start'].get('color', ''))
		elif 'left' in dc.keys():
			self.start = self.CellElement(self, 'left', dc['left'].get('value', 'single'), dc['left'].get('sz', ''),
					dc['left'].get('space', ''), dc['left'].get('shadow', ''),
					dc['left'].get('color', ''))
		elif 'all' in dc.keys():
			self.start = self.CellElement(self, 'start', dc['all'].get('value', 'single'), dc['all'].get('sz', ''),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

		self.end = None
		if 'end' in dc.keys():
			self.end = self.CellElement(self, 'end', dc['end'].get('value', 'single'), dc['end'].get('sz', ''),
					dc['end'].get('space', ''), dc['end'].get('shadow', ''),
					dc['end'].get('color', ''))
		elif 'all' in dc.keys():
			self.end = self.CellElement(self, 'end', dc['all'].get('value', 'single'), dc['all'].get('sz', ''),
					dc['all'].get('space', ''), dc['all'].get('shadow', ''),
					dc['all'].get('color', ''))

		self.insideH = None
		if 'insideH' in dc.keys():
			self.insideH = self.CellElement(self, 'insideH', dc['insideH'].get('value', 'single'),
					dc['insideH'].get('sz', ''),
					dc['insideH'].get('space', ''), dc['insideH'].get('shadow', ''),
					dc['insideH'].get('color', ''))

		self.insideV = None
		if 'insideV' in dc.keys():
			self.insideH = self.CellElement(self, 'insideV', dc['insideV'].get('value', 'single'),
					dc['insideV'].get('sz', ''),
					dc['insideV'].get('space', ''), dc['insideV'].get('shadow', ''),
					dc['insideV'].get('color', ''))

		self.tl2br = None
		if 'tl2br' in dc.keys():
			self.tl2br = self.CellElement(self, 'tl2br', dc['tl2br'].get('value', 'single'), dc['tl2br'].get('sz', ''),
					dc['tl2br'].get('space', ''), dc['tl2br'].get('shadow', ''),
					dc['tl2br'].get('color', ''))

		self.tr2bl = None
		if 'tr2bl' in dc.keys():
			self.tl2br = self.CellElement(self, 'tr2bl', dc['tr2bl'].get('value', 'single'), dc['tr2bl'].get('sz', ''),
					dc['tr2bl'].get('space', ''), dc['tr2bl'].get('shadow', ''),
					dc['tr2bl'].get('color', ''))

	def SetTop(self, value, sz='', space='', shadow=False, color=''):
		self.top = TableElement(self, 'top', value, sz, space, shadow, color)

	def SetBottom(self, value, sz='', space='', shadow=False, color=''):
		self.bottom = TableElement(self, 'bottom', value, sz, space, shadow, color)

	def SetStart(self, value, sz='', space='', shadow=False, color=''):
		self.start = TableElement(self, 'start', value, sz, space, shadow, color)

	def SetEnd(self, value, sz='', space='', shadow=False, color=''):
		self.end = TableElement(self, 'end', value, sz, space, shadow, color)

	def SetInsideH(self, value, sz='', space='', shadow=False, color=''):
		self.insideH = TableElement(self, 'insideH', value, sz, space, shadow, color)

	def SetInsideV(self, value, sz='', space='', shadow=False, color=''):
		self.insideV = TableElement(self, 'insideV', value, sz, space, shadow, color)

	def SetTl2br(self, value, sz='', space='', shadow=False, color=''):
		self.tl2br = TableElement(self, 'tl2br', value, sz, space, shadow, color)

	def SetTr2bl(self, value, sz='', space='', shadow=False, color=''):
		self.tr2bl = TableElement(self, 'tr2bl', value, sz, space, shadow, color)

	def get_Top(self):
		return self.top

	def get_Bottom(self):
		return self.bottom

	def get_Start(self):
		return self.start

	def get_End(self):
		return self.end

	def get_InsideH(self):
		return self.insideH

	def get_InsideV(self):
		return self.insideV

	def get_Tl2br(self):
		return self.tl2br

	def get_Tr2bl(self):
		return self.tr2bl

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		value = list()

		if self.top is not None:
			value.append(self.top.get_xml())
		if self.bottom is not None:
			value.append(self.bottom.get_xml())
		if self.start is not None:
			value.append(self.start.get_xml())
		if self.end is not None:
			value.append(self.end.get_xml())
		if self.insideH is not None:
			value.append(self.insideH.get_xml())
		if self.insideV is not None:
			value.append(self.insideV.get_xml())
		if self.tl2br is not None:
			value.append(self.tl2br.get_xml())
		if self.tr2bl is not None:
			value.append(self.tr2bl.get_xml())

		if value:
			value.insert(0, '%s<w:%s>' % (self.get_tab(), self.name))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))

		return self.separator.join(value)

