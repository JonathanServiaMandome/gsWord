#!/usr/bin/python
# -*- coding: utf-8 -*-

import element
import paragraph
import text
import rtf


class TableProperties(object):
	def __init__(self, tabla, bordes=None):
		self.name = 'tblPr'
		self.parent = tabla
		self.tab = tabla.tab
		self.separator = tabla.get_separator()
		self.indent = tabla.indent + 1

		self.shd = None
		self.tblBorders = None
		if bordes is not None:
			self.tblBorders = element.TableBorders(self, bordes)

		self.tblCaption = ''
		self.tblCellMar = None
		self.tblCellSpacing = None

		self.tblInd = None
		self.tblLayout = None
		'''Tables can be conditionally formatted based on such things as whether the content is in the first row, 
		last row, first column, or last column, or whether the rows or columns are to be banded (i.e., formatting 
		based on how the previous row or column was formatted). Such conditional formatting for tables is defined 
		in the referenced style for the table. See Table Styles. So a table may specify its style as 
		LightShading-Accent3 (<w:tblStyle w:val="LightShading-Accent3"/>). Inside LightShading-Accent3 may be a 
		style for the first row (<w:tblStylePr w:type="firstRow"/>). A given table which references the 
		LightShading-Accent3 may or may not use that style for the first row based on the attributes for the 
		<w:tblLook> element within the <w:tblPr> element. Since the firstRow attribute is set to true, 
		the conditional formatting will be applied. '''
		self.tblLook = None

		'''Since floating tables are outside of the normal flow of text and can be postiioned absolutely, 
		the potential exists for multiple floating tables to overlap. However, overlapping can be prevented with 
		the <w:tblOverlap> element within the <w:tblPr> element. This element specifies whether the table allows 
		other floating tables to overlap it. <w:tblOverlap> has just one attribute, val, with the following 
		possible values: never - the parent table cannot be displayed where it would overlap with another floating 
		table overlap - the parent table can be displayed where it overlaps another floating table '''
		self.tblOverlap = ''
		'''A floating table is a table which is not part of the main text flow in the document but is instead 
		absolutely positioned with a specific size and position relative to non-frame content in the document. A 
		floating table is specified with the <w:tblpPr> element within the <w:tblPr> element. Note: Positioning of 
		the table is relative to its top-left corner. Anchors (e.g., margin, page, or text) are specified in 
		attributes (tblpX and tblpY), from which measurements (also specified in attributes) for placement of the 
		table are specified. Distance from surrounding text can also be specified in other attributes. '''
		self.tblpPr = None

		'''Specifies the style ID of the table style if a table style is to be used to format the table. See 
		Defining a Style - Table Styles. Table styles are applied after document default styles, so they override 
		the defaults, but all other style types are applied after the table styles, so they override the table 
		style. In the sample below, a table style is referenced, but there is also direct formatting applied for 
		the table justification. '''
		self.tblStyle = 'Tablanormal'

		'''This element is applicable to a <tblPr> within a table style. Tables styles can have conditional 
		formatting which enables columns and/or rows to be "banded" by applying different formatting to 
		alternating columns or rows. Typically every other column is formatted the same. So, for example, 
		odd columns may have gray shading and even columns may have green shading. This element can be used to 
		group columns so that alternating groups of columns are formatted the same way rather than every other 
		column. For example, by setting the value of the val attribute to 3, the first three columns will be 
		formatted the same, and the second group of three will be formatted the same, and each group of three 
		thereafter will alternate their formatting. '''
		self.tblStyleColBandSize = ''

		'''This element is applicable to a <tblPr> within a table style. Tables styles can have conditional 
		formatting which enables columns and/or rows to be "banded" by applying different formatting to 
		alternating columns or rows. Typically every other row is formatted the same. So, for example, 
		odd rows may have gray shading and even rows may have green shading. This element can be used to group 
		rows so that alternating groups of rows are formatted the same way rather than every other row. For 
		example, by setting the value of the val attribute to 3, the first three rows will be formatted the same, 
		and the second group of three will be formatted the same, and each group of three thereafter will 
		alternate their formatting. '''
		self.tblStyleRowBandSize = ''

		'''The preferred width for the table is specified with the <w:tblW> element within the <w:tblPr> element. 
		All widths in a table are considered preferred because widths of columns can conflict and the table layout 
		rules can require a preference to be overridden. If this element is omitted, then the cell width is of 
		type auto. '''
		self.tblW = None

		self.paragraphSpacing = dict()

	def get_paragraph_spacing(self):
		return self.paragraphSpacing

	def set_paragraph_spacing(self, dc):
		for key in dc.keys():
			self.paragraphSpacing[key] = dc[key]

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_separator(self):
		return self.separator

	def get_table_width(self):
		return self.tblW

	def set_table_width(self, w, type_='dxa'):
		self.tblW = element.CellWidth(self, w, type_)

	def get_table_style_row_band_size(self):
		return self.tblStyleRowBandSize

	def set_table_style_row_band_size(self, value):  # ?
		self.tblStyleRowBandSize = value

	def get_table_style_col_band_size(self):
		return self.tblStyleColBandSize

	def set_table_style_col_band_size(self, value):  # ?
		self.tblStyleColBandSize = value

	def get_table_style(self):
		return self.tblStyle

	def set_table_style(self, value):
		self.tblStyle = value

	def get_floating_table(self):
		return self.tblpPr

	def get_table_overlap(self):
		return self.tblOverlap

	def set_table_overlap(self, value):
		self.tblOverlap = value

	def get_table_look(self):
		return self.tblLook

	def set_table_look(self, value='04A0', first_row='1', last_row='0', first_column='1', last_column='0',
						nohband='0', novband='1'):  # ??
		self.tblLook = element.TableLook(self, value, first_row, last_row, first_column, last_column, nohband,
											novband)

	def get_table_layout(self):
		return self.tblLayout

	def set_table_layout(self, value):
		self.tblLayout = element.TableLayout(self, value)

	def get_table_indentation(self):
		return self.tblInd

	def set_table_indentation(self, w, type_='dxa'):
		self.tblInd = element.TableIndentation(self, w, type_)

	def get_borders(self):
		return self.tblBorders

	def set_borders(self, dc):  # ok
		self.tblBorders = element.TableBorders(self, dc)

	def get_spacing(self):
		return self.tblCellSpacing

	def set_spacing(self, value, type_='dxa'):  # ok
		self.tblCellSpacing = element.CellSpacing(self, value, type_)

	def get_shading(self):
		return self.shd

	def set_shading(self, value, fill='', color=''):  # ok
		self.shd = element.Shading(self, value, fill, color)

	def get_cell_margin(self):
		return self.tblCellMar

	def set_cell_margin(self, dc):  # ok
		self.tblCellMar = element.CellMargin(self, dc)

	def get_table_caption(self):
		return self.tblCaption

	def set_table_caption(self, value):  # No funciona
		self.tblCaption = value

	def set_text_direction(self, value):
		for row in self.parent.rows:
			p = row.get_properties()
			p.set_text_direction(value)

	def get_text_direction(self):
		result = list()
		for row in self.parent.rows:
			p = row.get_properties()
			result.append(p.get_text_direction())
		return result

	def get_xml(self):
		value = list()
		if self.tblStyle:
			value.append('%s<w:tblStyle w:val="%s"/>' % (self.get_tab(1), self.tblStyle))
		if self.tblW is not None:
			value.append(self.tblW.get_xml())

		if self.tblCellSpacing is not None:
			value.append(self.tblCellSpacing.get_xml())
		if self.tblBorders is not None:
			value.append(self.tblBorders.get_xml())
		if self.shd is not None:
			value.append(self.shd.get_xml())
		if self.tblCellMar is not None:
			value.append(self.tblCellMar.get_xml())
		if self.tblLook is not None:
			value.append(self.tblLook.get_xml())
		if self.tblCaption:
			value.append('%s<w:tblCaption w:val="%s"/>' % (self.get_tab(1), self.tblCaption))

		if self.tblInd is not None:
			value.append(self.tblInd.get_xml())
		if self.tblLayout is not None:
			value.append(self.tblLayout.get_xml())
		if self.tblOverlap:
			value.append('%s<w:tblOverlap w:val="%s"/>' % (self.get_tab(1), self.tblCaption))
		if self.tblpPr is not None:
			value.append(self.tblpPr.get_xml())
		if self.tblStyleColBandSize:
			value.append('%s<w:tblStyleColBandSize w:val="%s"/>' % (self.get_tab(1), self.tblStyleColBandSize))
		if self.tblStyleRowBandSize:
			value.append('%s<w:tblStyleRowBandSize w:val="%s"/>' % (self.get_tab(1), self.tblStyleRowBandSize))

		if value:
			value.insert(0, '%s<w:%s>' % (self.get_tab(), self.name))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))

		return self.get_separator().join(value)


class Table(object):
	"""Class: Table
Atributtes:
		name:
		tblGrid:
		id:
		parent:
		tab:
		indent:
		separator:
		rows:
		properties:
		self.titulos"""

	def __init__(self, parent, idx, data=(), titulos=(), column_width=None, horizontal_alignment=(), borders=None):

		self.name = 'tbl'
		self.tblGrid = column_width
		self.id = idx
		self.parent = parent
		self.tab = parent.tab
		self.indent = parent.indent + 1
		self.separator = parent.get_separator()

		self.rows = []
		self.properties = TableProperties(self, borders)

		self._break = None
		self.titles = titulos
		if self.titles:
			self.rows.insert(0, Row(self, self.titles, ['c'] * len(titulos), is_header=True))

		if type(data) == str:
			data = [data]

		for k in range(len(data)):
			txt = data[k]

			if getattr(txt, 'name', '') == 'tr':
				txt.get_properties().set_width_cells(self.tblGrid)
				txt.set_horizontal_alignment(horizontal_alignment)

				self.rows.append(txt)
			elif getattr(txt, 'name', '') == 'p' or type(txt) == str:
				r = Row(self, [txt], horizontal_alignment)
				r.get_properties().set_width_cells(self.tblGrid)

				self.rows.append(r)
			elif type(txt) in [str, list]:
				if type(txt) == str:
					txt = [txt]
				r = Row(self, txt, horizontal_alignment)
				r.get_properties().set_width_cells(self.tblGrid)

				self.rows.append(r)
			else:
				raise ValueError(str(txt))

	def get_break(self):
		return self._break

	def set_break(self, _type='page'):
		self._break = _type

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_separator(self):
		return self.separator

	def get_parent(self):
		return self.parent

	def get_rows(self):
		return self.rows

	def get_row(self, index):
		return self.rows[index]

	def add_row(self):
		self.rows.append(Row(self, []))
		return self.rows[-1]

	def get_cell(self, row_index, cell_index):
		return self.get_row(row_index)[cell_index]

	def set_borders(self, dc):
		self.get_properties().set_borders(dc)

	def set_rows(self, rows):
		self.rows = rows

	def set_row(self, index, row):
		self.rows[index] = row

	def get_titles(self):
		return self.titles

	def set_vertical_merge(self, start, end):
		self.get_row(start[0]).get_cell(start[1]).get_properties().set_vertical_merge('start')
		self.get_row(end[0]).get_cell(end[1]).get_properties().set_vertical_merge(None)

	def set_paragraph_spacing(self, dc):
		self.get_properties().set_paragraph_spacing(dc)

	def set_cell_margin(self, dc):
		self.get_properties().set_cell_margin(dc)

	def set_titles(self, titulos):
		self.titles = titulos

	def set_font_size(self, size):
		for _row in self.get_rows():
			for _cell in _row.get_cells():
				for _paragraph in _cell.get_paragraphs():
					for _text in _paragraph.get_texts():
						_text.get_properties().set_font_size(size)

	def get_name(self):
		return self.name

	def get_id(self):
		return self.id

	def get_properties(self):
		return self.properties

	def get_columns_width(self):
		return self.tblGrid

	def add_paragraph(self, _text=(), horizontal_alignment='j', font_format='', font_size=None, is_null=False):
		x = self
		while getattr(x, 'tag', '') != 'document':
			x = x.get_parent()
		x.idx += 1
		p = paragraph.Paragraph(self, x.idx, _text, horizontal_alignment, font_format, font_size, nulo=is_null)

		r = Row(self, [p], horizontal_alignment)
		r.get_properties().set_width_cells(self.tblGrid)
		self.rows.append(r)
		return p, len(self.rows) - 1

	def set_columns_width(self, value):
		for row in self.get_rows():
			row.get_properties().set_width_cells(value)
		self.tblGrid = value

	def set_horizontal_alignment(self, value):
		for row in self.get_rows():
			for i in range(len(row.get_cells())):
				if len(value) > i:
					ha = value[i]
				else:
					ha = 'l'

				row.get_cell(i).get_properties().set_horizontal_alignment(ha)

	def set_foreground_colour(self, color):
		for row in self.get_rows():
			row.set_foreground_colour(color)

	def set_spacing(self, dc):
		for _row in self.get_rows():
			for _cell in _row.get_cells():
				for _element in _cell.elements:
					if hasattr(_element, 'set_spacing'):
						_element.set_spacing(dc)

	def set_font_format(self, dc):
		for key in dc.keys():
			if len(key) == 1:
				self.get_row(key[0]).set_font_format(dc[key])
			elif len(key) == 2:
				self.get_row(key[0]).get_cell(key[1]).set_font_format(dc[key])

	def get_table_grid(self):
		value = list()
		if self.tblGrid is not None:
			for w in self.tblGrid:
				if type(w) == str and '%' in w:
					w = int(w.replace('%', '').strip())
					section_width = self.get_parent().get_parent().get_body().get_active_section().get_width()
					section_width -= self.get_parent().get_parent().get_body().get_active_section().get_margin_left()
					section_width -= self.get_parent().get_parent().get_body().get_active_section().get_margin_rigth()

					w = int(section_width * (w / 100.))
				value.append('%s<w:gridCol w:w="%s"/>' % (self.get_tab(2), str(w)))
			if value:
				value.insert(0, '%s<w:tblGrid>' % self.get_tab(1))
				value.append('%s</w:tblGrid>' % self.get_tab(1))
		return self.get_separator().join(value)

	def get_xml(self):
		value = list()
		if self.get_break():
			value.append('%s<w:p>' % self.get_tab())
			value.append('%s<w:r>' % self.get_tab(1))
			value.append('%s<w:br w:type="page"/>' % self.get_tab(2))
			value.append('%s</w:r>' % self.get_tab(1))
			value.append('%s</w:p>' % self.get_tab())

		value.append('%s<w:tbl>' % self.get_tab())
		if self.properties is not None:
			value.append(self.properties.get_xml())

		table_grid = self.get_table_grid()
		if table_grid:
			value.append(table_grid)

		for row in self.rows:
			row.get_properties().set_paragraph_spacing(self.get_properties().get_paragraph_spacing())
			value.append(row.get_xml())

		value.append('%s</w:tbl>' % self.get_tab())
		return self.get_separator().join(value)


class CellProperties(object):
	def __init__(self, cell, horizontal_alignment='l', bordes=None):
		self.name = 'tcPr'
		self.parent = cell
		self.tab = cell.tab
		self.separator = cell.separator
		self.indent = cell.indent + 1
		self.jc = horizontal_alignment

		'''This element defines the number of logical columns across which the cell spans. It has a single 
		attribute w:val. See the discussion of <w:tblGrid> at Table Grid/Column Definition. '''
		self.gridSpan = ''

		'''he height of a row typically cannot be reduced below the size of the end-of-cell marker. This prevents 
		table rows from disappearing when they have no content. However, this makes it impossible to use a row as 
		a border by shading its cells or putting an image in the cells. To overcome this problem, 
		use the <w:hideMark /> element within the <w:tcPr /> of each cell in the row. This specifies that the 
		end-of-cell marker should be ignored for the cell, allowing it to collapse to the height of its contents 
		without formatting the cell's end-of-cell marker. '''
		self.hideMark = False

		'''This element will prevent text from wrapping in the cell under certain conditions. It is a boolean 
		property. E.g., <w:noWrap/>. Note: This only affects the behavior of the cell when the tblLayout for the 
		row is set to auto. If the cell width is fixed, then noWrap specifies that the cell will not be smaller 
		than that fixed value when other cells in the row are not at their minimum. If the cell width is set to 
		auto or pct, then the content of the cell will not wrap. '''
		self.noWrap = False

		'''Element.Shading'''
		self.shading = None

		'''Element.CellBorders'''
		self.tcBorders = bordes

		'''Text within a cell can be contracted or expanded to fit the width of the cell using the <w:tcFitText 
		w:val="true"/>. It has a single attribute val with boolean values of true/false. '''
		self.tcFitText = False

		'''Margins for a single table cell are specified with the <w:tcMar> element within the <tcPr> element. 
		This overrides the table-level cell margins, if present. Element.CellMargin '''
		self.tcMar = None

		'''Element.CellWidth'''
		self.tcW = None

		'''Specifies the vertical alignment for text between the top and bottom margins of the cell. Possible values are:
			bottom - Specifies that the text should be vertically aligned to the bottom margin.
			center - Specifies that the text should be vertically aligned to the center of the cell.
			top - Specifies that the text should be vertically aligned to the top margin.'''
		self.vAlign = ''

		'''This element specifies that the cell is part of a vertically merged set of cells. defines the number of 
		logical columns across which the cell spans. It has a single attribute w:val which specifies how the cell 
		is part of a vertically merged region. The cell can be part of an existing group of merged cells or it can 
		start a new group of merged cells. Possible values are: continue -- the current cell continues a 
		previously existing merge group restart -- the corrent cell starts a new merge group If omitted, 
		the value is assumed to be continue. See the discussion of <w:tblGrid> at Table Grid/Column Definition. '''
		self.vMerge = ''

		self.textDirection = None

		self.paragraphSpacing = dict()

	def get_paragraph_spacing(self):
		return self.paragraphSpacing

	def set_paragraph_spacing(self, dc):
		for key in dc.keys():
			self.paragraphSpacing[key] = dc[key]

	def get_parent(self):
		return self.parent

	def get_name(self):
		return self.name

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_separator(self):
		return self.separator

	def get_vertical_merge(self):
		return self.vMerge

	def set_vertical_merge(self, value):
		self.vMerge = value

	def get_text_direction(self):
		return self.textDirection

	def set_text_direction(self, value):
		self.textDirection = value

	def get_vertical_alignment(self):
		return self.vAlign

	def set_vertical_alignment(self, value):
		self.vAlign = value

	def get_horizontal_alignment(self):
		return self.jc

	def set_horizontal_alignment(self, value):
		self.jc = value

	def is_cell_fit_text(self):
		return self.tcFitText

	def set_cell_fit_text(self, value):
		self.tcFitText = value

	def get_cell_borders(self):
		return self.tcBorders

	def set_cell_borders(self, dc):
		self.tcBorders = element.CellBorders(self, dc)

	def get_shading(self):
		return self.shading

	def set_shading(self, value, fill='', color='auto'):
		self.shading = element.Shading(self, value, fill, color)

	def get_grid_span(self):
		return self.gridSpan

	def set_grid_span(self, value):
		self.gridSpan = value

	def is_hide_mark(self):
		return self.hideMark

	def hide_mark(self):
		self.hideMark = False

	def show_mark(self):
		self.hideMark = True

	def is_wrap(self):
		return self.hideMark

	def set_wrap(self, value):
		self.hideMark = value

	def get_CellMargin(self):
		return self.tcMar

	def set_cell_margin(self, dc):
		self.tcMar = element.CellMargin(self, dc)

	def get_table_cell_width(self):
		return self.tcW

	def set_table_cell_width(self, w, type_='dxa'):
		if type(w) == str and '%' in w:
			w = float(w.replace('%', '').strip())
			p = self
			while not hasattr(p, 'get_active_section'):
				if hasattr(p, 'get_parent'):
					p = p.get_parent()
				else:
					p = p.get_body()

			section_width = p.get_active_section().get_width()
			section_width -= p.get_active_section().get_margin_left()
			section_width -= p.get_active_section().get_margin_rigth()

			w = int(section_width * (w / 100.))
		self.tcW = element.CellWidth(self, int(w), type_)

	def get_xml(self):
		value = []
		if self.gridSpan:
			value.append('%s<w:gridSpan w:val="%s"/>' % (self.get_tab(1), self.gridSpan))
		if self.hideMark:
			value.append('%s<w:hideMark />' % self.get_tab(1))
		if self.noWrap:
			value.append('%s<w:noWrap />' % self.get_tab(1))

		if self.textDirection is not None:
			value.append('%s<w:textDirection w:val="%s"/>' % (self.get_tab(1), self.textDirection))

		if self.shading is not None:
			value.append(self.shading.get_xml())
		if self.tcBorders is not None:
			value.append(self.tcBorders.get_xml())
		if self.tcFitText:
			value.append('%s<w:tcFitText val="true"/>' % self.get_tab(1))
		if self.tcMar is not None:
			value.append(self.tcMar.get_xml())

		if self.tcW is not None:
			value.append(self.tcW.get_xml())

		if self.vAlign:
			value.append('%s<w:vAlign w:val="%s"/>' % (self.get_tab(1), self.vAlign))

		if self.vMerge is None:
			value.append('%s<w:vMerge/>' % (self.get_tab(1)))
		elif self.vMerge:
			value.append('%s<w:vMerge w:val="%s"/>' % (self.get_tab(1), self.vMerge))

		if value:
			value.insert(0, '%s<w:%s>' % (self.get_tab(), self.name))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))
			return self.get_separator().join(value)
		else:
			return ''


class Cell(object):

	def __init__(self, row, cell_number, data, horizontal_alignment='l'):
		self.name = 'tc'

		self.id = ''
		self.row = row
		self.tab = row.tab
		self.separator = row.separator
		self.indent = row.indent + 1
		self.properties = CellProperties(self, horizontal_alignment)
		self.cell_number = cell_number

		self.elements = list()

		if isinstance(data, Table):
			self.elements.append(data)
		elif getattr(data, 'name', '') == 'sdt':
			data.parent = self
			self.elements.append(data)
		else:
			if type(data) == str:
				data = [data]
			elif type(data) in [int, float]:
				data = [str(data)]
			if data is not None:
				for k in range(len(data)):
					txt = data[k]

					if isinstance(txt, paragraph.Paragraph):
						if txt.get_parent() is None:
							txt.SetParent(self)
						txt.set_horizontal_alignment(horizontal_alignment)
						self.elements.append(txt)
					else:
						self.add_paragraph(txt, horizontal_alignment)
					# self.elements.append(Paragraph.Paragraph(self, -1, txt, horizontal_alignment))

	def add_paragraph(self, _text=(), horizontal_alignment='j', font_format='', font_size=None, is_null=False):
		x = self
		while getattr(x, 'tag', '') != 'document':
			x = x.get_parent()
		x.idx += 1

		p = paragraph.Paragraph(self, x.idx, _text, horizontal_alignment, font_format, font_size, nulo=is_null)

		self.elements.append(p)
		return p

	def set_font_size(self, size):
		for _paragraph in self.elements:
			for _text in _paragraph.get_texts():
				_text.get_properties().set_font_size(size)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_parent(self):
		return self.row

	def get_separator(self):
		return self.separator

	def get_name(self):
		return self.name

	def get_row(self):
		return self.row

	def get_paragraphs(self):
		return self.elements

	def set_paragraphs(self, paragraphs):
		self.elements = paragraphs

	def get_paragraph(self, index):
		return self.elements[index]

	def set_paragraph(self, index, _paragraph):
		self.elements[index] = _paragraph

	def set_cell_borders(self, dc):
		self.get_properties().set_cell_borders(dc)

	def set_vertical_alignment(self, value):
		self.get_properties().set_vertical_alignment(value)

	def get_properties(self):
		return self.properties

	def set_properties(self, value):
		self.properties = value

	def cell_number(self):
		return self.cell_number

	def get_id(self):
		return self.id

	def set_id(self, value):
		self.id = value

	def add_element(self, _element):
		self.elements.append(_element)

	def set_element(self, index, _element):
		self.elements[index] = _element

	def set_font_format(self, font_format):
		for k in range(len(self.elements)):
			_element = self.elements[k]
			ff = ''
			if type(font_format) is str:
				ff = font_format
			elif type(font_format) is list:
				if len(font_format) > k:
					ff = font_format[k]

			for _text in _element.elements:
				if getattr(_text, 'name', '') == 'r':
					if 'b' in ff:
						_text.get_properties().bold()
					if 'i' in ff:
						_text.get_properties().italic()
					if 'u' in ff:
						_text.get_properties().underline()

	def add_rtf(self, rtf_text):
		_rtf = rtf.Rtf(text=rtf_text, parent=self)
		_rtf.get_value('word')

	def add_table(self, data=(), titles=(), column_width=(), horizontal_alignment=None, borders=None):
		_parent = self.get_row().get_table().get_parent().get_parent()
		while not hasattr(_parent, 'idx'):
			_parent = _parent.get_parent()

		_parent.idx += 1
		idx = _parent.idx
		table = Table(self, idx, data, titles, column_width, horizontal_alignment, borders)
		self.elements.append(table)
		return table

	def get_xml(self):
		value = list()
		id_ = ''
		if self.id:
			id_ = self.id
		value.append('%s<w:%s%s>' % (self.get_tab(), self.name, id_))
		if self.properties is not None:
			pr = self.properties.get_xml()
			if pr:
				value.append(pr)

		if not self.elements:
			raise ValueError('Celda sin elementos')
		for _element in self.elements:
			if getattr(_element, 'name', '') == 'p':
				paragraph_spacing = self.get_properties().get_paragraph_spacing()
				if _element.get_properties().get_spacing() is None:
					_element.get_properties().set_spacing(paragraph_spacing)
				else:
					if 'after' in paragraph_spacing.keys():
						_element.get_properties().get_spacing().set_after(paragraph_spacing['after'])
					if 'before' in paragraph_spacing.keys():
						_element.get_properties().get_spacing().set_before(paragraph_spacing['before'])
					if 'line' in paragraph_spacing.keys():
						_element.get_properties().get_spacing().set_line(paragraph_spacing['line'])
					if 'lineRule' in paragraph_spacing.keys():
						_element.get_properties().get_spacing().set_line_rule(paragraph_spacing['lineRule'])
					if 'beforeAutospacing' in paragraph_spacing.keys():
						_element.get_properties().get_spacing().set_before_autospacing(
							paragraph_spacing['beforeAutospacing'])
					if 'afterAutospacing' in paragraph_spacing.keys():
						_element.get_properties().get_spacing().set_after_autospacing(paragraph_spacing['afterAutospacing'])

			value.append(_element.get_xml())
			if isinstance(_element, Table):
				value.append('%s<w:p/>' % self.get_tab())
		value.append('%s</w:%s>' % (self.get_tab(), self.name))

		return self.get_separator().join(value)


class RowProperties(object):

	class RowHeight(object):
		def __init__(self, cell, alto, h_rule='exact'):
			self.name = 'trHeight'
			self.val = alto
			self.hRule = h_rule
			self.cell = cell
			self.tab = cell.tab
			self.separator = cell.separator
			self.indent = cell.indent + 1

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_separator(self):
			return self.separator

		def set_height(self, value):
			self.val = value

		def set_h_rule(self, value):
			self.hRule = value

		def get_height(self):
			return self.val

		def get_h_rule(self):
			return self.hRule

		def get_xml(self):
			value = ''

			if self.hRule:
				value += ' w:hRule="%s"' % self.hRule
			if self.val:
				value += ' w:val="%s"' % self.val
			if value:
				value = '%s<w:%s%s/>' % (self.get_tab(), self.name, value)

			return value

	class TableCellSpacing(object):
		def __init__(self, row, espaciado, tipo='dxa'):
			self.name = 'tblCellSpacing'
			self.w = espaciado
			self.type = tipo
			self.cell = row
			self.tab = row.tab
			self.indent = row.indent + 1
			self.separator = row.get_separator()

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_separator(self):
			return self.separator

		def set_cell_spacing(self, value):
			self.w = value

		def set_type(self, value):
			self.type = value

		def get_cell_spacing(self):
			return self.w

		def get_type(self):
			return self.type

		def get_name(self):
			return self.name

		def get_xml(self):
			value = ''
			if self.get_cell_spacing():
				value += ' w:w="%s"' % self.get_cell_spacing()
			if self.get_type():
				value += ' w:type="%s"' % self.get_type()
			if value:
				value = '%s<w:%s%s/>' % (self.get_tab(), self.get_name(), value)
			return value

	def __init__(self, row, cant_split=False, hidden=False, is_header=False):
		self.name = 'trPr'
		self.parent = row
		self.tab = row.tab
		self.indent = row.indent + 1
		self.separator = row.separator
		self.width_cells = None

		'''Elementos'''

		'''If present, it prevents the contents of the row from breaking across multiple pages by moving the start 
		of the row to the start of a new page. If the contents cannot fit on a single page, the row will start on 
		a new page and flow onto multiple pages. '''
		self.cantSplit = cant_split
		'''Hides the entire row. (Note, however, that applications can have settings that allow hidden content to 
		be displayed.) '''
		self.hidden = hidden

		'''Specifies the default cell spacing (space between adjacent cells and the edges of the table) for cells 
		in the row. E.g., <w:tblCellSpacing w:w="144" w:type="dxa"/> The attributes are: w -- Specifies the width 
		of the spacing. type -- Specifies the units of the width attribute above. If omitted, the value is assumed 
		to be dxa (twentieths of a point). Other possible values are nil. Note that you can specify auto or pct (
		width in percent of table width), but if you do, the width value will be ignored. '''
		self.tblCellSpacing = None

		'''Specifies that the current row should be repeated at the top each new page on which the table is 
		displayed. E.g, <w:tblHeader />. This can be specified for multiple rows to generate a multi-row header. 
		Note that if the row is not the first row, then the property will be ignored. This is a simple boolean 
		property, so you can specify a val attribute of true or false, or no value at all (defalut value is true). 
		'''
		self.tblHeader = is_header

		'''Specifies the height of the row. E.g, <w:trHeight w:hRule="exact" w: val="2189" />. If omitted, 
		the row is automatically resized to fit the content. The attributes are: hRule -- Specifies the meaning of 
		the height. Possible values are atLeast (height should be at leasat the value specified), exact (height 
		should be exactly the value specified), or auto (default value--height is determined based on the height 
		of the contents, so the value is ignored). val -- Specifies the row's height, in twentieths of a point. 
		Note: The height of a row typically cannot be reduced below the size of the end of the cell marker. This 
		prevents table rows from disappearing when they have no content. However, this makes it impossible to use 
		a row as a border by shading its cells or putting an image in the cells. To overcome this problem, 
		use the <w:hideMark /> element within the <w:tcPr /> of each cell in the row. See Hide End of Table Cell 
		Mark '''
		self.trHeight = None

		self.paragraphSpacing = dict()

	def get_paragraph_spacing(self):
		return self.paragraphSpacing

	def set_paragraph_spacing(self, dc):
		for key in dc.keys():
			self.paragraphSpacing[key] = dc[key]

	def set_width_cells(self, value):
		self.width_cells = value
		for k in range(len(self.parent.get_cells())):
			if len(value) > k:
				self.parent.get_cell(k).get_properties().set_table_cell_width(value[k])

	def get_width_cells(self):
		return self.width_cells

	def get_tab(self, number=0):
		return self.tab * number

	def get_name(self):
		return self.name

	def get_separator(self):
		return self.separator

	def set_text_direction(self, value):
		for cell in self.parent.cells:
			p = cell.get_properties()
			p.set_text_direction(value)

	def get_text_direction(self):
		result = []
		for cell in self.parent.cells:
			p = cell.get_properties()
			result.append(p.get_text_direction())
		return result

	def get_xml(self):
		value = []

		if self.cantSplit:
			value.append('%s<w:cantSplit/>' % self.get_tab(1))
		if self.hidden:
			value.append('%s<w:hidden/>' % self.get_tab(1))

		if self.tblCellSpacing is not None:
			value.append(self.get_table_cell_spacing().get_xml())
		if self.trHeight is not None:
			value.append(self.trHeight.get_xml())
		if self.tblHeader:
			value.append('%s<w:tblHeader/>' % self.get_tab(1))

		if value:
			value.insert(0, '%s<w:%s>' % (self.get_tab(), self.get_name()))
			value.append('%s</w:%s>' % (self.get_tab(), self.get_name()))

			return self.get_separator().join(value)
		else:
			return ''

	def is_splittable(self):
		return self.cantSplit

	def splittable(self):
		self.cantSplit = True

	def not_splittable(self):  # ok
		self.cantSplit = False

	def is_hidden(self):
		return self.hidden

	def hidden(self):  # ??
		self.hidden = True

	def not_hidden(self):
		self.hidden = False

	def is_table_header(self):
		return self.tblHeader

	def table_header(self):  # ok
		self.tblHeader = True

	def not_table_header(self):
		self.tblHeader = False

	def set_row_height(self, alto, h_rule='exact'):  # ok
		self.trHeight = self.RowHeight(self, alto, h_rule)

	def get_row_height(self):
		return self.trHeight

	def set_table_cell_spacing(self, spacing, _type='dxa'):  # ox
		self.tblCellSpacing = self.TableCellSpacing(self, spacing, _type)

	def get_table_cell_spacing(self):
		return self.tblCellSpacing


class Row(object):

	def __init__(self, table, data, horizontal_alignment=(), is_header=False, cant_split=False):
		self.name = 'tr'
		self.table = table
		self.tab = table.tab
		self.separator = table.separator
		self.indent = table.indent + 1
		'''Atributos'''
		self.rsidDel = ''
		self.rsidR = ''
		self.rsidRPr = ''
		self.rsidTr = ''

		self.tblPrEx = None

		self.cells = list()

		if getattr(data, 'name', '') == 'p':
			if horizontal_alignment in [(), None]:
				horizontal_alignment = ['l']
			c = Cell(self, 0, [data], horizontal_alignment[0])
			self.cells.append(c)
		else:
			for k in range(len(data)):
				txt = data[k]

				jc = 'l'
				if horizontal_alignment in [(), None]:
					horizontal_alignment = ['l']
				elif len(horizontal_alignment) > k:
					jc = horizontal_alignment[k]

				if getattr(txt, 'name', '') == 'tc':
					txt.get_properties().set_horizontal_alignment(jc)
					self.cells.append(txt)
				elif getattr(txt, 'name', '') == 'p':
					c = Cell(self, k, [txt], jc)
					self.cells.append(c)
				else:
					self.cells.append(Cell(self, k, txt, jc))

		self.properties = RowProperties(self, cant_split=cant_split, is_header=is_header)

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_name(self):
		return self.name

	def get_separator(self):
		return self.separator

	def set_background_colour(self, color, mode='auto'):
		for cell in self.get_cells():
			cell.get_properties().set_shading(color, mode)

	def set_cell_borders(self, borders):
		for k in range(len(self.get_cells())):
			cell = self.get_cell(k)
			if type(borders) == list:
				cell.set_cell_borders(borders[k])
			else:
				cell.set_cell_borders(borders)

	def set_foreground_colour(self, color):
		for _cell in self.get_cells():
			for _element in _cell.elements:
				for _text in _element.elements:
					if getattr(_text, 'name', '') == 'r':
						_text.get_properties().set_color(color)

	def set_font_format(self, font_format):
		for k in range(len(self.get_cells())):
			_cell = self.get_cell(k)
			ff = ''
			if type(font_format) is str:
				ff = font_format
			elif type(font_format) is list:
				if len(font_format) > k:
					ff = font_format[k]
			for _element in _cell.elements:
				for _text in _element.elements:
					if getattr(text, 'name', '') == 'r':
						if 'b' in ff:
							_text.get_properties().bold()
						if 'i' in ff:
							_text.get_properties().italic()
						if 'u' in ff:
							_text.get_properties().underline()

	def set_row_height(self, height):
		self.get_properties().set_row_height(height)

	def set_vertical_alignment(self, value):
		k = 0
		for cell in self.get_cells():
			if type(value) == str:
				cell.set_vertical_alignment(value)
			else:
				cell.set_vertical_alignment(value[k])
			k += 1

	def set_font_size(self, size):
		for k in range(len(self.get_cells())):
			cell = self.get_cell(k)
			for _paragraph in cell.get_paragraphs():
				for _text in _paragraph.get_texts():
					_text.get_properties().set_font_size(size)

	def get_table(self):
		return self.table

	def get_properties(self):
		return self.properties

	def get_parent(self):
		return self.table

	def set_properties(self, value):
		self.properties = value

	def get_table_property_exceptions(self):
		return self.tblPrEx

	def set_horizontal_alignment(self, horizontal_alignment):
		for k in range(len(self.get_cells())):
			cell = self.get_cell(k)
			jc = 'l'
			if len(horizontal_alignment) > k:
				jc = horizontal_alignment[k]
			cell.get_properties().set_horizontal_alignment(jc)

	def set_table_property_exceptions(self, value):
		self.tblPrEx = value

	def set_rsid_del(self, value):
		self.rsidDel = value

	def get_rsid_del(self):
		return self.rsidDel

	def set_rsid_r(self, value):
		self.rsidR = value

	def get_rsid_r(self):
		return self.rsidR

	def set_rsid_r_pr(self, value):
		self.rsidRPr = value

	def get_rsid_r_pr(self):
		return self.rsidRPr

	def set_rsid_tr(self, value):
		self.rsidTr = value

	def get_rsid_tr(self):
		return self.rsidTr

	def get_cells(self):
		return self.cells

	def set_cells(self, cells):
		self.cells = cells

	def get_cell(self, index):
		return self.cells[index]

	def set_cell(self, index, cell):
		self.cells[index] = cell

	def add_cell(self, data, horizontal_alignment='l'):
		self.cells.append(Cell(self, len(self.cells) + 1, data, horizontal_alignment))
		return self.cells[-1]

	def get_xml(self):
		value = list()
		if self.properties is not None:
			value.append(self.properties.get_xml())
		if self.tblPrEx is not None:
			value.append(self.properties.get_xml())
		for cell in self.cells:
			cell.get_properties().set_paragraph_spacing(self.get_properties().get_paragraph_spacing())
			cc = cell.get_xml()
			if not cc:
				continue
			value.append(cc)

		if value:
			attributes = str()
			if self.rsidDel:
				attributes += ' w:rsidDel="%s"' % self.rsidDel
			if self.rsidR:
				attributes += ' w:rsidDel="%s"' % self.rsidR
			if self.rsidRPr:
				attributes += ' w:rsidDel="%s"' % self.rsidRPr
			if self.rsidTr:
				attributes += ' w:rsidDel="%s"' % self.rsidTr

			value.insert(0, '%s<w:%s%s>' % (self.get_tab(), self.name, attributes))
			value.append('%s</w:%s>' % (self.get_tab(), self.name))

			return self.get_separator().join(value)
		else:
			return ''


if __name__ == '__main__':
	print Table.__doc__
