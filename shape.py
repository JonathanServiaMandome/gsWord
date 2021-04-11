#!/usr/bin/python
# -*- coding: utf-8 -*-




class PrstTxWarp(object):
	def __init__(self, _parent, prst='textNoShape', av_lst=True):
		self.name = 'a:prstTxWarp'
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1
		self.prst = prst
		self.av_lst = av_lst

	def get_prst(self):
		return self.prst

	def set_prst(self, prst):
		self.prst = prst

	def get_av_lst(self):
		return self.av_lst

	def set_av_lst(self, av_lst):
		self.av_lst = av_lst

	def get_xml(self):
		value = list()
		tx = str()
		if self.get_prst():
			tx += ' prst="%s"' % self.get_prst()
		value.append('%s<%s%s>' % (self.get_tab(), self.get_name(), tx))

		if self.get_av_lst():
			value.append('%s<a:avLst/>' % self.get_tab(1))

		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name


class Shape_body(object):
	def __init__(self, _parent, auto_fit=True):
		self.name = 'wps:bodyPr'
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1
		self.attributes = {
			"rot": "0",
			"spcFirstLastPara": "0",
			"vertOverflow": "overflow",
			"horzOverflow": "overflow",
			"vert": "horz",
			"wrap": "square",
			"lIns": "91440",
			"tIns": "45720",
			"rIns": "91440",
			"bIns": "45720",
			"numCol": "1",
			"spcCol": "0",
			"rtlCol": "0",
			"fromWordArt": "0",
			"anchor": "t",
			"anchorCtr": "0",
			"forceAA": "0",
			"compatLnSpc": "1"
		}
		self.auto_fit = auto_fit
		self.elements = list()

	def get_elements(self):
		return self.elements

	def set_elements(self, elements):
		self.elements = elements

	def add_element(self, elements):
		self.elements.append(elements)

	def get_attributes(self):
		return self.attributes

	def get_attribute(self, name):
		return self.attributes[name]

	def set_attributes(self, dc):
		self.attributes = dc

	def get_auto_fit(self):
		return self.auto_fit

	def set_auto_fit(self, auto_fit):
		self.auto_fit = auto_fit

	def set_attribute(self, name, value):
		self.attributes[name] = value

	def get_xml(self):
		value = list()
		tx = str()
		for key in self.get_attributes():
			tx += ' %s="%s"' % (key, self.get_attribute(key))

		value.append('%s<%s%s>' % (self.get_tab(), self.get_name(), tx))

		for element in self.get_elements():
			value.append(element.get_xml())

		if self.get_auto_fit():

			value.append('%s<a:spAutoFit/>' % self.get_tab(1))
		else:
			value.append('%s<a:noAutofit/>' % self.get_tab(1))

		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name


class Xform(object):

	def __init__(self, parent, flip_horizontal, flip_vertical, offset, ext):
		self.name = 'a:xfrm'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.elements = list()
		self.flip_horizontal = flip_horizontal
		self.flip_vertical = flip_vertical
		self.offset = offset
		self.ext = ext

	def get_xml(self):
		value = list()
		tx = '%s<%s' % (self.get_tab(), self.get_name())
		if self.get_flip_horizontal():
			tx += ' flipH="%s"' % self.get_flip_horizontal()
		if self.get_flip_vertical():
			tx += ' flipV="%s"' % self.get_flip_vertical()
		tx += '>'
		value.append(tx)

		value.append('%s<a:off x="%s" y="%s"/>' % (self.get_tab(1), self.get_offset(0), self.get_offset(1)))
		value.append('%s<a:ext cx="%s" cy="%s"/>' % (self.get_tab(1), self.get_ext(0), self.get_ext(1)))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)

	def get_ext(self, pos):
		return self.ext[pos]

	def set_ext(self, x, y):
		self.ext = (x, y)

	def get_offset(self, pos):
		return self.offset[pos]

	def set_offset(self, x, y):
		self.offset = (x, y)

	def get_flip_vertical(self):
		return self.flip_vertical

	def set_flip_vertical(self, value):
		self.flip_vertical = value

	def get_flip_horizontal(self):
		return self.flip_horizontal

	def set_flip_horizontal(self, value):
		self.flip_horizontal = value

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name


class PrstGeom(object):

	def __init__(self, parent, prst, avlst):
		self.name = 'a:prstGeom'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.elements = list()
		self.avlst = avlst
		self.prst = prst

	def get_prst(self):
		return self.prst

	def set_prst(self, value):
		self.prst = value

	def get_avlst(self):
		return self.avlst

	def set_avlst(self, value):
		self.avlst = value

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		if self.get_avlst():
			value.append('%s<%s prst="%s">' % (self.get_tab(), self.get_name(), self.get_prst()))
			value.append('%s<a:avLst/>' % self.get_tab(1))
			value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Line(object):
	def __init__(self, parent, width=''):
		self.name = 'a:ln'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.elements = list()
		self.width = width

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_width(self):
		return self.width

	def set_width(self, width):
		self.width = width

	def get_xml(self):
		value = list()
		for element in self.elements:
			value.append(element.get_xml())

		if value:
			value.insert(0, '%s<%s w="%s">' % (self.get_tab(), self.get_name(), str(self.get_width())))
			value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class SolidFill(object):

	def __init__(self, parent, prst_clr='', scheme_clr=''):
		self.name = 'a:solidFill'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.prst_clr = prst_clr  # black
		self.scheme_clr = scheme_clr  # lt1

	def get_prst_clr(self):
		return self.prst_clr

	def set_prst_clr(self, value):
		self.prst_clr = value

	def get_scheme_clr(self):
		return self.scheme_clr

	def set_scheme_clr(self, value):
		self.scheme_clr = value

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		if self.get_scheme_clr():
			value.append('%s<a:schemeClr val="%s"/>' % (self.get_tab(1), self.get_scheme_clr()))
		if self.get_prst_clr():
			value.append('%s<a:prstClr val="%s"/>' % (self.get_tab(1), self.get_prst_clr()))

		if value:
			value.insert(0, '%s<%s>' % (self.get_tab(), self.get_name()))
			value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class ShapeProperties(object):
	def __init__(self, parent):
		self.name = 'wps:spPr'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.elements = list()

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		value.append('%s<%s>' % (self.get_tab(), self.get_name()))
		for element in self.elements:
			value.append(element.get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Shape(object):

	def __init__(self, parent, shape_type='txBox'):
		self.name = 'wps:wps'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.properties = ShapeProperties(self)
		self.elements = list()
		self.bodyPr = None
		self.cNvSpPr_value = "1"
		self.cNvSpPr_type = shape_type

	def get_cnv_sp_pr_type(self):
		return self.cNvSpPr_type

	def set_cnv_sp_pr_type(self, _type):
		self.cNvSpPr_type = _type

	def get_cnv_sp_pr_value(self):
		return self.cNvSpPr_value

	def set_cnv_sp_pr_value(self, value):
		self.cNvSpPr_value = value

	def get_elements(self):
		return self.elements

	def set_elements(self, elements):
		self.elements = elements

	def add_element(self, elements):
		self.elements.append(elements)

	def get_body_pr(self):
		return self.bodyPr

	def set_body_pr(self, body_pr):
		self.bodyPr = body_pr

	def get_properties(self):
		return self.properties

	def set_properties(self, properties):
		self.properties = properties

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		value.append('%s<%s>' % (self.get_tab(), self.get_name()))
		value.append(
			'%s<wps:cNvSpPr %s="%s"/>' % (self.get_tab(1), self.get_cnv_sp_pr_type(), self.get_cnv_sp_pr_value()))
		value.append(self.get_properties().get_xml())
		for element in self.get_elements():
			value.append(element.get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)

class TextBox(object):

	def __init__(self, parent):
		self.name = 'wps:txbx'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 2
		self.elements = list()

	def get_name(self):
		return self.name

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_xml(self):
		value = list()
		for element in self.elements:
			value.append(element.get_xml())

		if value:
			value.insert(0, '%s<w:txbxContent>' % self.get_tab())
			value.insert(0, '%s<%s>' % (self.get_tab(-1), self.get_name()))
			value.append('%s</w:txbxContent>' % self.get_tab())
			value.append('%s</%s>' % (self.get_tab(-1), self.get_name()))

		return self.separator.join(value)

	def add_paragraph(self, text=(), horizontal_alignment='j', font_format='', font_size=None, is_null=False):
		_x = self
		while getattr(_x, 'tag', '') != 'document':
			_x = _x.parent
		_x.idx += 1
		# p = paragraph.Paragraph(self, _x.idx, text, horizontal_alignment, font_format, font_size, nulo=is_null)
		# self.elements.append(p)
		# return p


class GraphicData(object):
	def __init__(self, parent):
		self.name = 'a:graphicData'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.shape = None
		self.uri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

	def get_uri(self):
		return self.uri

	def get_shape(self):
		return self.shape

	def set_shape(self, _shape):
		self.shape = _shape

	def set_uri(self, uri):
		self.uri = uri

	def get_name(self):
		return self.name

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_xml(self):
		value = list()
		value.append('%s<%s uri="%s">' % (self.get_tab(), self.get_name(), self.get_uri()))
		value.append(self.get_shape().get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))
		return self.separator.join(value)


class Graphic(object):
	def __init__(self, parent):
		self.name = 'a:graphic'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.graphicData = GraphicData(self)
		self.xmls = 'http://schemas.openxmlformats.org/drawingml/2006/main'

	def get_xmls(self):
		return self.xmls

	def set_xmls(self, _xmls):
		self.xmls = _xmls

	def get_graphic_data(self):
		return self.graphicData

	def set_graphic_data(self, _object):
		self.graphicData = _object

	def get_name(self):
		return self.name

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_xml(self):
		value = list()
		value.append('%s<%s xmlns:a="%s">' % (self.get_tab(), self.get_name(), self.get_xmls()))
		value.append(self.get_graphic_data().get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))
		return self.separator.join(value)


class SizeRelative(object):
	def __init__(self, parent, orientation='horizontal', relative_from='margin', value=0):
		self.orientation = orientation
		self.name = ''
		self.set_name(orientation)
		self.relative_from = relative_from
		self.value = value
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def set_name(self, orientation):
		if orientation is 'horizontal':
			self.name = 'wp14:sizeRelH'
		elif orientation is 'vertical':
			self.name = 'wp14:sizeRelV'
		else:
			raise ValueError("La orietacion tiene que ser horizontal o vertical")

	def get_value(self):
		return self.value

	def set_value(self, value):
		self.value = value

	def get_orientation(self):
		return self.orientation

	def set_orientation(self, orientation):
		self.orientation = orientation
		self.set_name(orientation)

	def get_relative_from(self):
		return self.relative_from

	def set_relative_from(self, relative_from):
		self.relative_from = relative_from

	def get_xml(self):
		value = list()
		value.append('%s<%s relativeFrom="%s">' % (self.get_tab(), self.get_name(), self.get_relative_from()))
		if self.get_orientation() is 'horizontal':
			value.append('%s<wp14:pctWidth>%s</wp14:pctWidth>' % (self.get_tab(1), self.get_value()))
		else:
			value.append('%s<wp14:pctHeight>%s</wp14:pctHeight>' % (self.get_tab(1), self.get_value()))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))
		return self.separator.join(value)


class Position(object):
	def __init__(self, parent, orientation='horizontal', relative_from='paragraph', align=None, position_offset=None):
		self.orientation = orientation
		self.name = ''
		self.set_name(orientation)
		self.relative_from = relative_from
		self.position_offset = position_offset
		self.align = align
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def set_name(self, orientation):
		if orientation is 'horizontal':
			self.name = 'wp:positionH'
		elif orientation is 'vertical':
			self.name = 'wp:positionV'
		else:
			raise ValueError("La orietacion tiene que ser horizontal o vertical")

	def get_align(self):
		return self.align

	def set_align(self, align):
		self.align = align

	def get_position_offset(self):
		return self.position_offset

	def set_position_offset(self, position_offset):
		self.position_offset = position_offset

	def get_orientation(self):
		return self.orientation

	def set_orientation(self, orientation):
		self.orientation = orientation
		self.set_name(orientation)

	def get_relative_from(self):
		return self.relative_from

	def set_relative_from(self, relative_from):
		self.relative_from = relative_from

	def get_xml(self):
		value = list()
		value.append('%s<%s relativeFrom="%s">' % (self.get_tab(), self.get_name(), self.get_relative_from()))
		if self.get_align() is not None:
			value.append('%s<wp:align>%s</wp:align>' % (self.get_tab(1), self.get_align()))
		if self.get_position_offset() is not None:
			value.append('%s<wp:posOffset>%s</wp:posOffset>' % (self.get_tab(1), str(self.get_position_offset())))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))
		return self.separator.join(value)


# noinspection PyShadowingNames
class Drawing(object):

	def __init__(self, parent, pos=(0, 0), size=(10, 20), size_relative=None):
		self.name = 'w:drawing'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 2

		self.anchor = {
			"distT": 0,
			"distB": 0,
			"distL": 114300,
			"distR": 114300,
			"simplePos": 0,
			"relativeHeight": 251659264,
			"behindDoc": 0,
			"locked": 0,
			"layoutInCell": 1,
			"allowOverlap": 1,
			"wp14:anchorId": "45C456EB",
			"wp14:editId": "0CFAEE44"
		}
		self.simplePos = {'x': pos[0], 'y': pos[1]}
		width = size[0] * 635
		height = size[1] * 635
		self.extent = {'cx': width, 'cy': height}
		self.effectExtent = {'l': 0, 't': 0, 'r': 0, 'b': 0}
		self.wrapSquare = None  # {'wrapText': 'bothSides'}
		_id = 1
		self.docPr = {'id': str(_id), 'name': "Shape %d" % _id}
		self.cNvGraphicFramePr = True
		self.graphic = Graphic(self)
		self.alternateContent = None  # AlternateContent(self, requires='wp14')
		self.position_horizontal = Position(self, 'horizontal', position_offset='0')
		self.position_vertical = Position(self, 'vertical', position_offset='0')

		self.size_relative_vertical = None
		self.size_relative_horizontal = None
		if size_relative and 'vertical' in size_relative.keys():
			SizeRelative(self, 'vertical', size_relative['vertical']['relative'], size_relative['vertical']['value'])
		if size_relative and 'horizontal' in size_relative.keys():
			SizeRelative(self, 'horizontal', size_relative['horizontal']['relative'],
						 size_relative['horizontal']['value'])

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_alternate_content(self):
		return self.alternateContent

	def set_alternate_content(self, _object):
		self.alternateContent = _object

	def get_position_horizontal(self):
		return self.position_horizontal

	def set_position_horizontal(self, _object):
		self.position_horizontal = _object

	def get_position_vertical(self):
		return self.position_vertical

	def set_position_vertical(self, _object):
		self.position_vertical = _object

	def get_size_relative_vertical(self):
		return self.size_relative_vertical

	def set_size_relative_vertical(self, size_relative_vertical):
		self.size_relative_vertical = size_relative_vertical

	def get_size_relative_horizontal(self):
		return self.size_relative_horizontal

	def set_size_relative_horizontal(self, size_relative_horizontal):
		self.size_relative_horizontal = size_relative_horizontal

	def get_anchor(self):
		return self.anchor

	def set_anchor(self, dc):
		self.anchor = dc

	def set_anchor_value(self, key, value):
		self.anchor[key] = value

	def get_simple_pos(self):
		return self.simplePos

	def set_simple_pos(self, dc):
		self.simplePos = dc

	def set_simple_pos_value(self, key, value):
		self.simplePos[key] = value

	def get_extent(self):
		return self.extent

	def set_extent(self, dc):
		self.extent = dc

	def set_extent_value(self, key, value):
		self.extent[key] = value

	def get_effect_extent(self):
		return self.effectExtent

	def set_effect_extent(self, dc):
		self.effectExtent = dc

	def set_effect_extent_value(self, key, value):
		self.effectExtent[key] = value

	def get_wrap_square(self):
		return self.wrapSquare

	def set_wrap_square(self, dc):
		self.wrapSquare = dc

	def set_wrap_square_value(self, key, value):
		self.wrapSquare[key] = value

	def get_doc_pr(self):
		return self.docPr

	def set_doc_pr(self, dc):
		self.docPr = dc

	def set_doc_pr_value(self, key, value):
		self.docPr[key] = value

	def get_cnv_graphic_frame_pr(self):
		return self.cNvGraphicFramePr

	def set_cnv_graphic_frame_pr(self, value):
		self.cNvGraphicFramePr = value

	def get_graphic(self):
		return self.graphic

	def set_graphic(self, graphic):
		self.graphic = graphic

	def get_xml(self):
		def add_to_value(_aux, _value):
			for ln in _aux:
				name, dic = ln
				tx = '%s<%s' % (self.get_tab(), name)
				for key in dic.keys():
					tx += ' %s="%s"' % (key, str(dic[key]))
				if name != 'wp:anchor':
					tx += '/'
				tx += '>'
				_value.append(tx)

		value = list()

		aux = list()
		if self.get_anchor():
			aux.append(['wp:anchor', self.get_anchor()])

		if self.get_simple_pos():
			aux.append(['wp:simplePos', self.get_simple_pos()])

		add_to_value(aux, value)

		if self.get_position_horizontal():
			value.append(self.get_position_horizontal().get_xml())

		if self.get_position_vertical():
			value.append(self.get_position_vertical().get_xml())

		if self.get_alternate_content():
			value.append(self.get_alternate_content().get_xml())

		aux = list()
		if self.get_extent():
			aux.append(['wp:extent', self.get_extent()])

		if self.get_effect_extent():
			aux.append(['wp:effectExtent', self.get_effect_extent()])

		if self.get_wrap_square():
			aux.append(['wp:wrapSquare', self.get_wrap_square()])
		else:
			value.append("%s<wp:wrapNone/>" % self.get_tab())

		if self.get_doc_pr():
			aux.append(['wp:docPr', self.get_doc_pr()])

		add_to_value(aux, value)

		if self.get_cnv_graphic_frame_pr():
			value.append("%s<wp:cNvGraphicFramePr/>" % self.get_tab())

		if self.get_graphic():
			value.append(self.get_graphic().get_xml())

		if self.get_size_relative_horizontal():
			value.append(self.get_size_relative_horizontal().get_xml())

		if self.get_size_relative_vertical():
			value.append(self.get_size_relative_vertical().get_xml())

		if self.get_anchor():
			value.append('%s</wp:anchor>' % self.get_tab(1))
		if value:
			value.insert(0, '%s<%s>' % (self.get_tab(-1), self.get_name()))
			value.append('%s</%s>' % (self.get_tab(-1), self.get_name()))

		return self.separator.join(value)


class Choice(object):
	def __init__(self, parent, requires, pos=(0, 0), size=(20, 10), size_relative=()):
		self.name = 'mc:Choice'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.requires = requires
		self.drawing = Drawing(self, pos, size, size_relative)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_requires(self):
		return self.requires

	def set_requires(self, requires):
		self.requires = requires

	def get_drawing(self):
		return self.drawing

	def set_drawing(self, drawing):
		self.drawing = drawing

	def get_xml(self):
		value = list()
		t = ''
		if self.get_requires():
			t = ' Requires="%s"' % self.get_requires()

		value.append('%s<%s%s>' % (self.get_tab(), self.get_name(), t))
		value.append(self.get_drawing().get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class AlternateContent(object):
	def __init__(self, parent, requires='wps', pos=(0, 0), size=(20, 10), size_relative=(), text=''):
		self.name = 'mc:AlternateContent'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.choice = Choice(self, requires, pos, size, size_relative)
		self.fallBack = None
		self.txt = text.decode('iso-8859-1').encode('utf8').replace('%EURO%', '€').replace('&', '&amp;').replace(
					'<', '&lt;').replace('>', '&gt;')

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_choice(self):
		return self.choice

	def get_name(self):
		return self.name

	def get_fall_back(self):
		return self.fallBack

	def get_xml(self):
		tx = """
				<mc:AlternateContent>
					<mc:Choice Requires="wps">
						<w:drawing>
							<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251660288" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="3454B481" wp14:editId="6B87A082">
								<wp:simplePos x="0" y="0"/>
								<wp:positionH relativeFrom="margin">
									<wp:posOffset>-3078955</wp:posOffset>
								</wp:positionH>
								<wp:positionV relativeFrom="paragraph">
									<wp:posOffset>3491230</wp:posOffset>
								</wp:positionV>
								<wp:extent cx="5514975" cy="323215"/>
								<wp:effectExtent l="5080" t="0" r="0" b="0"/>
								<wp:wrapSquare wrapText="bothSides"/>
								<wp:docPr id="3" name="Cuadro de texto 2"/>
								<wp:cNvGraphicFramePr>
									<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
								</wp:cNvGraphicFramePr>
								<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
									<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
										<wps:wsp>
											<wps:cNvSpPr txBox="1">
												<a:spLocks noChangeArrowheads="1"/>
											</wps:cNvSpPr>
											<wps:spPr bwMode="auto">
												<a:xfrm rot="16200000">
													<a:off x="0" y="0"/>
													<a:ext cx="5514975" cy="323215"/>
												</a:xfrm>
												<a:prstGeom prst="rect">
													<a:avLst/>
												</a:prstGeom>
												<a:solidFill>
													<a:srgbClr val="FFFFFF"/>
												</a:solidFill>
												<a:ln w="9525">
													<a:noFill/>
													<a:miter lim="800000"/>
													<a:headEnd/>
													<a:tailEnd/>
												</a:ln>
											</wps:spPr>
											<wps:txbx>
												<w:txbxContent>
													<w:p w14:paraId="79920982" w14:textId="180F8E3D" w:rsidR="005B7084" w:rsidRDefault="005B7084" w:rsidP="005B7084">
														<w:pPr>
															<w:jc w:val="center"/>
															<w:sz w:val="14"/>
														</w:pPr>
														<w:r>
															<w:t>%s</w:t>
														</w:r>
													</w:p>
												</w:txbxContent>
											</wps:txbx>
											<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0">
												<a:noAutofit/>
											</wps:bodyPr>
										</wps:wsp>
									</a:graphicData>
								</a:graphic>
								<wp14:sizeRelH relativeFrom="margin">
									<wp14:pctWidth>0</wp14:pctWidth>
								</wp14:sizeRelH>
								<wp14:sizeRelV relativeFrom="margin">
									<wp14:pctHeight>0</wp14:pctHeight>
								</wp14:sizeRelV>
							</wp:anchor>
						</w:drawing>
					</mc:Choice>
					<mc:Fallback>
						<w:pict>
							<v:shapetype w14:anchorId="3454B481" id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
								<v:stroke joinstyle="miter"/>
								<v:path gradientshapeok="t" o:connecttype="rect"/>
							</v:shapetype>
							<v:shape id="Cuadro de texto 2" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:-355.05pt;margin-top:378.1pt;width:592.65pt;height:20.25pt;rotation:-90;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top" o:gfxdata="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&#10;" stroked="f">
								<v:textbox>
									<w:txbxContent>
										<w:p w14:paraId="79920982" w14:textId="180F8E3D" w:rsidR="005B7084" w:rsidRDefault="005B7084" w:rsidP="005B7084">
											<w:pPr>
												<w:jc w:val="center"/>
												<w:sz w:val="14"/>
											</w:pPr>
											<w:r>
												<w:t>%s</w:t>
											</w:r>
										</w:p>
									</w:txbxContent>
								</v:textbox>
								<w10:wrap type="square" anchorx="margin"/>
							</v:shape>
						</w:pict>
					</mc:Fallback>
				</mc:AlternateContent>
	""" % (self.txt, self.txt)
		print tx
		return tx

	def get_xml2(self):
		value = list()
		t = ''

		value.append('%s<%s%s>' % (self.get_tab(), self.get_name(), t))
		if self.get_choice():
			value.append(self.get_choice().get_xml())
		if self.get_fall_back():
			value.append(self.get_fall_back().get_xml())
		value.append("""<mc:Fallback>
						<w:pict>
							<v:shapetype w14:anchorId="34AD94B0" id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
								<v:stroke joinstyle="miter"/>
								<v:path gradientshapeok="t" o:connecttype="rect"/>
							</v:shapetype>
							<v:shape id="Cuadro de texto 2" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:0;margin-top:14.4pt;width:185.9pt;height:110.6pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:400;mso-height-percent:200;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:center;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:400;mso-height-percent:200;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top" o:gfxdata="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&#10;">
								<v:textbox style="mso-fit-shape-to-text:t">
									<w:txbxContent>
										<w:p w14:paraId="7501BB31" w14:textId="5E774222" w:rsidR="009517B3" w:rsidRDefault="009517B3">
											<w:r>
												<w:t>HOLA</w:t>
											</w:r>
										</w:p>
									</w:txbxContent>
								</v:textbox>
								<w10:wrap type="square"/>
							</v:shape>
						</w:pict>
					</mc:Fallback>""")
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


if __name__ == '__main__':
	class A(object):
		def __init__(self):
			self.tag = 'document'
			self.tab = '\t'
			self.separator = '\n'
			self.indent = 0
			self.idx = 1


	from document import Document
	parent = Document('c:users/jonathan/desktop/pyword/','t.docx')
	parent.EmptyDocument()
	header = parent.get_DefaultHeader()
	x = AlternateContent(header, text='ccccccccc')
	header.add_paragraph(x)
	'''sr = {'vertical': {}}
	sr['vertical']['relative'] = 'margin'
	sr['vertical']['value'] = '0'
	x = AlternateContent(parent, requires='wps', pos=(5, 5), size=(200, 150), size_relative=sr)
	graphic_data = x.get_choice().get_drawing().get_graphic().get_graphic_data()
	shape = Shape(graphic_data)
	shape_pr = shape.get_properties()
	xf = Xform(shape_pr, '', '', ('5', '5'), ('1285875', '523875'))
	shape_pr.elements.append(xf)
	pg = PrstGeom(shape_pr, 'rect', True)
	shape_pr.elements.append(pg)
	sf = SolidFill(shape_pr, scheme_clr='lt1')
	shape_pr.elements.append(sf)
	ln = Line(shape_pr, width='4000')

	sf2 = SolidFill(ln, prst_clr="black")
	ln.elements.append(sf2)
	shape_pr.elements.append(ln)
	txbx = TextBox(shape)
	txbx.add_paragraph('Hola')
	shape.add_element(txbx)
	shape_body = Shape_body(shape)
	shape_body.add_element(PrstTxWarp(shape_body))
	shape.add_element(shape_body)
	graphic_data.set_shape(shape)

	body = parent.get_body()
	section = body.get_active_section()
	p = body.add_paragraph(x)'''

	parent.Save()

