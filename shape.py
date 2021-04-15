#!/usr/bin/python
# -*- coding: utf-8 -*-
from parts.word.elements import paragraph


class TextBox(object):
	def __init__(self, parent, text, position, size, rotation=0, r_position=(), simple_position=(0, 0),
					background_color='FFFFFF', flip_vertical='', flip_horizontal='', horizontal_alignment='j',
					font_format='', font_size=None):
		self.name = 'shape'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1

		self.content = AlternateContent(parent)
		choice = Choice()
		self.content.set_choice(choice)
		drawing = Drawing()
		choice.set_drawing(drawing)
		anchor = Anchor()
		drawing.set_element(anchor)
		simple_pos = SimplePosition(*simple_position)
		simple_pos.set_parent(anchor)
		anchor.add_element(simple_pos)

		for rp in r_position:
			position_r = Position(rp['position'], orientation=rp['orientation'], relative_from=rp['relative'])
			position_r.set_parent(anchor)
			anchor.add_element(position_r)

		extent = Extent(size)
		extent.set_parent(anchor)
		anchor.add_element(extent)
		effect_exent = EffectExent()
		effect_exent.set_parent(anchor)
		anchor.add_element(effect_exent)
		wrap_square = WrapSquare()
		wrap_square.set_parent(anchor)
		anchor.add_element(wrap_square)
		doc_pr = DocPr(217)
		doc_pr.set_parent(anchor)
		anchor.add_element(doc_pr)
		graphic_frame = CNvGraphicFramePr()
		graphic_frame.set_parent(anchor)
		anchor.add_element(graphic_frame)
		graphic = Graphic()
		graphic.set_parent(anchor)
		anchor.add_element(graphic)
		shape = Shape()
		shape.set_parent(graphic)
		graphic.set_shape(shape)

		prv = PositionRelative(orietation='horizontal')
		prv.set_parent(anchor)
		anchor.add_element(prv)
		prh = PositionRelative(orietation='vertical')
		prh.set_parent(anchor)
		anchor.add_element(prh)

		sp_pr = CNvSpPr()
		shape.add_element(sp_pr)
		shape_pr = ShapeProperties()
		shape.add_element(shape_pr)
		shape_pr.add_element(Xform(position, rotation=rotation, flip_vertical=flip_vertical, flip_horizontal=flip_horizontal))
		shape_pr.add_element(PrstGeom())
		shape_pr.add_element(SolidFill(background_color))
		line = Line('9525')
		line.add_properties_list('noFill')
		line.add_properties_list(('miter', 'lim', '800000'))
		line.add_properties_list('headEnd')
		line.add_properties_list('tailEnd')
		shape_pr.add_element(line)

		txbx = Txbx()
		shape.add_element(txbx)
		txbx.add_paragraph(text, horizontal_alignment, font_format, font_size)

		shape_body = Shape_body(auto_fit=False)
		shape.add_element(shape_body)

		fall_back = FallBack()
		self.content.set_fall_back(fall_back)
		pict = Pict()
		fall_back.set_pict(pict)
		shape_type = ShapeType()
		pict.add_element(shape_type)
		shape_type.add_element({'v:stroke': {'joinstyle': "miter"}})
		shape_type.add_element({'v:path': {'gradientshapeok': "t", 'o:connecttype': "rect"}})
		fall_shape = FallShape(wrap='square')
		pict.add_element(fall_shape)
		text_box = Txbx()
		text_box.name = 'v:textbox'
		fall_shape.set_object(text_box)
		text_box.add_paragraph(text)

	def get_separator(self):
		return self.separator

	def get_parent(self):
		return self.parent

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_xml(self):
		return self.content.get_xml()


class AlternateContent(object):
	def __init__(self, parent):
		self.name = 'mc:AlternateContent'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.choice = None
		self.fallBack = None
		'''self.txt = text.decode('iso-8859-1').encode('utf8').replace('%EURO%', '€').replace('&', '&amp;').replace(
					'<', '&lt;').replace('>', '&gt;')'''

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1
		print _parent.indent

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_parent(self):
		return self.parent

	def get_choice(self):
		return self.choice

	def set_choice(self, choice):
		self.choice = choice
		self.choice.set_parent(self)

	def get_name(self):
		return self.name

	def get_fall_back(self):
		return self.fallBack

	def set_fall_back(self, fall_back):
		self.fallBack = fall_back
		self.fallBack.set_parent(self)

	def get_xml(self):
		value = list()
		t = ''

		value.append('%s<%s%s>' % (self.get_tab(), self.get_name(), t))
		if self.get_choice():
			value.append(self.get_choice().get_xml())
		if self.get_fall_back():
			value.append(self.get_fall_back().get_xml())

		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Choice(object):
	def __init__(self, requires='wps'):
		self.name = 'mc:Choice'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.requires = requires
		self.drawing = None

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

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
		self.drawing.set_parent(self)

	def get_xml(self):
		value = list()
		t = ''
		if self.get_requires():
			t = ' Requires="%s"' % self.get_requires()

		value.append('%s<%s%s>' % (self.get_tab(), self.get_name(), t))
		value.append(self.get_drawing().get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class FallBack(object):
	def __init__(self):
		self.name = 'mc:Fallback'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.pict = None

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_name(self):
		return self.name

	def get_pict(self):
		return self.pict

	def set_pict(self, pict):
		pict.set_parent(self)
		self.pict = pict

	def get_xml(self):
		value = list()

		value.append('%s<%s>' % (self.get_tab(), self.get_name()))
		value.append(self.get_pict().get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Drawing(object):

	def __init__(self):
		self.name = 'w:drawing'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.element = None

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_element(self):
		return self.element

	def set_element(self, element):
		self.element = element
		self.element.set_parent(self)

	def get_xml(self):
		value = list()
		value.append('%s<%s>' % (self.get_tab(), self.get_name()))
		value.append(self.get_element().get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Anchor(object):

	def __init__(self, allow_overlap="1", dist_b="0", dist_l="114300", dist_r="114300", dist_t="0", layout_in_cell="1",
					behind_doc="0", anchor_id="45C456EB", locked="0", wp14_edit_id="0CFAEE44", simple_pos="0",
					relative_height="251659264"):
		self.name = 'wp:anchor'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.allow_overlap = allow_overlap
		self.dist_b = dist_b
		self.dist_l = dist_l
		self.dist_r = dist_r
		self.dist_t = dist_t
		self.layout_in_cell = layout_in_cell
		self.behind_doc = behind_doc
		self.anchor_id = anchor_id
		self.locked = locked
		self.wp14_edit_id = wp14_edit_id
		self.simple_pos = simple_pos
		self.relative_height = relative_height

		self.elements = list()

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def add_element(self, element):
		self.elements.append(element)

	def get_elements(self):
		return self.elements

	def set_elements(self, elements):
		self.elements = elements

	def get_allow_overlap(self):
		return self.allow_overlap

	def set_allow_overlap(self, allow_overlap):
		self.allow_overlap = allow_overlap

	def get_dist_b(self):
		return self.dist_b

	def set_dist_b(self, dist_b):
		self.dist_b = dist_b

	def get_dist_l(self):
		return self.dist_l

	def set_dist_l(self, dist_l):
		self.dist_l = dist_l

	def get_dist_r(self):
		return self.dist_r

	def set_dist_r(self, dist_r):
		self.dist_r = dist_r

	def get_dist_t(self):
		return self.dist_t

	def set_dist_t(self, dist_t):
		self.dist_t = dist_t

	def get_layout_in_cell(self):
		return self.layout_in_cell

	def set_layout_in_cell(self, layout_in_cell):
		self.layout_in_cell = layout_in_cell

	def get_behind_doc(self):
		return self.behind_doc

	def set_behind_doc(self, behind_doc):
		self.behind_doc = behind_doc

	def get_anchor_id(self):
		return self.anchor_id

	def set_anchor_id(self, anchor_id):
		self.anchor_id = anchor_id

	def get_locked(self):
		return self.locked

	def set_locked(self, locked):
		self.locked = locked

	def get_wp14_edit_id(self):
		return self.wp14_edit_id

	def set_wp14_edit_id(self, wp14_edit_id):
		self.wp14_edit_id = wp14_edit_id

	def get_simple_pos(self):
		return self.simple_pos

	def set_simple_pos(self, simple_pos):
		self.simple_pos = simple_pos

	def get_relative_height(self):
		return self.relative_height

	def set_relative_height(self, relative_height):
		self.relative_height = relative_height

	def get_xml(self):
		args = list()

		if self.get_dist_t():
			args.append('distT="' + self.get_dist_t() + '"')

		if self.get_dist_b():
			args.append('distB="' + self.get_dist_b() + '"')

		if self.get_dist_l():
			args.append('distL="' + self.get_dist_l() + '"')

		if self.get_dist_r():
			args.append('distR="' + self.get_dist_r() + '"')

		if self.get_simple_pos():
			args.append('simplePos="' + self.get_simple_pos() + '"')

		if self.get_relative_height():
			args.append('relativeHeight="' + self.get_relative_height() + '"')

		if self.get_behind_doc():
			args.append('behindDoc="' + self.get_behind_doc() + '"')

		if self.get_locked():
			args.append('locked="' + self.get_locked() + '"')

		if self.get_layout_in_cell():
			args.append('layoutInCell="' + self.get_layout_in_cell() + '"')

		if self.get_allow_overlap():
			args.append('allowOverlap="' + self.get_allow_overlap() + '"')

		if self.get_anchor_id():
			args.append('anchorId="' + self.get_anchor_id() + '"')

		if self.get_wp14_edit_id():
			args.append('wp14_editId="' + self.get_wp14_edit_id() + '"')

		args = ' '.join(args)
		if args:
			args = ' ' + args
		value = list()
		value.append('%s<%s%s>' % (self.get_tab(), self.get_name(), args))
		for element in self.elements:
			value.append(element.get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class SimplePosition(object):

	def __init__(self, x, y):
		self.name = 'wp:simplePos'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.y = y
		self.x = x

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_x(self):
		return self.x

	def set_x(self, x):
		self.x = x

	def get_y(self):
		return self.y

	def set_y(self, y):
		self.y = y

	def set_position(self, x, y):
		self.y = y
		self.x = x

	def get_xml(self):
		value = list()
		value.append('%s<%s x="%s" y="%s"/>' % (self.get_tab(), self.get_name(), str(self.get_x()), str(self.get_y())))

		return self.separator.join(value)


class Position(object):

	def __init__(self, position, orientation='vertical', relative_from='margin'):
		if orientation == 'vertical':
			self.name = 'wp:positionV'
		else:
			self.name = 'wp:positionH'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.relative_from = relative_from
		self.position = position * 635

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_relative_from(self):
		return self.relative_from

	def set_relative_from(self, relative_from):
		self.relative_from = relative_from

	def get_position(self):
		return self.position

	def set_position(self, position):
		self.position = position * 635

	def get_xml(self):
		value = list()
		value.append('%s<%s relativeFrom="%s">' % (self.get_tab(), self.get_name(), self.get_relative_from()))
		value.append('%s<wp:posOffset>%s</wp:posOffset>' % (self.get_tab(1), self.get_position()))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class PositionRelative(object):

	def __init__(self, position=0, orietation='vertical', relative_from='margin'):
		if orietation == 'vertical':
			self.name = 'wp14:sizeRelV'
			self.tag = 'wp14:pctHeight'
		else:
			self.name = 'wp14:sizeRelH'
			self.tag = 'wp14:pctWidth'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.relative_from = relative_from
		self.position = position * 635

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_tag(self):
		return self.tag

	def get_relative_from(self):
		return self.relative_from

	def set_relative_from(self, relative_from):
		self.relative_from = relative_from

	def get_position(self):
		return self.position

	def set_position(self, position):
		self.position = position * 635

	def get_xml(self):
		value = list()
		value.append('%s<%s relativeFrom="%s">' % (self.get_tab(), self.get_name(), self.get_relative_from()))
		value.append('%s<%s>%s</%s>' % (self.get_tab(1), self.get_tag(), self.get_position(), self.get_tag()))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Extent(object):

	def __init__(self, size):
		self.name = 'wp:extent'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.size = size

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_size(self):
		return self.size

	def set_size(self, size):
		self.size = size

	def get_xml(self):
		value = list()
		value.append('%s<%s cx="%s" cy="%s"/>' % (self.get_tab(), self.get_name(), str(self.get_size()[0] * 635),
													str(self.get_size()[1] * 635)))

		return self.separator.join(value)


class EffectExent(object):

	def __init__(self, left=0, right=0, top=0, bottom=0):
		self.name = 'wp:effectExtent'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.left = left
		self.right = right
		self.top = top
		self.bottom = bottom

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_left(self):
		return self.left

	def set_left(self, left):
		self.left = left

	def get_right(self):
		return self.right

	def set_right(self, right):
		self.right = right

	def get_top(self):
		return self.top

	def set_top(self, top):
		self.top = top

	def get_bottom(self):
		return self.bottom

	def set_bottom(self, bottom):
		self.bottom = bottom

	def get_xml(self):
		value = list()
		value.append('%s<%s l="%s" t="%s" r="%s" b="%s"/>' % (self.get_tab(),
																self.get_name(),
																self.get_left(),
																self.get_top(),
																self.get_right(),
																self.get_bottom()))

		return self.separator.join(value)


class WrapSquare(object):

	def __init__(self, wrap_text="bothSides"):
		self.name = 'wp:wrapSquare'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.wrap_text = wrap_text

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_wrap_text(self):
		return self.wrap_text

	def set_wrap_text(self, wrap_text):
		self.wrap_text = wrap_text

	def get_xml(self):
		value = list()
		value.append('%s<%s wrapText="%s"/>' % (self.get_tab(), self.get_name(), self.get_wrap_text()))

		return self.separator.join(value)


class DocPr(object):

	def __init__(self, shape_id, shape_name='Text box'):
		self.name = 'wp:docPr'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.shape_name = shape_name
		self.shape_id = shape_id

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_shape_id(self):
		return self.shape_id

	def set_shape_id(self, shape_id):
		self.shape_id = shape_id

	def get_shape_name(self):
		return self.shape_name

	def set_shape_name(self, shape_name):
		self.shape_name = shape_name

	def get_xml(self):
		value = list()
		value.append('%s<%s id="%s" name="%s"/>' % (self.get_tab(),
													self.get_name(),
													str(self.get_shape_id()),
													self.get_shape_name()))

		return self.separator.join(value)


class CNvGraphicFramePr(object):

	def __init__(self):
		self.name = 'wp:cNvGraphicFramePr'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.graphic_frame_locks = "http://schemas.openxmlformats.org/drawingml/2006/main"

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_graphic_frame_locks(self):
		return self.graphic_frame_locks

	def set_graphic_frame_locks(self, graphic_frame_locks):
		self.graphic_frame_locks = graphic_frame_locks

	def get_xml(self):
		value = list()
		value.append('%s<%s>' % (self.get_tab(), self.get_name()))
		value.append('%s<a:graphicFrameLocks xmlns:a="%s"/>' % (self.get_tab(1), self.get_graphic_frame_locks()))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Graphic(object):
	def __init__(self):
		self.name = 'a:graphic'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.graphicData = GraphicData(self)
		self.xmls = 'http://schemas.openxmlformats.org/drawingml/2006/main'

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1
		self.graphicData.set_parent(self)

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

	def set_shape(self, shape):
		self.get_graphic_data().set_shape(shape)
		self.get_graphic_data().get_shape().set_parent(self)

	def get_xml(self):
		value = list()
		value.append('%s<%s xmlns:a="%s">' % (self.get_tab(), self.get_name(), self.get_xmls()))
		value.append(self.get_graphic_data().get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))
		return self.separator.join(value)


class GraphicData(object):
	def __init__(self, parent):
		self.name = 'a:graphicData'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.shape = None
		self.uri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

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


class Shape(object):

	def __init__(self):
		self.name = 'wps:wsp'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.elements = list()

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_elements(self):
		return self.elements

	def set_elements(self, elements):
		self.elements = elements

	def add_element(self, element):
		element.set_parent(self)
		self.elements.append(element)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		value.append('%s<%s>' % (self.get_tab(), self.get_name()))
		for element in self.get_elements():
			value.append(element.get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class CNvSpPr(object):
	def __init__(self, tx_box="1", no_change_arrowheads="1"):
		self.name = 'wps:cNvSpPr'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.tx_box = tx_box
		self.no_change_arrowheads = no_change_arrowheads

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_no_change_arrowheads(self):
		return self.no_change_arrowheads

	def set_no_change_arrowheads(self, no_change_arrowheads):
		self.no_change_arrowheads = no_change_arrowheads

	def get_tx_box(self):
		return self.tx_box

	def set_tx_box(self, tx_box):
		self.tx_box = tx_box

	def get_xml(self):
		value = list()
		value.append('%s<%s txBox="%s">' % (self.get_tab(), self.get_name(), self.get_tx_box()))
		value.append('%s<a:spLocks noChangeArrowheads="%s"/>' % (self.get_tab(1), self.get_no_change_arrowheads()))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class ShapeProperties(object):
	def __init__(self, mode='auto'):
		self.name = 'wps:spPr'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.mode = mode
		self.elements = list()

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_elements(self):
		return self.elements

	def set_elements(self, elements):
		self.elements = elements

	def add_element(self, element):
		element.set_parent(self)
		self.elements.append(element)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_mode(self):
		return self.mode

	def set_mode(self, mode):
		self.mode = mode

	def get_xml(self):
		value = list()
		value.append('%s<%s bwMode="%s">' % (self.get_tab(), self.get_name(), self.get_mode()))
		for element in self.elements:
			value.append(element.get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Xform(object):

	def __init__(self, ext, flip_horizontal='', flip_vertical='', rotation=0, offset=(0, 0)):
		self.name = 'a:xfrm'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.elements = list()
		self.flip_horizontal = flip_horizontal
		self.flip_vertical = flip_vertical
		self.offset = offset
		self.ext = (ext[0] * 635, ext[1] * 635)
		self.rotation = rotation * 60000

	def get_parent(self):
		return self.parent

	def get_xml(self):
		value = list()
		tx = '%s<%s' % (self.get_tab(), self.get_name())
		if self.get_flip_horizontal():
			tx += ' flipH="%s"' % self.get_flip_horizontal()
		if self.get_flip_vertical():
			tx += ' flipV="%s"' % self.get_flip_vertical()
		if self.get_rotation():
			tx += ' rot="%s"' % self.get_rotation()
		tx += '>'
		value.append(tx)

		value.append('%s<a:off x="%s" y="%s"/>' % (self.get_tab(1), self.get_offset(0), self.get_offset(1)))
		value.append('%s<a:ext cx="%s" cy="%s"/>' % (self.get_tab(1), self.get_ext(0), self.get_ext(1)))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_ext(self, pos):
		return self.ext[pos]

	def set_ext(self, x, y):
		self.ext = (x, y)

	def get_rotation(self):
		return self.rotation

	def set_rotation(self, rotation):
		self.rotation = rotation

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

	def __init__(self, prst='rect', avlst=True):
		self.name = 'a:prstGeom'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.elements = list()
		self.avlst = avlst
		self.prst = prst

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

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
	def __init__(self, width=''):
		self.name = 'a:ln'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.properties_list = list()
		self.width = width

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_properties_list(self):
		return self.properties_list

	def set_properties_list(self, properties_list):
		self.properties_list = properties_list

	def add_properties_list(self, prop):
		self.properties_list.append(prop)

	def get_name(self):
		return self.name

	def get_width(self):
		return self.width

	def set_width(self, width):
		self.width = width

	def get_xml(self):
		value = list()
		for element in self.properties_list:
			if type(element) == str:
				value.append('%s<a:%s/>' % (self.get_tab(1), element))
			elif type(element) in [tuple, list]:
				value.append('%s<a:%s %s="%s"/>' % (self.get_tab(1), element[0], element[1], element[2]))

		if value:
			value.insert(0, '%s<%s w="%s">' % (self.get_tab(), self.get_name(), str(self.get_width())))
			value.append('%s</%s>' % (self.get_tab(), self.get_name()))
		else:
			value.append('%s<%s w="%s"/>' % (self.get_tab(), self.get_name(), str(self.get_width())))

		return self.separator.join(value)


class SolidFill(object):

	def __init__(self, color='FFFFFF'):
		self.name = 'a:solidFill'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.color = color

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_color(self):
		return self.color

	def set_color(self, value):
		self.color = value

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		if self.get_color():
			value.append('%s<a:srgbClr val="%s"/>' % (self.get_tab(1), self.get_color()))

		if value:
			value.insert(0, '%s<%s>' % (self.get_tab(), self.get_name()))
			value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class Txbx(object):
	def __init__(self):
		self.name = 'wps:txbx'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.elements = list()

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

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
		p = paragraph.Paragraph(self, _x.idx, text, horizontal_alignment, font_format, font_size, nulo=is_null)
		p.get_properties().set_pstyle('')
		self.elements.append(p)
		return p


class Shape_body(object):
	def __init__(self, auto_fit=True):
		self.name = 'wps:bodyPr'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.attributes = {
			"rot": "0",
			"vert": "horz",
			"wrap": "square",
			"lIns": "91440",
			"tIns": "45720",
			"rIns": "91440",
			"bIns": "45720",
			"anchor": "t",
			"anchorCtr": "0",
		}
		self.auto_fit = auto_fit
		self.elements = list()

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

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
		ls = ["rot", "vert", "wrap", "lIns", "tIns", "rIns", "bIns", "anchor", "anchorCtr"]
		for key in ls:
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


class Pict(object):

	def __init__(self):
		self.name = 'w:pict'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.elements = list()

	def get_parent(self):
		return self.parent

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_elements(self):
		return self.elements

	def set_elements(self, elements):
		self.elements = elements

	def add_element(self, element):
		element.set_parent(self)
		self.elements.append(element)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		value.append('%s<%s>' % (self.get_tab(), self.get_name()))
		for element in self.get_elements():
			value.append(element.get_xml())
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class ShapeType(object):

	def __init__(self, w14_anchor_id="3454B481", _id="_x0000_t202", coordsize="21600,21600", o_spt="202",
					path="m,l,21600r21600,l21600,xe"):
		self.name = 'v:shapetype'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0

		self.w14_anchor_id = w14_anchor_id
		self.id = _id
		self.coordsize = coordsize
		self.o_spt = o_spt
		self.path = path

		self.elements = list()

	def get_parent(self):
		return self.parent

	def get_w14_anchor_id(self):
		return self.w14_anchor_id

	def set_w14_anchor_id(self, w14_anchor_id):
		self.w14_anchor_id = w14_anchor_id

	def get_id(self):
		return self.id

	def set_id(self, _id):
		self.id = _id

	def get_coordsize(self):
		return self.coordsize

	def set_coordsize(self, coordsize):
		self.coordsize = coordsize

	def get_o_spt(self):
		return self.o_spt

	def set_o_spt(self, o_spt):
		self.o_spt = o_spt

	def get_path(self):
		return self.path

	def set_path(self, path):
		self.path = path

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_elements(self):
		return self.elements

	def set_elements(self, elements):
		self.elements = elements

	def add_element(self, element):
		self.elements.append(element)

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		args = list()
		if self.get_w14_anchor_id():
			args.append('w14:anchorId="' + self.get_w14_anchor_id() + '"')

		if self.get_id():
			args.append('id="' + self.get_id() + '"')

		if self.get_coordsize():
			args.append('coordsize="' + self.get_coordsize() + '"')

		if self.get_o_spt():
			args.append('o:spt="' + self.get_o_spt() + '"')

		if self.get_path():
			args.append('path="' + self.get_path() + '"')

		value = list()
		value.append('%s<%s %s>' % (self.get_tab(), self.get_name(), ' '.join(args)))
		for element in self.get_elements():
			for key in element.keys():
				vals = list()
				for key_ in element[key].keys():
					vals.append('%s="%s"' % (key_, element[key][key_]))
				value.append('%s<%s %s/>' % (self.get_tab(1), key, ' '.join(vals)))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)


class FallShape(object):

	def __init__(self, _id="Shape", o_spid="_x0000_s1026", _type="#_x0000_t202", stroked='f', wrap=''):
		self.name = 'v:shape'
		self.parent = None
		self.tab = ''
		self.separator = ''
		self.indent = 0
		self.id = _id
		self.o_spid = o_spid
		self.type = _type
		self.stroked = stroked
		self.wrap = wrap
		self.style = "position:absolute;margin-left:-355.05pt;margin-top:378.1pt;width:592.65pt;height:20.25pt;"
		self.style += "rotation:-90;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;"
		self.style += "mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;"
		self.style += "mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:absolute;"
		self.style += "mso-position-horizontal-relative:text;mso-position-vertical:absolute;"
		self.style += "mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;"
		self.style += "mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top"
		self.o_gfxdata = "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF&#10"
		self.o_gfxdata += ";90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA&#10"
		self.o_gfxdata += ";0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD&#10"
		self.o_gfxdata += ";OlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893&#10"
		self.o_gfxdata += ";SUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y&#10;JsXecLq0q"
		self.o_gfxdata += "+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl&#10;bHOkkMFqwzAMhu"
		self.o_gfxdata += "+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR&#10;JVI2sOt6UJgd"
		self.o_gfxdata += "+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY&#10;22YiTra2kYMu1l1tQD30"
		self.o_gfxdata += "/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i&#10;OWA14Fm+Q8a1a8+Bvu/d"
		self.o_gfxdata += "/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA&#10"
		self.o_gfxdata += ";IQAWQs4yLQIAADMEAAAOAAAAZHJzL2Uyb0RvYy54bWysU9uO2yAQfa/Uf0C8N46tONm14qy22aaq&#10"
		self.o_gfxdata += ";tL1I234ABhyjAuMCiZ1+fQccJWn7VpUHxDDD4cyZmfXDaDQ5SucV2Jrmszkl0nIQyu5r+u3r7s0d&#10"
		self.o_gfxdata += ";JT4wK5gGK2t6kp4+bF6/Wg99JQvoQAvpCIJYXw19TbsQ+irLPO+kYX4GvbTobMEZFtB0+0w4NiC6&#10;0Vkxny"
		self.o_gfxdata += "+zAZzoHXDpPd4+TU66SfhtK3n43LZeBqJritxC2l3am7hnmzWr9o71neJnGuwfWBimLH56&#10"
		self.o_gfxdata += ";gXpigZGDU39BGcUdeGjDjIPJoG0VlykHzCaf/5HNS8d6mXJBcXx/kcn/P1j+6fjFESVqWuQrSiwz&#10"
		self.o_gfxdata += ";WKTtgQkHREgS5BiAFFGmofcVRr/0GB/GtzBiuVPKvn8G/t0TC9uO2b18dA6GTjKBNPP4Mrt5OuH4&#10;CNIMH0Hgb"
		self.o_gfxdata += "+wQIAGNrTPEAdYoX2JtcaVrFIngZ1i906ViSItwvFyVxXJZlpRw9BXlKl+V6UdWRbBY&#10"
		self.o_gfxdata += ";kN758F6CIfFQU4cdkVDZ8dmHSO4aEsM9aCV2SutkuH2z1Y4cGXbPLq0z+m9h2pKhpvdlUSZkC/F9&#10"
		self.o_gfxdata += ";aiyjAna3Vqamd1NC6TqK886KdA5M6emMTLQ9qxUFmqQKYzNiYJSwAXFC3ZJCqAZOHSbUgftJyYAd&#10;XFP"
		self.o_gfxdata += "/48CcpER/sKj9fb5YxJZPxqJcFWi4W09z62GWI1RNAyXTcRvSmEQdLDxijVqV9LoyOXPFzkwy&#10"
		self.o_gfxdata += ";nqcotv6tnaKus775BQAA//8DAFBLAwQUAAYACAAAACEAqtzyR+UAAAANAQAADwAAAGRycy9kb3du&#10"
		self.o_gfxdata += ";cmV2LnhtbEyPT0vDQBDF74LfYRnBi6SbNJI0MZsixT/0ItgWobdtMibB7GzIbtvop+940uO89+PN&#10"
		self.o_gfxdata += ";e8VyMr044eg6SwqiWQgCqbJ1R42C3fY5WIBwXlOte0uo4BsdLMvrq0LntT3TO542vhEcQi7XClrv&#10"
		self.o_gfxdata += ";h1xKV7VotJvZAYm9Tzsa7fkcG1mP+szhppfzMEyk0R3xh1YPuGqx+tocjYL09S3Z+5X56fYv4Tp7&#10"
		self.o_gfxdata += ";ujPr4f5Dqdub6fEBhMfJ/8HwW5+rQ8mdDvZItRO9giCK05RZdhYxj2AkyNI5iAMrcZJFIMtC/l9R&#10;XgAAAP"
		self.o_gfxdata += "//AwBQSwECLQAUAAYACAAAACEAtoM4kv4AAADhAQAAEwAAAAAAAAAAAAAAAAAAAAAAW0Nv&#10"
		self.o_gfxdata += ";bnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAAAAAAAAAAAAA&#10"
		self.o_gfxdata += ";AC8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQAWQs4yLQIAADMEAAAOAAAAAAAAAAAAAAAA&#10"
		self.o_gfxdata += ";AC4CAABkcnMvZTJvRG9jLnhtbFBLAQItABQABgAIAAAAIQCq3PJH5QAAAA0BAAAPAAAAAAAAAAAA&#10"
		self.o_gfxdata += ";AAAAAIcEAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAQABADzAAAAmQUAAAAA&#10;"
		self.object = None

	def get_parent(self):
		return self.parent

	def get_id(self):
		return self.id

	def set_id(self, _id):
		self.id = _id

	def get_o_spid(self):
		return self.o_spid

	def set_o_spid(self, o_spid):
		self.o_spid = o_spid

	def get_type(self):
		return self.type

	def set_type(self, _type):
		self.type = _type

	def get_wrap(self):
		return self.wrap

	def set_wrap(self, wrap):
		self.wrap = wrap

	def get_object(self):
		return self.object

	def set_object(self, _object):
		_object.set_parent(self)
		self.object = _object

	def get_style(self):
		return self.style

	def set_style(self, style):
		self.style = style

	def get_o_gfxdata(self):
		return self.o_gfxdata

	def set_o_gfxdata(self, o_gfxdata):
		self.o_gfxdata = o_gfxdata

	def get_stroked(self):
		return self.stroked

	def set_stroked(self, stroked):
		self.stroked = stroked

	def set_parent(self, _parent):
		self.parent = _parent
		self.parent = _parent
		self.tab = _parent.tab
		self.separator = _parent.separator
		self.indent = _parent.indent + 1

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_name(self):
		return self.name

	def get_xml(self):
		args = list()

		if self.get_id():
			args.append('id="'+self.get_id()+'"')

		if self.get_o_spid():
			args.append('o:spid="'+self.get_o_spid()+'"')

		if self.get_type():
			args.append('type="'+self.get_type()+'"')

		if self.get_style():
			args.append('style="'+self.get_style()+'"')

		if self.get_o_gfxdata():
			args.append('o:gfxdata="'+self.get_o_gfxdata()+'"')

		if self.get_stroked():
			args.append('stroked="'+self.get_type()+'"')

		value = list()
		value.append('%s<%s %s>' % (self.get_tab(), self.get_name(), ' '.join(args)))
		value.append(self.get_object().get_xml())
		if self.get_wrap():
			value.append('%s<w14:wrap type="%s"/>' % (self.get_tab(1), self.get_wrap()))
		value.append('%s</%s>' % (self.get_tab(), self.get_name()))

		return self.separator.join(value)
