#!/usr/bin/python
# -*- coding: utf-8 -*-


class Picture(object):
	class InlineProperties(object):

		class Graphic(object):
			def __init__(self, parent, rid, description, width, height):
				self.name = 'wp:cNvGraphicFrame'
				self.parent = parent
				self.tab = parent.tab
				self.separator = parent.separator
				self.indent = parent.indent + 1
				self.width = width
				self.height = height
				self.rId = rid
				self.name = "Imagen %d" % rid
				self.description = description

			def get_Description(self):
				return self.description

			def get_RId(self):
				return self.rId

			def SetRId(self, rid):
				self.rId = rid

			def get_name(self):
				return self.name

			def set_name(self, name):
				self.name = name

			def get_width(self):
				return self.width

			def get_height(self):
				return self.height

			def SetWidth(self, width):
				self.width = width

			def set_height(self, height):
				self.height = height

			def get_tab(self, number=0):
				return self.tab * (number + self.indent)

			def get_xml_old(self):
				value = list()
				value.append('%s<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' %
				             (self.get_tab()))
				value.append('%s<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(1)))
				value.append('%s<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(2)))
				value.append('%s<pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:cNvPr id="%d" name="%s"/>' % (self.get_tab(4), self.get_RId(), self.get_name()))
				value.append('%s<pic:cNvPicPr/>' % (self.get_tab(4)))
				value.append('%s</pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:blipFill>' % (self.get_tab(3)))
				# value.append('%s<a:blip r:embed="rId%d">' % (self.get_tab(4), self.get_RId()))
				value.append('%s<a:blip r:embed="rId1">' % (self.get_tab(4)))
				value.append('%s<a:extLst>' % (self.get_tab(5)))
				value.append('%s<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' % (self.get_tab(6)))
				value.append(
					'%s<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>' %
					(self.get_tab(7)))
				value.append('%s</a:ext>' % (self.get_tab(6)))
				value.append('%s</a:extLst>' % (self.get_tab(5)))
				value.append('%s</a:blip>' % (self.get_tab(4)))
				value.append('%s<a:stretch>' % (self.get_tab(4)))
				value.append('%s<a:fillRect/>' % (self.get_tab(5)))
				value.append('%s</a:stretch>' % (self.get_tab(4)))
				value.append('%s</pic:blipFill>' % (self.get_tab(3)))
				value.append('%s<pic:spPr>' % (self.get_tab(3)))
				value.append('%s<a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:off x="0" y="0"/>' % (self.get_tab(5)))
				value.append('%s<a:ext cx="%d" cy="%d"/>' % (self.get_tab(5), self.get_width(), self.get_height()))
				value.append('%s</a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:prstGeom prst="rect">' % (self.get_tab(4)))
				value.append('%s<a:avLst/>' % (self.get_tab(5)))
				value.append('%s</a:prstGeom>' % (self.get_tab(4)))
				value.append('%s</pic:spPr>' % (self.get_tab(3)))
				value.append('%s</pic:pic>' % (self.get_tab(2)))
				value.append('%s</a:graphicData>' % (self.get_tab(1)))
				value.append('%s</a:graphic>' % (self.get_tab(0)))
				return self.separator.join(value)

			def get_xml(self):
				value = list()
				value.append('%s<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' %
				             (self.get_tab()))
				value.append('%s<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(1)))
				value.append('%s<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(2)))
				value.append('%s<pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:cNvPr id="%d" name="%s"/>' % (self.get_tab(4), self.get_RId(), self.get_name()))
				value.append('%s<pic:cNvPicPr/>' % (self.get_tab(4)))
				value.append('%s</pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:blipFill>' % (self.get_tab(3)))
				# value.append('%s<a:blip r:embed="rId%d">' % (self.get_tab(4), self.get_RId()))
				value.append('%s<a:blip r:embed="rId1">' % (self.get_tab(4)))
				value.append('%s<a:extLst>' % (self.get_tab(5)))
				value.append('%s<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' % (self.get_tab(6)))
				value.append(
					'%s<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>' %
					(self.get_tab(7)))
				value.append('%s</a:ext>' % (self.get_tab(6)))
				value.append('%s</a:extLst>' % (self.get_tab(5)))
				value.append('%s</a:blip>' % (self.get_tab(4)))
				value.append('%s<a:stretch>' % (self.get_tab(4)))
				value.append('%s<a:fillRect/>' % (self.get_tab(5)))
				value.append('%s</a:stretch>' % (self.get_tab(4)))
				value.append('%s</pic:blipFill>' % (self.get_tab(3)))
				value.append('%s<pic:spPr>' % (self.get_tab(3)))
				value.append('%s<a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:off x="0" y="0"/>' % (self.get_tab(5)))
				value.append('%s<a:ext cx="%d" cy="%d"/>' % (self.get_tab(5), self.get_width(), self.get_height()))
				value.append('%s</a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:prstGeom prst="rect">' % (self.get_tab(4)))
				value.append('%s<a:avLst/>' % (self.get_tab(5)))
				value.append('%s</a:prstGeom>' % (self.get_tab(4)))
				value.append('%s</pic:spPr>' % (self.get_tab(3)))
				value.append('%s</pic:pic>' % (self.get_tab(2)))
				value.append('%s</a:graphicData>' % (self.get_tab(1)))
				value.append('%s</a:graphic>' % (self.get_tab(0)))
				return self.separator.join(value)

		class GraphicFramePr(object):
			def __init__(self, parent):
				self.name = 'wp:cNvGraphicFramePr'
				self.parent = parent
				self.tab = parent.tab
				self.separator = parent.separator
				self.indent = parent.indent + 1

			def get_xml(self):
				value = list()
				value.append('%s<%s/>' % (self.get_tab(), self.get_name()))

				return self.separator.join(value)

			def get_tab(self, number=0):
				return self.tab * (number + self.indent)

			def get_name(self):
				return self.name

		def __init__(self, parent, rid, path, width, height):
			self.name = 'inline'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1
			self.inline = {'distT': 0, 'distB': 0, 'distL': 0, 'distR': 0, 'simplePos': 0,
			               'wp14:anchorId': "56601072", 'wp14:editId': "0C39F830"}
			width *= 635
			height *= 635
			self.extent = {'cx': width, 'cy': height}
			self.effectExtent = {'l': 0, 't': 0, 'r': 0, 'b': 0}
			self.wrap_square = 'bothSides'
			self.id_picture = rid
			self.name_picture = "Imagen %d" % rid
			self.description_picture = path
			self.graphic_framePr = self.GraphicFramePr(self)
			self.graphic = self.Graphic(self, rid, path, width, height)

		def get_tab(self, number=0):
			return self.tab * (number + self.indent)

		def get_description_picture(self):
			return self.description_picture

		def set_description_picture(self, description_picture):
			self.description_picture = description_picture

		def get_name_picture(self):
			return self.name_picture

		def set_name_picture(self, name_picture):
			self.name_picture = name_picture

		def get_id_picture(self):
			return self.id_picture

		def set_id_picture(self, id_picture):
			self.id_picture = id_picture

		def get_name(self):
			return self.name

		def get_extent(self):
			return self.extent

		def get_effect_extent(self):
			return self.effectExtent

		def get_inline(self):
			return self.inline

		def set_name(self, name):
			self.name = name

		def get_graphic(self):
			return self.graphic

		def get_graphic_frame_pr(self):
			return self.graphic_framePr

		def get_xml(self):
			value = list()
			t = '%s<wp:%s ' % (self.get_tab(), self.get_name())
			t += 'distT="%d" distB="%d" ' % (self.get_inline()['distT'], self.get_inline()['distB'])
			t += 'distL="%d" distR="%d" ' % (self.get_inline()['distL'], self.get_inline()['distR'])
			t += 'wp14:anchorId="%s" wp14:editId="%s">' % (
				self.get_inline()['wp14:anchorId'], self.get_inline()['wp14:editId'])
			value.append(t)

			value.append('%s<wp:extent cx="%s" cy="%s"/>' % (
				self.get_tab(1), self.get_extent()['cx'], self.get_extent()['cy']))
			value.append('%s<wp:effectExtent l="%s" t="%s" r="%s" b="%s"/>' %
			             (self.get_tab(1), self.get_effect_extent()['l'], self.get_effect_extent()['t'],
			              self.get_effect_extent()['r'], self.get_effect_extent()['b']))
			value.append('%s<wp:docPr id="%s" name="%s" descr="%s"/>' %
			             (self.get_tab(1), self.get_id_picture(), self.get_name_picture(),
			              self.get_description_picture()))

			# value.append(self.get_graphicFrame().get_xml())
			value.append(self.get_graphic_frame_pr().get_xml())
			value.append(self.get_graphic().get_xml())

			value.append('%s</wp:%s>' % (self.get_tab(), self.get_name()))

			return self.separator.join(value)

	class Properties(object):

		class Graphic(object):
			def __init__(self, parent, rid, description, width, height):
				self.name = 'wp:cNvGraphicFrame'
				self.parent = parent
				self.tab = parent.tab
				self.separator = parent.separator
				self.indent = parent.indent + 1
				self.width = width
				self.height = height
				self.rId = rid
				self.name = "Imagen %d" % rid
				self.description = description

			def get_Description(self):
				return self.description

			def get_RId(self):
				return self.rId

			def SetRId(self, rid):
				self.rId = rid

			def get_name(self):
				return self.name

			def set_name(self, name):
				self.width = name

			def get_width(self):
				return self.width

			def get_height(self):
				return self.height

			def SetWidth(self, width):
				self.width = width

			def set_height(self, height):
				self.height = height

			def get_tab(self, number=0):
				return self.tab * (number + self.indent)

			def get_xml_old(self):
				value = list()
				value.append('%s<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' %
				             (self.get_tab()))
				value.append('%s<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(1)))
				value.append('%s<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(2)))
				value.append('%s<pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:cNvPr id="%d" name="%s"/>' % (self.get_tab(4), self.get_RId(), self.get_name()))
				value.append('%s<pic:cNvPicPr/>' % (self.get_tab(4)))
				value.append('%s</pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:blipFill>' % (self.get_tab(3)))
				# value.append('%s<a:blip r:embed="rId%d">' % (self.get_tab(4), self.get_RId()))
				value.append('%s<a:blip r:embed="rId1">' % (self.get_tab(4)))
				value.append('%s<a:extLst>' % (self.get_tab(5)))
				value.append('%s<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' % (self.get_tab(6)))
				value.append(
					'%s<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>' %
					(self.get_tab(7)))
				value.append('%s</a:ext>' % (self.get_tab(6)))
				value.append('%s</a:extLst>' % (self.get_tab(5)))
				value.append('%s</a:blip>' % (self.get_tab(4)))
				value.append('%s<a:stretch>' % (self.get_tab(4)))
				value.append('%s<a:fillRect/>' % (self.get_tab(5)))
				value.append('%s</a:stretch>' % (self.get_tab(4)))
				value.append('%s</pic:blipFill>' % (self.get_tab(3)))
				value.append('%s<pic:spPr>' % (self.get_tab(3)))
				value.append('%s<a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:off x="0" y="0"/>' % (self.get_tab(5)))
				value.append('%s<a:ext cx="%d" cy="%d"/>' % (self.get_tab(5), self.get_width(), self.get_height()))
				value.append('%s</a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:prstGeom prst="rect">' % (self.get_tab(4)))
				value.append('%s<a:avLst/>' % (self.get_tab(5)))
				value.append('%s</a:prstGeom>' % (self.get_tab(4)))
				value.append('%s</pic:spPr>' % (self.get_tab(3)))
				value.append('%s</pic:pic>' % (self.get_tab(2)))
				value.append('%s</a:graphicData>' % (self.get_tab(1)))
				value.append('%s</a:graphic>' % (self.get_tab(0)))
				return self.separator.join(value)

			def get_xml(self):
				value = list()
				value.append('%s<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' %
				             (self.get_tab()))
				value.append('%s<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(1)))
				value.append('%s<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' %
				             (self.get_tab(2)))
				value.append('%s<pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:cNvPr id="%d" name="%s"/>' % (self.get_tab(4), self.get_RId(), self.get_name()))
				value.append('%s<pic:cNvPicPr/>' % (self.get_tab(4)))
				value.append('%s</pic:nvPicPr>' % (self.get_tab(3)))
				value.append('%s<pic:blipFill>' % (self.get_tab(3)))
				# value.append('%s<a:blip r:embed="rId%d">' % (self.get_tab(4), self.get_RId()))
				value.append('%s<a:blip r:embed="rId1">' % (self.get_tab(4)))
				value.append('%s<a:extLst>' % (self.get_tab(5)))
				value.append('%s<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' % (self.get_tab(6)))
				value.append(
					'%s<a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>' %
					(self.get_tab(7)))
				value.append('%s</a:ext>' % (self.get_tab(6)))
				value.append('%s</a:extLst>' % (self.get_tab(5)))
				value.append('%s</a:blip>' % (self.get_tab(4)))
				value.append('%s<a:stretch>' % (self.get_tab(4)))
				value.append('%s<a:fillRect/>' % (self.get_tab(5)))
				value.append('%s</a:stretch>' % (self.get_tab(4)))
				value.append('%s</pic:blipFill>' % (self.get_tab(3)))
				value.append('%s<pic:spPr>' % (self.get_tab(3)))
				value.append('%s<a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:off x="0" y="0"/>' % (self.get_tab(5)))
				value.append('%s<a:ext cx="%d" cy="%d"/>' % (self.get_tab(5), self.get_width(), self.get_height()))
				value.append('%s</a:xfrm>' % (self.get_tab(4)))
				value.append('%s<a:prstGeom prst="rect">' % (self.get_tab(4)))
				value.append('%s<a:avLst/>' % (self.get_tab(5)))
				value.append('%s</a:prstGeom>' % (self.get_tab(4)))
				value.append('%s</pic:spPr>' % (self.get_tab(3)))
				value.append('%s</pic:pic>' % (self.get_tab(2)))
				value.append('%s</a:graphicData>' % (self.get_tab(1)))
				value.append('%s</a:graphic>' % (self.get_tab(0)))
				return self.separator.join(value)

		class GraphicFrame(object):
			def __init__(self, parent):
				self.name = 'wp:cNvGraphicFrame'
				self.parent = parent
				self.tab = parent.tab
				self.separator = parent.separator
				self.indent = parent.indent + 1
				self.no_change_aspect = True
				self.no_crop = False
				self.no_rotation = False
				self.no_select = False

			def get_xml(self):
				value = list()
				value.append('%s<%s>' % (self.get_tab(), self.get_name()))
				attributes = list()
				if self.get_ChangeAspect():
					attributes.append('noChangeAspect="1"')
				if self.get_Selection():
					attributes.append('noSelect="1"/>')
				if self.get_Rotation():
					attributes.append('noRot="1"/>')
				if self.get_Crop():
					attributes.append('noCrop="1"/>')
				t = ' '.join(attributes)
				if t:
					t = ' ' + t
				value.append(
					'%s<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"%s/>' %
					(self.get_tab(1), t))
				value.append('%s</%s>' % (self.get_tab(), self.get_name()))
				return self.separator.join(value)

			def get_tab(self, number=0):
				return self.tab * (number + self.indent)

			def get_name(self):
				return self.name

			def EnableSelection(self):
				self.no_select = True

			def DisableSelection(self):
				self.no_select = False

			def get_Selection(self):
				return self.no_rotation

			def EnableRotation(self):
				self.no_rotation = True

			def DisableRotation(self):
				self.no_rotation = False

			def get_Rotation(self):
				return self.no_rotation

			def EnableCrop(self):
				self.no_crop = True

			def DisableCrop(self):
				self.no_crop = False

			def get_Crop(self):
				return self.no_crop

			def EnableChangeAspect(self):
				self.no_change_aspect = True

			def DisableChangeAspect(self):
				self.no_change_aspect = False

			def get_ChangeAspect(self):
				return self.no_change_aspect

		class GraphicFramePr(object):
			def __init__(self, parent):
				self.name = 'wp:cNvGraphicFramePr'
				self.parent = parent
				self.tab = parent.tab
				self.separator = parent.separator
				self.indent = parent.indent + 1

			def get_xml(self):
				value = list()
				value.append('%s<%s/>' % (self.get_tab(), self.get_name()))

				return self.separator.join(value)

			def get_tab(self, number=0):
				return self.tab * (number + self.indent)

			def get_name(self):
				return self.name

		class Position(object):
			def __init__(self, parent, orientation='horizontal', relative_from='paragraph', align=None,
			             position_offset=None):
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

			def get_Align(self):
				return self.align

			def SetAlign(self, align):
				self.align = align

			def get_PositionOffset(self):
				return self.position_offset

			def SetPositionOffset(self, position_offset):
				self.position_offset = position_offset

			def get_Orientation(self):
				return self.orientation

			def SetOrientation(self, orientation):
				self.orientation = orientation
				self.set_name(orientation)

			def get_RelativeFrom(self):
				return self.relative_from

			def SetRelativeFrom(self, relative_from):
				self.relative_from = relative_from

			def get_xml(self):
				value = list()
				value.append('%s<%s relativeFrom="%s">' % (self.get_tab(), self.get_name(), self.get_RelativeFrom()))
				if self.get_Align() is not None:
					value.append('%s<wp:align>%s</wp:align>' % (self.get_tab(1), self.get_Align()))
				if self.get_PositionOffset() is not None:
					value.append('%s<wp:posOffset>%d</wp:posOffset>' % (self.get_tab(1), self.get_PositionOffset()*635))
				value.append('%s</%s>' % (self.get_tab(), self.get_name()))
				return self.separator.join(value)

		class SizeRelative(object):
			def __init__(self, parent, orientation='horizontal', relative_from='margin'):
				self.orientation = orientation
				self.name = ''
				self.set_name(orientation)
				self.relative_from = relative_from
				self.value = 0
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

			def get_Value(self):
				return self.value

			def SetValue(self, value):
				self.value = value

			def get_Orientation(self):
				return self.orientation

			def SetOrientation(self, orientation):
				self.orientation = orientation
				self.set_name(orientation)

			def get_RelativeFrom(self):
				return self.relative_from

			def SetRelativeFrom(self, relative_from):
				self.relative_from = relative_from

			def get_xml(self):
				value = list()
				value.append('%s<%s relativeFrom="%s">' % (self.get_tab(), self.get_name(), self.get_RelativeFrom()))
				if self.get_Orientation() is 'horizontal':
					value.append('%s<wp14:pctWidth>%s</wp14:pctWidth>' % (self.get_tab(1), self.get_Value()))
				else:
					value.append('%s<wp14:pctHeight>%s</wp14:pctHeight>' % (self.get_tab(1), self.get_Value()))
				value.append('%s</%s>' % (self.get_tab(), self.get_name()))
				return self.separator.join(value)

		def __init__(self, parent, rid, path, width, height):
			self.name = 'wp'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1
			self.anchor = {'distT': 0, 'distB': 0, 'distL': 114300, 'distR': 114300, 'simplePos': 0,
			               'relativeHeight': 251658240, 'behindDoc': 0, 'locked': 0, 'layoutInCell': 1,
			               'allowOverlap': 1, 'wp14:anchorId': "68AB4B87", 'wp14:editId': "0BEDF127"}
			width *= 635
			height *= 635
			self.simple_position = {'x': 4595854, 'y': 453224}
			self.position_horizontal = self.Position(self, relative_from='margin', align='left')
			self.position_vertical = self.Position(self, 'vertical', position_offset=3810)
			self.extent = {'cx': width, 'cy': height}
			self.effectExtent = {'l': 0, 't': 0, 'r': 0, 'b': 0}
			self.wrap_square = 'bothSides'
			self.id_picture = rid
			self.name_picture = "Imagen %d" % rid
			self.description_picture = path
			self.graphic_frame = self.GraphicFrame(self)
			self.graphic_framePr = self.GraphicFramePr(self)
			self.graphic = self.Graphic(self, rid, path, width, height)
			self.size_relative_horizontal = self.SizeRelative(self)
			self.size_relative_vertical = self.SizeRelative(self, 'vertical')

		def get_xml(self):
			value = list()
			t = '%s<wp:anchor ' % (self.get_tab())
			t += 'distT="%d" distB="%d" ' % (self.get_Anchor()['distT'], self.get_Anchor()['distB'])
			t += 'distL="%d" distR="%d" ' % (self.get_Anchor()['distL'], self.get_Anchor()['distR'])
			t += 'simplePos="%d" relativeHeight="%d" ' % (
				self.get_Anchor()['simplePos'], self.get_Anchor()['relativeHeight'])
			t += 'behindDoc="%d" locked="%d" ' % (self.get_Anchor()['behindDoc'], self.get_Anchor()['locked'])
			t += 'layoutInCell="%d" allowOverlap="%d" ' % (
				self.get_Anchor()['layoutInCell'], self.get_Anchor()['allowOverlap'])
			t += 'wp14:anchorId="%s" wp14:editId="%s">' % (
			self.get_Anchor()['wp14:anchorId'], self.get_Anchor()['wp14:editId'])
			value.append(t)
			value.append('%s<wp:simplePos x="%d" y="%d"/>' %
			             (self.get_tab(1), self.get_SimplePosition()['x'], self.get_SimplePosition()['y']))
			value.append(self.get_PositionHorizontal().get_xml())
			value.append(self.get_PositionVertical().get_xml())

			value.append(
				'%s<wp:extent cx="%s" cy="%s"/>' % (self.get_tab(1), self.get_Extent()['cx'], self.get_Extent()['cy']))
			value.append('%s<wp:effectExtent l="%s" t="%s" r="%s" b="%s"/>' %
			             (self.get_tab(1), self.get_effect_extent()['l'], self.get_effect_extent()['t'],
			              self.get_effect_extent()['r'], self.get_effect_extent()['b']))

			if self.get_WrapSquare():
				value.append('%s<wp:wrapSquare wrapText="%s"/>' % (self.get_tab(1), self.get_WrapSquare()))
			else:
				value.append("%s<wp:wrapNone/>" % self.get_tab(1))

			value.append('%s<wp:docPr id="%s" name="%s" descr="%s"/>' %
			             (self.get_tab(1), self.get_id_picture(), self.get_name_picture(),
			              self.get_description_picture()))

			# value.append(self.get_graphicFrame().get_xml())
			value.append(self.get_graphic_frame_pr().get_xml())
			value.append(self.get_graphic().get_xml())
			value.append(self.get_SizeRelativeHorizontal().get_xml())
			value.append(self.get_SizeRelativeVertical().get_xml())
			value.append('</wp:anchor>')

			return self.separator.join(value)

		def get_SizeRelativeHorizontal(self):
			return self.size_relative_horizontal

		def get_SizeRelativeVertical(self):
			return self.size_relative_vertical

		def get_graphic(self):
			return self.graphic

		def get_graphicFrame(self):
			return self.graphic_frame

		def get_graphic_frame_pr(self):
			return self.graphic_framePr

		def get_description_picture(self):
			return self.description_picture

		def set_description_picture(self, description_picture):
			self.description_picture = description_picture

		def get_name_picture(self):
			return self.name_picture

		def set_name_picture(self, name_picture):
			self.name_picture = name_picture

		def get_id_picture(self):
			return self.id_picture

		def set_id_picture(self, id_picture):
			self.id_picture = id_picture

		def get_WrapSquare(self):
			return self.wrap_square

		def SetWrapSquare(self, wrap_square):
			self.wrap_square = wrap_square

		def get_PositionHorizontal(self):
			return self.position_horizontal

		def SetPositionHorizontal(self, relative_from='paragraph', align=None, position_offset=None):
			self.position_horizontal = self.Position(self, 'horizontal', relative_from, align, position_offset)

		def get_PositionVertical(self):
			return self.position_vertical

		def SetPositionVertical(self, relative_from='paragraph', align=None, position_offset=None):
			self.position_vertical = self.Position(self, 'vertical', relative_from, align, position_offset)

		def get_effect_extent(self):
			return self.effectExtent

		def SetEffectExtent(self, dc):
			self.effectExtent = dc

		def SetEffectExtentValue(self, key, value):
			if key not in self.extent.keys():
				raise ValueError('La clave %s no en válida' % key)
			self.effectExtent[key] = value

		def get_Extent(self):
			return self.extent

		def SetExtent(self, dc):
			self.extent = dc

		def SetExtentValue(self, key, value):
			if key not in self.extent.keys():
				raise ValueError('La clave %s no en válida' % key)
			self.extent[key] = value

		def get_SimplePosition(self):
			return self.simple_position

		def SetSimplePosition(self, dc):
			self.simple_position = dc

		def SetSimplePositionValue(self, key, value):
			if key not in self.simple_position.keys():
				raise ValueError('La clave %s no en válida' % key)
			self.simple_position[key] = value

		def get_Anchor(self):
			return self.anchor

		def SetAnchor(self, dc):
			self.anchor = dc

		def SetAnchorValue(self, key, value):
			if key not in self.anchor.keys():
				raise ValueError('La clave %s no en válida' % key)
			self.anchor[key] = value

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_separator(self):
			return self.separator

		def get_name(self):
			return self.name

	def __init__(self, parent, rid, name, width, height, anchor='inline'):
		self.name = 'drawing'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1

		if anchor == 'inline':
			self.properties = self.InlineProperties(self, rid, name, width, height)
		else:
			self.properties = self.Properties(self, rid, name, width, height)

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_separator(self):
		return self.separator

	def get_properties(self):
		return self.properties

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()

		value.append('%s<w:%s>' % (self.get_tab(), self.get_name()))
		value.append(self.get_properties().get_xml())
		value.append('%s</w:%s>' % (self.get_tab(), self.get_name()))

		return self.get_separator().join(value)
