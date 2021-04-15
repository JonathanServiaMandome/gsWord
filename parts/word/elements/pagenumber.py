#!/usr/bin/python
# -*- coding: utf-8 -*-
from parts.word.elements import paragraph


class Properties(object):
	class PartObject(object):
		def __init__(self, parent, doc_part_gallery='Page Numbers (Top of Page)', doc_part_unique=True):
			self.name = 'docPartObj'
			self.parent = parent
			self.tab = parent.tab
			self.separator = parent.separator
			self.indent = parent.indent + 1
			self.docPartGallery = doc_part_gallery
			self.docPartUnique = doc_part_unique

		def get_parent(self):
			return self.parent

		def set_doc_part_gallery(self, value):
			self.docPartGallery = value

		def set_doc_part_unique(self, value):
			self.docPartUnique = value

		def get_doc_part_gallery(self):
			return self.docPartGallery

		def get_separator(self):
			return self.separator

		def get_doc_part_unique(self):
			return self.docPartUnique

		def get_name(self):
			return self.name

		def get_tab(self, number=0):
			return self.tab * (self.indent + number)

		def get_xml(self):
			value = list()
			value.append('%s<w:%s>' % (self.get_tab(), self.get_name()))
			value.append('%s<w:docPartGallery w:val="%s"/>' % (self.get_tab(1), self.get_doc_part_gallery()))
			if self.get_doc_part_unique():
				value.append('%s<w:docPartUnique/>' % self.get_tab(1))
			value.append('%s</w:%s>' % (self.get_tab(), self.get_name()))
			return self.get_separator().join(value)

	def __init__(self, parent, _id='98381352'):
		self.name = 'sdtPr'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.id = _id
		self.part_object = self.PartObject(self)

	def get_parent(self):
		return self.parent

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_separator(self):
		return self.separator

	def get_id(self):
		return self.id

	def get_part_object(self):
		return self.part_object

	def set_id(self, _id):
		self.id = _id

	def set_part_object(self, part_object):
		self.part_object = part_object

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()
		value.append('%s<w:%s>' % (self.get_tab(), self.get_name()))
		value.append('%s<w:id w:val="%s"/>' % (self.get_tab(1), self.get_id()))
		value.append(self.get_part_object().get_xml())
		value.append('%s</w:%s>' % (self.get_tab(), self.get_name()))

		return self.get_separator().join(value)


class Content(object):

	def __init__(self, parent, title, text_separator, font_format, font_size, horizontal_alignment):
		self.name = 'sdtContent'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.title = title
		self.text_separator = text_separator
		_x = self
		while getattr(_x, 'tag', '') != 'document':
			_x = _x.parent
		_x.idx += 1
		self.paragraph = paragraph.Paragraph(self, _x.idx, title, font_format=font_format, font_size=font_size,
												horizontal_alignment=horizontal_alignment)
		t1 = self.paragraph.add_text((), font_format, font_size)
		t1.set_field_char('begin')

		t2 = self.paragraph.add_text((), font_format, font_size)
		t2.set_instr_text('PAGE')
		t3 = self.paragraph.add_text((), font_format, font_size)
		t3.set_field_char('separate')
		t4 = self.paragraph.add_text('1', font_format, font_size)
		t4.get_properties().set_font_size(font_size)
		t4.get_properties().set_font_format(font_format)
		t4.get_properties().set_font_format_cs(font_format)
		t4.get_properties().set_font_size_cs(font_size)
		t5 = self.paragraph.add_text((), font_format, font_size)
		t5.set_field_char('end')
		self.paragraph.add_text(text_separator, font_format, font_size)
		t6 = self.paragraph.add_text((), font_format, font_size)
		t6.set_field_char('begin')
		t7 = self.paragraph.add_text((), font_format, font_size)
		t7.set_instr_text('NUMPAGES')
		t8 = self.paragraph.add_text((), font_format, font_size)
		t8.set_field_char('separate')
		t9 = self.paragraph.add_text('1', font_format, font_size)
		t9.get_properties().set_font_size(font_size)
		t9.get_properties().set_font_format(font_format)
		t9.get_properties().set_font_format_cs(font_format)
		t9.get_properties().set_font_size_cs(font_size)
		t10 = self.paragraph.add_text((), font_format, font_size)
		t10.set_field_char('end')

	def get_parent(self):
		return self.parent

	def set_title(self, title):
		self.title = title

	def get_title(self):
		return self.title

	def set_text_separator(self, text_separator):
		self.text_separator = text_separator

	def get_text_separator(self):
		return self.text_separator

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def get_separator(self):
		return self.separator

	def get_name(self):
		return self.name

	def get_paragraph(self):
		return self.paragraph

	def get_xml(self):
		value = list()
		value.append('%s<w:%s>' % (self.get_tab(), self.get_name()))
		value.append(self.get_paragraph().get_xml())
		value.append('%s</w:%s>' % (self.get_tab(), self.get_name()))

		return self.get_separator().join(value)


class PageNumber(object):

	def __init__(self, parent, title='', text_separator=' / ', font_format='b', font_size=10, horizontal_alignment='r'):
		self.name = 'sdt'
		self.parent = parent
		self.tab = parent.tab
		self.separator = parent.separator
		self.indent = parent.indent + 1
		self.title = title
		self.text_separator = text_separator

		self.properties = Properties(self)
		self.content = Content(self, title, text_separator, font_format, font_size, horizontal_alignment)

	def get_parent(self):
		return self.parent

	def get_tab(self, number=0):
		return self.tab * (self.indent + number)

	def set_title(self, title):
		self.title = title

	def get_title(self):
		return self.title

	def set_text_separator(self, text_separator):
		self.text_separator = text_separator

	def get_text_separator(self):
		return self.text_separator

	def get_properties(self):
		return self.properties

	def get_content(self):
		return self.content

	def get_separator(self):
		return self.separator

	def get_name(self):
		return self.name

	def get_xml(self):
		value = list()

		value.append('%s<w:%s>' % (self.get_tab(), self.get_name()))
		value.append(self.get_properties().get_xml())
		value.append(self.get_content().get_xml())
		value.append('%s</w:%s>' % (self.get_tab(), self.get_name()))

		return self.get_separator().join(value)
