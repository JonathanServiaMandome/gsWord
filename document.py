#!/usr/bin/python
# -*- coding: utf-8 -*-

from zipfile import ZipFile  # , ZIP_DEFLATED
# from shutil import copyfile

import os

import parts
from parts import content_types

from parts import rels
from parts import word

from pagenumber import PageNumber


def new_page_number(parent, title):
	return PageNumber(parent, title)


class Document:
	def __init__(self, ruta_plantilla='', plantilla_=''):
		self.header_num = 0
		self.footer_num = 0
		self.tab = '\t'
		self.separator = '\n'
		self.tag = 'document'
		self._debug = False
		self.ruta_plantilla = ruta_plantilla
		self._plantilla = plantilla_
		self.variables = dict()
		# ./
		self.content_types = content_types.ContentType(self)

		# Componentes
		self.rId = 1
		self.parts = dict()

		self.idx = 1
		self.images = list()
		self.indent = 0

	def get_images(self):
		return self.images

	def add_image(self, target, image, rid):
		self.images.append([target, image, rid])

	def get_rid(self):
		"""Devuelve el RId"""
		return self.rId

	def get_content_types(self):
		return self.content_types

	def get_parts(self):
		return self.parts

	def set_parts(self, dc):
		self.parts = dc

	def add_part_rel(self, name):
		if name in self.parts.keys():
			raise ValueError("La parte %s ya está añadida al documento." % name)
		rel = parts.documentrels.documentRels.DocumentRels(self)
		rel.name = rel.name.replace('document.xml', name)

		self.parts[name] = rel

	def get_part(self, name):
		return self.parts[name]

	def set_part(self, name, dc):
		self.parts[name] = dc

	def get_path(self):
		return self.ruta_plantilla

	def get_file_name(self):
		return self._plantilla

	def set_path(self, value):
		self.ruta_plantilla = value

	def set_file_name(self, value):
		self._plantilla = value

	def get_variables(self):
		return self.variables

	def get_variable(self, name):
		return self.variables.get(name, '')

	def set_variables(self, dc):
		self.variables = dc

	def update_variables(self, dc):
		self.variables.update(dc)

	def get_tab(self):
		return self.tab

	def get_name(self):
		return self.set_tag()

	def get_tag(self):
		return self.tag

	def set_tag(self):
		return self.tag

	def get_separator(self):
		return self.separator

	def get_body(self):
		return self.parts["body"]

	def get_default_header(self):
		header_ = self.parts.get('header2')
		if header_ is None:
			header_ = self.parts.get('header1')
		return header_

	def get_default_footer(self):
		footer_ = self.parts.get('footer2')
		if footer_ is None:
			footer_ = self.parts.get('footer1')
		return footer_

	def new_table(self, parent, data=(), titles=(), column_width=None, horizontal_alignment=(), borders=None):
		self.idx += 1
		return parts.word.elements.table.Table(parent, self.idx, data, titles, column_width, horizontal_alignment,
		                                       borders)

	def new_paragraph(self, parent, text=(), horizontal_alignment='j', font_format='', font_size=None, null=False):
		self.idx += 1
		return parts.word.elements.paragraph.Paragraph(parent, self.idx, text, horizontal_alignment, font_format,
		                                               font_size,
		                                               nulo=null)

	def new_image(self, parent, path, width, heigth, anchor='inline', horizontal_alignment='l'):
		self.idx += 1
		pa = parts.word.elements.paragraph.Paragraph(None, self.idx, horizontal_alignment=horizontal_alignment)
		pa.AddPicture(parent, path, width, heigth, anchor)
		return pa

	def add_part(self, name, value):
		if name in self.parts.keys():
			raise ValueError("La parte %s ya está añadida al documento." % name)
		self.parts[name] = value

	def add_header_section(self, type_part='default'):

		self.header_num += 1
		self.parts["header%d" % self.header_num] = word.part.Part(self, 'hdr', self.header_num, type_part)
		self.parts["header%d" % self.header_num].set_rid(self.idx)
		section = self.get_part('body').get_active_section()
		section.add_header_reference(type_part, self.idx)
		self.idx += 1

	def empty_document(self, headers=True):
		# ./word
		self.parts["font_table"] = word.fonttable.FontTable(self)
		self.parts["font_table"].set_rid(self.idx)
		self.idx += 1
		self.parts["settings"] = word.settings.Settings(self)
		self.parts["settings"].set_rid(self.idx)
		self.idx += 1
		self.parts["web_settings"] = word.websettings.WebSettings(self)
		self.parts["web_settings"].set_rid(self.idx)
		self.idx += 1
		self.parts["style"] = word.styles.Styles(self)
		self.idx += 1

		self.parts["body"] = word.body.Body(self)
		section = self.parts["body"].add_section()
		if headers:
			# for _type in ['even', 'default', 'first']:
			for _type in ['default']:
				self.header_num += 1
				self.parts["header%d" % self.header_num] = word.part.Part(self, 'hdr', self.header_num, _type)
				self.parts["header%d" % self.header_num].set_rid(self.idx)
				section.add_header_reference(_type, self.idx)
				self.idx += 1

				self.footer_num += 1
				self.parts["footer%d" % self.footer_num] = word.part.Part(self, 'ftr', self.footer_num, _type)
				self.parts["footer%d" % self.header_num].set_rid(self.idx)
				section.add_footer_reference(_type, self.idx)
				self.idx += 1

			self.parts["footernotes"] = word.part.Notes(self, 'footernotes')
			self.parts["footernotes"].set_rid(self.idx)
			self.idx += 1
			self.parts["endnotes"] = word.part.Notes(self, 'endnotes')
			self.parts["endnotes"].set_rid(self.idx)
			self.idx += 1

		print section.get_HeaderReferences()
		# ./word/theme
		self.parts["theme1"] = word.theme1.Theme1(self)
		self.parts["theme1"].DefaultValues()
		self.parts["theme1"].set_rid(self.idx)
		self.idx += 1
		# ./rels
		self.parts["rels"] = rels.Rels(self)
		self.parts["rels"].DefaultValues()
		# ./docProps
		self.parts["app"] = parts.docProps.app.App(self)
		self.parts["app"].DefaultValues()
		self.parts["core"] = parts.docProps.core.Core(self)
		self.parts["core"].DefaultValues()

	def create_document_rels(self):
		doc_rels = word.documentRels.DocumentRels(self)

		for part_name in self.get_parts().keys():
			part = self.get_part(part_name)

			if hasattr(part, 'set_rid'):
				pass
				# part.set_rid(self.rId)
				# self.rId += 1
			else:
				continue
			name = part.get_name().replace('word/', '')
			name_ = name.replace('media/', '')
			if '.' in name_:
				name_ = name_.split('.')[0]
			while name_[-1].isdigit():
				name_ = name_[:-1]
			dc = {'Id': 'rId%d' % part.get_rid(),
			      "Type": 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/%s' % name_}
			doc_rels.add_part(name, dc)

		for img in self.get_images():
			dc = {'Id': 'rId%d' % img[2],
			      "Type": 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'}
			doc_rels.add_part(img[0], dc)
		return doc_rels

	# noinspection PyBroadException
	def save(self):
		# ./word/rels
		self.parts["document_rels"] = self.create_document_rels()

		# self.parts["body"].AddPrincipalSection()
		try:
			os.makedirs(self.get_path())
		except Exception:
			pass

		try:
			zout = ZipFile(self.get_path() + self.get_file_name(), 'w')
		except IOError as e:
			if e.args and e.args[0] == 13:
				raise ValueError("Tiene abierto un documento con el mismo nombre y en la misma ubicación.")
			elif e.args and e.args[0] == 32:
				raise ValueError("Tiene abierto un documento con el mismo nombre y en la misma ubicación.")
			else:
				raise ValueError(e)
		except Exception as e:
			raise ValueError(e)

		for img in self.get_images():
			zout.writestr('word/' + img[0], img[1])

		for part in self.parts.items():
			name, part = part

			if not hasattr(part, 'get_xml'):
				continue

			zout.writestr(part.get_name(), part.get_xml())

			if hasattr(part, 'get_Images'):
				for img in part.get_images():
					zout.writestr('word/' + img[0], img[1])

			if self._debug:
				try:
					os.makedirs(self.get_path() + 'temp/')
				except Exception:
					pass

				try:
					part_split = part.get_name().split('/')
					os.makedirs(self.get_path() + 'temp/' + '/'.join(part_split[:-1]))
				except Exception:
					pass
				with open(self.get_path() + 'temp/' + part.get_name(), 'w') as file_part:
					file_part.write(part.get_xml())

		zout.writestr(self.get_content_types().get_name(), self.get_content_types().get_xml())

		if self._debug:
			with open(self.get_path() + 'temp/' + self.get_content_types().get_name(), 'w') as file_part:
				file_part.write(self.get_content_types().get_xml())
		zout.close()

	def get_xml(self):
		value = list()
		value.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
		value.append('</w:%s>' % self.set_tag())
		try:
			open('c:/users/jonathan/desktop/body.txt', 'w').write(self.separator.join(value))
			os.system('start c:/users/jonathan/desktop/body.txt')
		except IOError:
			pass
		return self.separator.join(value)
