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

		self.idx = 0
		self.idx_header = 1
		self.idx_footer = 1
		self.images = list()
		self.indent = 0

	def get_Images(self):
		return self.images

	def AddImage(self, target, image, rid):
		self.images.append([target, image, rid])

	def get_RId(self):
		"""Devuelve el RId"""
		return self.rId

	def AddRId(self):
		self.rId += 1

	def get_ContentTypes(self):
		return self.content_types

	def get_parts(self):
		return self.parts

	def SetParts(self, dc):
		self.parts = dc

	def add_part(self, name, value):
		if name in self.parts.keys():
			raise ValueError("La parte %s ya está añadida al documento." % name)
		self.parts[name] = value

	def add_part_rel(self, name):
		if name in self.parts.keys():
			raise ValueError("La parte %s ya está añadida al documento." % name)
		rel = parts.documentrels.documentRels.DocumentRels(self)
		rel.name = rel.name.replace('document.xml', name)

		self.parts[name] = rel

	def get_part(self, name):
		return self.parts[name]

	def SetPart(self, name, dc):
		self.parts[name] = dc

	def get_RutaPlantilla(self):
		return self.ruta_plantilla

	def get_Plantilla(self):
		return self._plantilla

	def SetRutaPlantilla(self, value):
		self.ruta_plantilla = value

	def SetPlantilla(self, value):
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
		return self.get_tag()

	def get_tag(self):
		return self.tag

	def get_Tag(self):
		return self.tag

	def get_separator(self):
		return self.separator

	def get_body(self):
		return self.parts["body"]

	def get_DefaultHeader(self):
		header_ = self.parts.get('header2')
		if header_ is None:
			header_ = self.parts.get('header1')
		return header_

	def get_DefaultFooter(self):
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

	def empty_document(self, headers=True):
		# ./word
		self.parts["font_table"] = word.fonttable.FontTable(self)
		self.idx += 1
		self.parts["settings"] = word.settings.Settings(self)
		self.idx += 1
		self.parts["web_settings"] = word.websettings.WebSettings(self)
		self.idx += 1
		self.parts["style"] = word.styles.Styles(self)
		self.idx += 1
		self.parts["body"] = word.body.Body(self)
		section = self.parts["body"].get_active_section()
		if headers:
			self.parts["header1"] = word.part.Part(self, 'hdr', 1, 'even')
			section.AddHeaderReference(self.parts["header1"].get_TypeReference(), self.parts["header1"])
			self.idx += 1
			self.idx_header += 1
			self.parts["header2"] = word.part.Part(self, 'hdr', 2, 'default')
			section.AddHeaderReference(self.parts["header2"].get_TypeReference(), self.parts["header2"])
			self.idx += 1
			self.idx_header += 1
			self.parts["footer1"] = word.part.Part(self, 'ftr', 1, 'even')
			section.AddFooterReference(self.parts["footer1"].get_TypeReference(), self.parts["footer1"])
			self.idx += 1
			self.idx_footer += 1
			self.parts["footer2"] = word.part.Part(self, 'ftr', 2, 'default')
			section.AddFooterReference(self.parts["footer2"].get_TypeReference(), self.parts["footer2"])
			self.idx += 1
			self.idx_footer += 1
			self.parts["header3"] = word.part.Part(self, 'hdr', 3, 'first')
			section.AddHeaderReference(self.parts["header3"].get_TypeReference(), self.parts["header3"])
			self.idx += 1
			self.idx_header += 1
			self.parts["footer3"] = word.part.Part(self, 'ftr', 3, 'first')
			section.AddFooterReference(self.parts["footer3"].get_TypeReference(), self.parts["footer3"])
			self.idx += 1
			self.idx_footer += 1
			self.parts["footernotes"] = word.part.Notes(self, 'footernotes')
			self.idx += 1
			self.parts["endnotes"] = word.part.Notes(self, 'endnotes')
			self.idx += 1

		# ./word/theme
		self.parts["theme1"] = word.theme1.Theme1(self)
		self.parts["theme1"].DefaultValues()
		self.idx += 1
		# ./rels
		self.parts["rels"] = rels.Rels(self)
		self.parts["rels"].DefaultValues()
		# ./docProps
		self.parts["app"] = parts.docProps.app.App(self)
		self.parts["app"].DefaultValues()
		self.parts["core"] = parts.docProps.core.Core(self)
		self.parts["core"].DefaultValues()

	def CreateDocumentRels(self):
		doc_rels = word.documentRels.DocumentRels(self)

		for part_name in self.get_parts().keys():
			part = self.get_part(part_name)

			if hasattr(part, 'SetRId'):
				part.SetRId(self.rId)
				self.rId += 1
			else:
				continue
			name = part.get_name().replace('word/', '')
			name_ = name.replace('media/', '')
			if '.' in name_:
				name_ = name_.split('.')[0]
			while name_[-1].isdigit():
				name_ = name_[:-1]
			dc = {
				'Id': 'rId%d' % part.get_RId(),
				"Type": 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/%s' % name_
			}
			doc_rels.add_part(name, dc)

		for img in self.get_Images():
			dc = {
				'Id': 'rId%d' % img[2],
				"Type": 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
			}
			doc_rels.add_part(img[0], dc)
		return doc_rels

	def new_header(self, section, _type='default'):
		number = self.idx_header
		self.idx_header += 1
		self.parts["header%d" % number] = word.part.Part(self, 'hdr', number, _type)

		section.AddHeaderReference(self.parts["header%d" % number].get_TypeReference(), self.parts["header%d" % number])

		self.parts["header%d" % number].SetRId((100 * number) + 200)

		return self.parts["header%d" % number]

	# noinspection PyBroadException
	def save(self):
		# ./word/rels
		self.parts["document_rels"] = self.CreateDocumentRels()

		# self.parts["body"].AddPrincipalSection()
		try:
			os.makedirs(self.get_RutaPlantilla())
		except Exception:
			pass

		try:
			zout = ZipFile(self.get_RutaPlantilla() + self.get_Plantilla(), 'w')
		except IOError as e:
			if e.args and e.args[0] == 13:
				raise ValueError("Tiene abierto un documento con el mismo nombre y en la misma ubicación.")
			else:
				raise ValueError(e)
		except Exception as e:
			raise ValueError(e)

		for img in self.get_Images():
			zout.writestr('word/' + img[0], img[1])

		for part in self.parts.items():
			name, part = part

			if not hasattr(part, 'get_xml'):
				continue

			zout.writestr(part.get_name(), part.get_xml())

			if hasattr(part, 'get_Images'):
				for img in part.get_Images():
					zout.writestr('word/' + img[0], img[1])

			if self._debug:
				try:
					os.makedirs(self.get_RutaPlantilla() + 'temp/')
				except Exception:
					pass

				try:
					part_split = part.get_name().split('/')
					os.makedirs(self.get_RutaPlantilla() + 'temp/' + '/'.join(part_split[:-1]))
				except Exception:
					pass
				with open(self.get_RutaPlantilla() + 'temp/' + part.get_name(), 'w') as file_part:
					file_part.write(part.get_xml())

		zout.writestr(self.get_ContentTypes().get_name(), self.get_ContentTypes().get_xml())

		if self._debug:
			with open(self.get_RutaPlantilla() + 'temp/' + self.get_ContentTypes().get_name(), 'w') as file_part:
				file_part.write(self.get_ContentTypes().get_xml())
		zout.close()

	def get_xml(self):
		value = list()
		value.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
		value.append('</w:%s>' % self.get_Tag())
		try:
			open('c:/users/jonathan/desktop/body.txt', 'w').write(self.separator.join(value))
			os.system('start c:/users/jonathan/desktop/body.txt')
		except IOError:
			pass
		return self.separator.join(value)
