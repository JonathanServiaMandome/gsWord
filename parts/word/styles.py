#!/usr/bin/python
# -*- coding: utf-8 -*-
"""
@author: Jonathan Servia Mandome
@version: 0.1.0
Created on 15/03/2020
Genera la hoja de estilos por defecto del word

@version: 0.2.0
Updated on 31/03/2021
Se crea como objeto para poder modificar sus propiedades y crear estilos personalizados
"""
from .elements import table, text, paragraph


class Styles(object):
	class LatentStyle(object):
		def __init__(self, parent):
			self.name = 'latentStyles'
			self.subname = 'lsdException'
			self.parent = parent
			self.tab = parent.tab
			self.indent = parent.indent + 1
			self.separator = parent.separator
			self.values = list()

			self.values.append({"Normal": {'uiPriority': "0", 'qFormat': "1"}})
			self.values.append({"heading 1": {'uiPriority': "9", 'qFormat': "1"}})
			self.values.append(
				{"heading 2": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append(
				{"heading 3": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append(
				{"heading 4": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append(
				{"heading 5": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append(
				{"heading 6": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append(
				{"heading 7": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append(
				{"heading 8": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append(
				{"heading 9": {'semiHidden': "1", 'uiPriority': "9", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append({"index 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 6": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 7": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 8": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index 9": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 1": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 2": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 3": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 4": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 5": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 6": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 7": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 8": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"toc 9": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1"}})
			self.values.append({"Normal Indent": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"footnote text": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"annotation text": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"header": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"footer": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"index heading": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append(
				{"caption": {'semiHidden': "1", 'uiPriority': "35", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append({"table of figures": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"envelope address": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"envelope return": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"footnote reference": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"annotation reference": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"line number": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"page number": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"endnote reference": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"endnote text": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"table of authorities": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"macro": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"toa heading": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Bullet": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Number": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Bullet 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Bullet 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Bullet 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Bullet 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Number 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Number 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Number 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Number 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Title": {'uiPriority': "10", 'qFormat': "1"}})
			self.values.append({"Closing": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Signature": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append(
				{"Default Paragraph Font": {'semiHidden': "1", 'uiPriority': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text Indent": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Continue": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Continue 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Continue 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Continue 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"List Continue 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Message Header": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Subtitle": {'uiPriority': "11", 'qFormat': "1"}})
			self.values.append({"Salutation": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Date": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text First Indent": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text First Indent 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Note Heading": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text Indent 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Body Text Indent 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Block Text": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Hyperlink": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"FollowedHyperlink": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Strong": {'uiPriority': "22", 'qFormat': "1"}})
			self.values.append({"Emphasis": {'uiPriority': "20", 'qFormat': "1"}})
			self.values.append({"Document Map": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Plain Text": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"E-mail Signature": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Top of Form": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Bottom of Form": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Normal (Web)": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Acronym": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Address": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Cite": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Code": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Definition": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Keyboard": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Preformatted": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Sample": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Typewriter": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"HTML Variable": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Normal Table": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"annotation subject": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"No List": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Outline List 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Outline List 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Outline List 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Simple 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Simple 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Simple 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Classic 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Classic 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Classic 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Classic 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Colorful 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Colorful 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Colorful 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Columns 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Columns 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Columns 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Columns 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Columns 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 6": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 7": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid 8": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 4": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 5": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 6": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 7": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table List 8": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table 3D effects 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table 3D effects 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table 3D effects 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Contemporary": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Elegant": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Professional": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Subtle 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Subtle 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Web 1": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Web 2": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Web 3": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Balloon Text": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Table Grid": {'uiPriority': "39"}})
			self.values.append({"Table Theme": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Placeholder Text": {'semiHidden': "1"}})
			self.values.append({"No Spacing": {'uiPriority': "1", 'qFormat': "1"}})
			self.values.append({"Light Shading": {'uiPriority': "60"}})
			self.values.append({"Light List": {'uiPriority': "61"}})
			self.values.append({"Light Grid": {'uiPriority': "62"}})
			self.values.append({"Medium Shading 1": {'uiPriority': "63"}})
			self.values.append({"Medium Shading 2": {'uiPriority': "64"}})
			self.values.append({"Medium List 1": {'uiPriority': "65"}})
			self.values.append({"Medium List 2": {'uiPriority': "66"}})
			self.values.append({"Medium Grid 1": {'uiPriority': "67"}})
			self.values.append({"Medium Grid 2": {'uiPriority': "68"}})
			self.values.append({"Medium Grid 3": {'uiPriority': "69"}})
			self.values.append({"Dark List": {'uiPriority': "70"}})
			self.values.append({"Colorful Shading": {'uiPriority': "71"}})
			self.values.append({"Colorful List": {'uiPriority': "72"}})
			self.values.append({"Colorful Grid": {'uiPriority': "73"}})
			self.values.append({"Light Shading Accent 1": {'uiPriority': "60"}})
			self.values.append({"Light List Accent 1": {'uiPriority': "61"}})
			self.values.append({"Light Grid Accent 1": {'uiPriority': "62"}})
			self.values.append({"Medium Shading 1 Accent 1": {'uiPriority': "63"}})
			self.values.append({"Medium Shading 2 Accent 1": {'uiPriority': "64"}})
			self.values.append({"Medium List 1 Accent 1": {'uiPriority': "65"}})
			self.values.append({"Revision": {'semiHidden': "1"}})
			self.values.append({"List Paragraph": {'uiPriority': "34", 'qFormat': "1"}})
			self.values.append({"Quote": {'uiPriority': "29", 'qFormat': "1"}})
			self.values.append({"Intense Quote": {'uiPriority': "30", 'qFormat': "1"}})
			self.values.append({"Medium List 2 Accent 1": {'uiPriority': "66"}})
			self.values.append({"Medium Grid 1 Accent 1": {'uiPriority': "67"}})
			self.values.append({"Medium Grid 2 Accent 1": {'uiPriority': "68"}})
			self.values.append({"Medium Grid 3 Accent 1": {'uiPriority': "69"}})
			self.values.append({"Dark List Accent 1": {'uiPriority': "70"}})
			self.values.append({"Colorful Shading Accent 1": {'uiPriority': "71"}})
			self.values.append({"Colorful List Accent 1": {'uiPriority': "72"}})
			self.values.append({"Colorful Grid Accent 1": {'uiPriority': "73"}})
			self.values.append({"Light Shading Accent 2": {'uiPriority': "60"}})
			self.values.append({"Light List Accent 2": {'uiPriority': "61"}})
			self.values.append({"Light Grid Accent 2": {'uiPriority': "62"}})
			self.values.append({"Medium Shading 1 Accent 2": {'uiPriority': "63"}})
			self.values.append({"Medium Shading 2 Accent 2": {'uiPriority': "64"}})
			self.values.append({"Medium List 1 Accent 2": {'uiPriority': "65"}})
			self.values.append({"Medium List 2 Accent 2": {'uiPriority': "66"}})
			self.values.append({"Medium Grid 1 Accent 2": {'uiPriority': "67"}})
			self.values.append({"Medium Grid 2 Accent 2": {'uiPriority': "68"}})
			self.values.append({"Medium Grid 3 Accent 2": {'uiPriority': "69"}})
			self.values.append({"Dark List Accent 2": {'uiPriority': "70"}})
			self.values.append({"Colorful Shading Accent 2": {'uiPriority': "71"}})
			self.values.append({"Colorful List Accent 2": {'uiPriority': "72"}})
			self.values.append({"Colorful Grid Accent 2": {'uiPriority': "73"}})
			self.values.append({"Light Shading Accent 3": {'uiPriority': "60"}})
			self.values.append({"Light List Accent 3": {'uiPriority': "61"}})
			self.values.append({"Light Grid Accent 3": {'uiPriority': "62"}})
			self.values.append({"Medium Shading 1 Accent 3": {'uiPriority': "63"}})
			self.values.append({"Medium Shading 2 Accent 3": {'uiPriority': "64"}})
			self.values.append({"Medium List 1 Accent 3": {'uiPriority': "65"}})
			self.values.append({"Medium List 2 Accent 3": {'uiPriority': "66"}})
			self.values.append({"Medium Grid 1 Accent 3": {'uiPriority': "67"}})
			self.values.append({"Medium Grid 2 Accent 3": {'uiPriority': "68"}})
			self.values.append({"Medium Grid 3 Accent 3": {'uiPriority': "69"}})
			self.values.append({"Dark List Accent 3": {'uiPriority': "70"}})
			self.values.append({"Colorful Shading Accent 3": {'uiPriority': "71"}})
			self.values.append({"Colorful List Accent 3": {'uiPriority': "72"}})
			self.values.append({"Colorful Grid Accent 3": {'uiPriority': "73"}})
			self.values.append({"Light Shading Accent 4": {'uiPriority': "60"}})
			self.values.append({"Light List Accent 4": {'uiPriority': "61"}})
			self.values.append({"Light Grid Accent 4": {'uiPriority': "62"}})
			self.values.append({"Medium Shading 1 Accent 4": {'uiPriority': "63"}})
			self.values.append({"Medium Shading 2 Accent 4": {'uiPriority': "64"}})
			self.values.append({"Medium List 1 Accent 4": {'uiPriority': "65"}})
			self.values.append({"Medium List 2 Accent 4": {'uiPriority': "66"}})
			self.values.append({"Medium Grid 1 Accent 4": {'uiPriority': "67"}})
			self.values.append({"Medium Grid 2 Accent 4": {'uiPriority': "68"}})
			self.values.append({"Medium Grid 3 Accent 4": {'uiPriority': "69"}})
			self.values.append({"Dark List Accent 4": {'uiPriority': "70"}})
			self.values.append({"Colorful Shading Accent 4": {'uiPriority': "71"}})
			self.values.append({"Colorful List Accent 4": {'uiPriority': "72"}})
			self.values.append({"Colorful Grid Accent 4": {'uiPriority': "73"}})
			self.values.append({"Light Shading Accent 5": {'uiPriority': "60"}})
			self.values.append({"Light List Accent 5": {'uiPriority': "61"}})
			self.values.append({"Light Grid Accent 5": {'uiPriority': "62"}})
			self.values.append({"Medium Shading 1 Accent 5": {'uiPriority': "63"}})
			self.values.append({"Medium Shading 2 Accent 5": {'uiPriority': "64"}})
			self.values.append({"Medium List 1 Accent 5": {'uiPriority': "65"}})
			self.values.append({"Medium List 2 Accent 5": {'uiPriority': "66"}})
			self.values.append({"Medium Grid 1 Accent 5": {'uiPriority': "67"}})
			self.values.append({"Medium Grid 2 Accent 5": {'uiPriority': "68"}})
			self.values.append({"Medium Grid 3 Accent 5": {'uiPriority': "69"}})
			self.values.append({"Dark List Accent 5": {'uiPriority': "70"}})
			self.values.append({"Colorful Shading Accent 5": {'uiPriority': "71"}})
			self.values.append({"Colorful List Accent 5": {'uiPriority': "72"}})
			self.values.append({"Colorful Grid Accent 5": {'uiPriority': "73"}})
			self.values.append({"Light Shading Accent 6": {'uiPriority': "60"}})
			self.values.append({"Light List Accent 6": {'uiPriority': "61"}})
			self.values.append({"Light Grid Accent 6": {'uiPriority': "62"}})
			self.values.append({"Medium Shading 1 Accent 6": {'uiPriority': "63"}})
			self.values.append({"Medium Shading 2 Accent 6": {'uiPriority': "64"}})
			self.values.append({"Medium List 1 Accent 6": {'uiPriority': "65"}})
			self.values.append({"Medium List 2 Accent 6": {'uiPriority': "66"}})
			self.values.append({"Medium Grid 1 Accent 6": {'uiPriority': "67"}})
			self.values.append({"Medium Grid 2 Accent 6": {'uiPriority': "68"}})
			self.values.append({"Medium Grid 3 Accent 6": {'uiPriority': "69"}})
			self.values.append({"Dark List Accent 6": {'uiPriority': "70"}})
			self.values.append({"Colorful Shading Accent 6": {'uiPriority': "71"}})
			self.values.append({"Colorful List Accent 6": {'uiPriority': "72"}})
			self.values.append({"Colorful Grid Accent 6": {'uiPriority': "73"}})
			self.values.append({"Subtle Emphasis": {'uiPriority': "19", 'qFormat': "1"}})
			self.values.append({"Intense Emphasis": {'uiPriority': "21", 'qFormat': "1"}})
			self.values.append({"Subtle Reference": {'uiPriority': "31", 'qFormat': "1"}})
			self.values.append({"Intense Reference": {'uiPriority': "32", 'qFormat': "1"}})
			self.values.append({"Book Title": {'uiPriority': "33", 'qFormat': "1"}})
			self.values.append({"Bibliography": {'semiHidden': "1", 'uiPriority': "37", 'unhideWhenUsed': "1"}})
			self.values.append(
				{"TOC Heading": {'semiHidden': "1", 'uiPriority': "39", 'unhideWhenUsed': "1", 'qFormat': "1"}})
			self.values.append({"Plain Table 1": {'uiPriority': "41"}})
			self.values.append({"Plain Table 2": {'uiPriority': "42"}})
			self.values.append({"Plain Table 3": {'uiPriority': "43"}})
			self.values.append({"Plain Table 4": {'uiPriority': "44"}})
			self.values.append({"Plain Table 5": {'uiPriority': "45"}})
			self.values.append({"Grid Table Light": {'uiPriority': "40"}})
			self.values.append({"Grid Table 1 Light": {'uiPriority': "46"}})
			self.values.append({"Grid Table 2": {'uiPriority': "47"}})
			self.values.append({"Grid Table 3": {'uiPriority': "48"}})
			self.values.append({"Grid Table 4": {'uiPriority': "49"}})
			self.values.append({"Grid Table 5 Dark": {'uiPriority': "50"}})
			self.values.append({"Grid Table 6 Colorful": {'uiPriority': "51"}})
			self.values.append({"Grid Table 7 Colorful": {'uiPriority': "52"}})
			self.values.append({"Grid Table 1 Light Accent 1": {'uiPriority': "46"}})
			self.values.append({"Grid Table 2 Accent 1": {'uiPriority': "47"}})
			self.values.append({"Grid Table 3 Accent 1": {'uiPriority': "48"}})
			self.values.append({"Grid Table 4 Accent 1": {'uiPriority': "49"}})
			self.values.append({"Grid Table 5 Dark Accent 1": {'uiPriority': "50"}})
			self.values.append({"Grid Table 6 Colorful Accent 1": {'uiPriority': "51"}})
			self.values.append({"Grid Table 7 Colorful Accent 1": {'uiPriority': "52"}})
			self.values.append({"Grid Table 1 Light Accent 2": {'uiPriority': "46"}})
			self.values.append({"Grid Table 2 Accent 2": {'uiPriority': "47"}})
			self.values.append({"Grid Table 3 Accent 2": {'uiPriority': "48"}})
			self.values.append({"Grid Table 4 Accent 2": {'uiPriority': "49"}})
			self.values.append({"Grid Table 5 Dark Accent 2": {'uiPriority': "50"}})
			self.values.append({"Grid Table 6 Colorful Accent 2": {'uiPriority': "51"}})
			self.values.append({"Grid Table 7 Colorful Accent 2": {'uiPriority': "52"}})
			self.values.append({"Grid Table 1 Light Accent 3": {'uiPriority': "46"}})
			self.values.append({"Grid Table 2 Accent 3": {'uiPriority': "47"}})
			self.values.append({"Grid Table 3 Accent 3": {'uiPriority': "48"}})
			self.values.append({"Grid Table 4 Accent 3": {'uiPriority': "49"}})
			self.values.append({"Grid Table 5 Dark Accent 3": {'uiPriority': "50"}})
			self.values.append({"Grid Table 6 Colorful Accent 3": {'uiPriority': "51"}})
			self.values.append({"Grid Table 7 Colorful Accent 3": {'uiPriority': "52"}})
			self.values.append({"Grid Table 1 Light Accent 4": {'uiPriority': "46"}})
			self.values.append({"Grid Table 2 Accent 4": {'uiPriority': "47"}})
			self.values.append({"Grid Table 3 Accent 4": {'uiPriority': "48"}})
			self.values.append({"Grid Table 4 Accent 4": {'uiPriority': "49"}})
			self.values.append({"Grid Table 5 Dark Accent 4": {'uiPriority': "50"}})
			self.values.append({"Grid Table 6 Colorful Accent 4": {'uiPriority': "51"}})
			self.values.append({"Grid Table 7 Colorful Accent 4": {'uiPriority': "52"}})
			self.values.append({"Grid Table 1 Light Accent 5": {'uiPriority': "46"}})
			self.values.append({"Grid Table 2 Accent 5": {'uiPriority': "47"}})
			self.values.append({"Grid Table 3 Accent 5": {'uiPriority': "48"}})
			self.values.append({"Grid Table 4 Accent 5": {'uiPriority': "49"}})
			self.values.append({"Grid Table 5 Dark Accent 5": {'uiPriority': "50"}})
			self.values.append({"Grid Table 6 Colorful Accent 5": {'uiPriority': "51"}})
			self.values.append({"Grid Table 7 Colorful Accent 5": {'uiPriority': "52"}})
			self.values.append({"Grid Table 1 Light Accent 6": {'uiPriority': "46"}})
			self.values.append({"Grid Table 2 Accent 6": {'uiPriority': "47"}})
			self.values.append({"Grid Table 3 Accent 6": {'uiPriority': "48"}})
			self.values.append({"Grid Table 4 Accent 6": {'uiPriority': "49"}})
			self.values.append({"Grid Table 5 Dark Accent 6": {'uiPriority': "50"}})
			self.values.append({"Grid Table 6 Colorful Accent 6": {'uiPriority': "51"}})
			self.values.append({"Grid Table 7 Colorful Accent 6": {'uiPriority': "52"}})
			self.values.append({"List Table 1 Light": {'uiPriority': "46"}})
			self.values.append({"List Table 2": {'uiPriority': "47"}})
			self.values.append({"List Table 3": {'uiPriority': "48"}})
			self.values.append({"List Table 4": {'uiPriority': "49"}})
			self.values.append({"List Table 5 Dark": {'uiPriority': "50"}})
			self.values.append({"List Table 6 Colorful": {'uiPriority': "51"}})
			self.values.append({"List Table 7 Colorful": {'uiPriority': "52"}})
			self.values.append({"List Table 1 Light Accent 1": {'uiPriority': "46"}})
			self.values.append({"List Table 2 Accent 1": {'uiPriority': "47"}})
			self.values.append({"List Table 3 Accent 1": {'uiPriority': "48"}})
			self.values.append({"List Table 4 Accent 1": {'uiPriority': "49"}})
			self.values.append({"List Table 5 Dark Accent 1": {'uiPriority': "50"}})
			self.values.append({"List Table 6 Colorful Accent 1": {'uiPriority': "51"}})
			self.values.append({"List Table 7 Colorful Accent 1": {'uiPriority': "52"}})
			self.values.append({"List Table 1 Light Accent 2": {'uiPriority': "46"}})
			self.values.append({"List Table 2 Accent 2": {'uiPriority': "47"}})
			self.values.append({"List Table 3 Accent 2": {'uiPriority': "48"}})
			self.values.append({"List Table 4 Accent 2": {'uiPriority': "49"}})
			self.values.append({"List Table 5 Dark Accent 2": {'uiPriority': "50"}})
			self.values.append({"List Table 6 Colorful Accent 2": {'uiPriority': "51"}})
			self.values.append({"List Table 7 Colorful Accent 2": {'uiPriority': "52"}})
			self.values.append({"List Table 1 Light Accent 3": {'uiPriority': "46"}})
			self.values.append({"List Table 2 Accent 3": {'uiPriority': "47"}})
			self.values.append({"List Table 3 Accent 3": {'uiPriority': "48"}})
			self.values.append({"List Table 4 Accent 3": {'uiPriority': "49"}})
			self.values.append({"List Table 5 Dark Accent 3": {'uiPriority': "50"}})
			self.values.append({"List Table 6 Colorful Accent 3": {'uiPriority': "51"}})
			self.values.append({"List Table 7 Colorful Accent 3": {'uiPriority': "52"}})
			self.values.append({"List Table 1 Light Accent 4": {'uiPriority': "46"}})
			self.values.append({"List Table 2 Accent 4": {'uiPriority': "47"}})
			self.values.append({"List Table 3 Accent 4": {'uiPriority': "48"}})
			self.values.append({"List Table 4 Accent 4": {'uiPriority': "49"}})
			self.values.append({"List Table 5 Dark Accent 4": {'uiPriority': "50"}})
			self.values.append({"List Table 6 Colorful Accent 4": {'uiPriority': "51"}})
			self.values.append({"List Table 7 Colorful Accent 4": {'uiPriority': "52"}})
			self.values.append({"List Table 1 Light Accent 5": {'uiPriority': "46"}})
			self.values.append({"List Table 2 Accent 5": {'uiPriority': "47"}})
			self.values.append({"List Table 3 Accent 5": {'uiPriority': "48"}})
			self.values.append({"List Table 4 Accent 5": {'uiPriority': "49"}})
			self.values.append({"List Table 5 Dark Accent 5": {'uiPriority': "50"}})
			self.values.append({"List Table 6 Colorful Accent 5": {'uiPriority': "51"}})
			self.values.append({"List Table 7 Colorful Accent 5": {'uiPriority': "52"}})
			self.values.append({"List Table 1 Light Accent 6": {'uiPriority': "46"}})
			self.values.append({"List Table 2 Accent 6": {'uiPriority': "47"}})
			self.values.append({"List Table 3 Accent 6": {'uiPriority': "48"}})
			self.values.append({"List Table 4 Accent 6": {'uiPriority': "49"}})
			self.values.append({"List Table 5 Dark Accent 6": {'uiPriority': "50"}})
			self.values.append({"List Table 6 Colorful Accent 6": {'uiPriority': "51"}})
			self.values.append({"List Table 7 Colorful Accent 6": {'uiPriority': "52"}})
			self.values.append({"Mention": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Smart Hyperlink": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Hashtag": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Unresolved Mention": {'semiHidden': "1", 'unhideWhenUsed': "1"}})
			self.values.append({"Smart Link": {'semiHidden': "1", 'unhideWhenUsed': "1"}})

		def get_tab(self, number=0):
			return self.tab * (number + self.indent)

		def get_name(self):
			return self.name

		def get_subname(self):
			return self.subname

		def get_values(self):
			return self.values

		def set_values(self, dc):
			self.values = dc

		def add_value(self, value):
			self.values.append(value)

		def get_xml(self):
			value = list()
			for _value in self.get_values():
				for _k in _value.keys():
					aux = _value[_k]
					if aux:
						temp = '%s<w:%s w:name="%s"' % (self.get_tab(1), self.get_subname(), _k)
						for k in aux.keys():
							temp += ' w:%s="%s"' % (k, aux[k])
						temp += '/>'
						value.append(temp)
			if value:
				value.insert(0, '%s<w:%s>' % (self.get_tab(0), self.get_name()))
				value.append('%s</w:%s>' % (self.get_tab(0), self.get_name()))
			return self.separator.join(value)

	class DocDefaluts(object):

		class PropertiesDefault(object):
			def __init__(self, parent, name, subname, values):
				self.name = name
				self.subname = subname
				self.parent = parent
				self.tab = parent.tab
				self.indent = parent.indent + 1
				self.separator = parent.separator
				self.values = values

			def get_tab(self, number=0):
				return self.tab * (number + self.indent)

			def get_name(self):
				return self.name

			def get_subname(self):
				return self.subname

			def get_values(self):
				return self.values

			def set_values(self, dc):
				self.values = dc

			def set_value(self, _type, key, value):
				if _type not in self.values.keys():
					self.values[_type] = {}
				self.values[_type][key] = value

			def get_xml(self):
				value = list()
				if self.get_values():
					for _k in self.get_values().keys():
						aux = self.get_values()[_k]
						if aux:
							temp = '%s<w:%s' % (self.get_tab(2), self.get_name())
							for k in aux.keys():
								temp += ' w:%s="%s"' % (k, aux[k])
							temp += '/>'
							value.append(temp)

				if value:
					value.insert(0, '%s<w:%s>' % (self.get_tab(1), self.get_subname()))
					value.insert(0, '%s<w:%s>' % (self.get_tab(0), self.get_name()))
					value.append('%s</w:%s>' % (self.get_tab(1), self.get_subname()))
					value.append('%s</w:%s>' % (self.get_tab(0), self.get_name()))
				return self.separator.join(value)

		def __init__(self, parent):
			self.name = 'docDefaults'
			self.parent = parent
			self.tab = parent.tab
			self.indent = parent.indent + 1
			self.separator = parent.separator
			rprdefault = {
				'rFonts': {
					'asciiTheme': 'minorHAnsi',
					'eastAsiaTheme': 'minorHAnsi',
					'hAnsiTheme': 'minorHAnsi',
					'cstheme': 'minorHAnsi',
				},
				'sz': {
					'val': 20
				},
				'szCs': {
					'val': 19
				},
				'lang': {
					'val': 'es-ES',
					'eastAsia': 'en-US',
					'bidi': 'ar-SA'
				}
			}

			pprdefault = {
				'spacing': {
					'after': "160",
					'line': "259",
					'lineRule': "auto"
				}
			}

			self.rPrDefaults = self.PropertiesDefault(self, 'rPrDefault', 'rPr', rprdefault)
			self.pPrDefaults = self.PropertiesDefault(self, 'pPrDefault', 'pPr', pprdefault)
			self.style = list()

		def get_rpr_defaults(self):
			return self.rPrDefaults

		def set_rpr_defaults(self, _object):
			self.rPrDefaults = _object

		def get_ppr_defaults(self):
			return self.pPrDefaults

		def set_ppr_defaults(self, _object):
			self.pPrDefaults = _object

		def get_name(self):
			return self.name

		def get_tab(self, number=0):
			return self.tab * (number + self.indent)

		def get_xml(self):
			return self.separator.join([
				'%s<w:%s>' % (self.get_tab(), self.get_name()),
				self.get_rpr_defaults().get_xml(),
				self.get_ppr_defaults().get_xml(),
				'%s</w:%s>' % (self.get_tab(), self.get_name())
			])

	class Style(object):
		def __init__(self, parent, _type, style_id, value_name, default='', priority='99', semi_hidden=False,
						unhide_when_used=False, q_format=False, rsid='', based_on='', link='', custom_style=''):
			self.name = 'style'
			self.parent = parent
			self.tab = parent.tab
			self.indent = parent.indent + 1
			self.separator = parent.separator
			self.type = _type
			self.styleId = style_id
			self.default = default
			self.value_name = value_name
			self.priority = priority
			self.semiHidden = semi_hidden
			self.unhideWhenUsed = unhide_when_used
			self.qFormat = q_format
			self.rsid = rsid
			self.basedOn = based_on
			self.link = link
			self.custom_style = custom_style
			self.elements = list()

		def get_name(self):
			return self.name

		def get_type(self):
			return self.type

		def set_type(self, _type):
			self.type = _type

		def get_link(self):
			return self.link

		def set_link(self, link):
			self.link = link

		def get_style_id(self):
			return self.styleId

		def set_style_id(self, style_id):
			self.styleId = style_id

		def get_custom_style(self):
			return self.custom_style

		def set_custom_style(self, custom_style):
			self.custom_style = custom_style

		def get_default(self):
			return self.default

		def set_default(self, default):
			self.default = default

		def get_value_name(self):
			return self.value_name

		def set_value_name(self, value_name):
			self.value_name = value_name

		def get_priority(self):
			return self.priority

		def set_priority(self, priority):
			self.priority = priority

		def get_based_on(self):
			return self.basedOn

		def set_based_on(self, based_on):
			self.basedOn = based_on

		def get_semi_hidden(self):
			return self.semiHidden

		def set_semi_hidden(self, semi_hidden):
			self.semiHidden = semi_hidden

		def get_unhide_when_used(self):
			return self.unhideWhenUsed

		def set_unhide_when_used(self, unhide_when_used):
			self.unhideWhenUsed = unhide_when_used

		def get_q_format(self):
			return self.qFormat

		def set_q_format(self, q_format):
			self.qFormat = q_format

		def get_rsid(self):
			return self.rsid

		def set_rsid(self, rsid):
			self.rsid = rsid

		def get_tab(self, number=0):
			return self.tab * (number + self.indent)

		def get_xml(self):
			temp = ''
			if self.get_default():
				temp = ' w:default="%s"' % self.get_default()
			elif self.get_custom_style():
				temp = ' w:customStyle="%s"' % self.get_custom_style()
			value = list()
			value.append('%s<%s w:type="%s"%s w:styleId="%s">' % (
				self.get_tab(), self.get_name(), self.get_type(), temp, self.get_style_id()
			))

			value.append('%s<w:name w:val="%s"/>' % (self.get_tab(1), self.get_value_name()))
			if self.get_based_on():
				value.append('%s<w:basedOn w:val="%s"/>' % (self.get_tab(1), self.get_based_on()))
			if self.get_priority():
				value.append('%s<w:uiPriority w:val="%s"/>' % (self.get_tab(1), self.get_priority()))
			if self.get_semi_hidden():
				value.append('%s<w:semiHidden/>' % self.get_tab(1))
			if self.get_unhide_when_used():
				value.append('%s<w:unhideWhenUsed/>' % self.get_tab(1))
			if self.get_link():
				value.append('%s<w:link w:val="%s"/>' % (self.get_tab(1), self.get_link()))
			if self.get_q_format():
				value.append('%s<w:qFormat/>' % self.get_tab(1))
			if self.get_rsid():
				value.append('%s<w:rsid w:val="%s"/>' % (self.get_tab(1), self.get_rsid()))

			for _element in self.elements:
				value.append(_element.get_xml())

			value.append('%s</w:%s>' % (self.get_tab(), self.get_name()))
			return self.separator.join(value)

	def __init__(self, parent):
		self.name = 'word/styles.xml'
		self.tag = 'styles'
		self.parent = parent
		self.separator = parent.separator
		self.tab = parent.tab
		self.indent = 0
		self.xml_header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if not self.separator:
			self.xml_header += '\n'
		self.xmlns = [
			'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
			'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
			'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
			'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
			'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
			'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"',
			'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"',
			'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"',
			'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"',
			'mc:Ignorable="w14 w15 w16se w16cid w16 w16cex"',
		]

		self.docDefaluts = self.DocDefaluts(self)
		self.latentStyles = self.LatentStyle(self)
		self.styles = list()
		self.styles.append(self.Style(self, 'paragraph', 'Normal', 'Normal', default='1', priority='', q_format=True))
		self.styles.append(
			self.Style(
				self, 'character', 'Fuentedeprrafopredeter', 'Default Paragraph Font', default='1',
				priority='1', semi_hidden=True, unhide_when_used=True))

		self.styles.append(self.Style(
			self, 'table', 'Tablanormal', 'No List', default='1', priority='99',
			semi_hidden=True, unhide_when_used=True))

		self.styles.append(self.Style(
			self, 'table', 'Descripcin', 'caption', default='', priority='99',
			semi_hidden=True, unhide_when_used=True))

		pr = table.TableProperties(self)
		pr.indent += 1
		pr.set_table_style('')
		pr.set_table_indentation("0")
		pr.set_cell_margin({
			'top': {'w': "0", 'type': "dxa"},
			'bottom': {'w': "0", 'type': "dxa"},
			'left': {'w': "108", 'type': "dxa"},
			'right': {'w': "108", 'type': "dxa"}
		})
		self.styles[-1].elements.append(pr)
		self.styles.append(self.Style(
			self, 'numbering', 'Sinlista', 'Normal', default='1', priority='99',
			semi_hidden=True, unhide_when_used=True))
		self.styles.append(self.Style(
			self, 'paragraph', 'Prrafodelista', 'List Paragraph', default='', priority='34',
			q_format=True, rsid='00BE31C9', based_on='Normal'))

		sparagraph, character = self.new_paragraph_character_style('Principal')

		pa = paragraph.Properties(sparagraph)
		pa.set_pstyle('')

		t = text.Properties(character)
		t.set_font_size(20)
		t.set_font('Arial')

		sparagraph.elements.append(pa)
		sparagraph.elements.append(t)
		character.elements.append(t)

	def new_paragraph_character_style(self, name, based_on='Normal', rsid='000F5F75'):
		link = name + 'Car'
		nlink = name + ' Car'
		_paragraph = self.Style(
			self, 'paragraph', name, name, default='', priority='', custom_style='1',
			q_format=True, rsid=rsid, based_on=based_on, link=link)
		self.styles.append(_paragraph)
		_character = self.Style(
			self, 'character', link, nlink, default='', priority='1', custom_style='1',
			rsid=rsid, based_on='Fuentedeprrafopredeter', link=name)
		self.styles.append(_character)
		return _paragraph, _character

	def get_separator(self):
		return self.separator

	def get_latentStyles(self):
		return self.latentStyles

	def set_latentStyles(self, _object):
		self.latentStyles = _object

	def get_doc_defaluts(self):
		return self.docDefaluts

	def set_doc_defaluts(self, doc_defaluts):
		self.docDefaluts = doc_defaluts

	def get_xmlHeader(self):
		return self.xml_header

	def SetXMLHeader(self, value):
		self.xml_header = value

	def get_parent(self):
		return self.parent

	def get_name(self):
		return self.name

	def GetTag(self):
		return self.tag

	def get_name(self):
		return self.name

	def get_tab(self, number=0):
		return self.tab * (number + self.indent)

	def get_xml(self):
		value = list()
		value.append(self.get_xmlHeader())
		xmlns = ' '.join(self.xmlns)
		if xmlns:
			xmlns = ' ' + xmlns
		value.append('%s<%s%s>' % (self.get_tab(), self.GetTag(), xmlns))
		value.append(self.get_doc_defaluts().get_xml())
		value.append(self.get_latentStyles().get_xml())

		for style in self.styles:
			value.append(style.get_xml())

		value.append('%s</%s>' % (self.get_tab(), self.GetTag()))

		return self.separator.join(value)
