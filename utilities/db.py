# coding=utf-8
import os
import re
from StringIO import StringIO
from cPickle import Pickler, Unpickler
from datetime import time
from time import *
from types import ListType, StringType
from zlib import *
from _bsddb import *

import rotor

_rt2 = rotor.newrotor('K;:-=)Tz', 17)
MX = chr(30)

cl = None
GS_INS = "D:/Server"
gpx = ('jonathan', 'sat', 'dgt')
apl = gpx[1]

def error(cl, txt):
	raise ValueError(txt)

def bkopen(_archivo, forma='c', tipo='b', borrar='', inx=(), cache=1, huf=0, tmk=0):
	fm = DB_CREATE
	if borrar == 's':
		fm = fm | DB_TRUNCATE
	if forma == 'r':
		fm = DB_RDONLY
	tp = DB_BTREE
	if tipo == 'h':
		tp = DB_HASH
	d = DBcp(_archivo, borrar, inx, cache, huf, tmk)
	d.open(_archivo, dbtype=tp, flags=fm | DB_THREAD)
	return d


class DBcp:
	__module__ = __name__

	def __init__(self, _archivo, borrar, inx, cache, huf, tmk):
		self.db = DB(None)
		# try:
		if tmk != 0:
			tama = tmk
		else:
			tama = os.path.getsize(_archivo)
		if tama < 20480:
			chd = 20480
		elif tama < 32768:
			chd = 24576
		elif tama < 65536:
			chd = 45056
		elif tama < 131072:
			chd = 102400
		elif tama < 262144:
			chd = 225280
		elif tama < 524288:
			chd = 0
		elif tama < 921600:
			chd = 307200
		elif tama < 2560000:
			chd = 358400
		elif tama < 5120000:
			chd = 409600
		else:
			chd = 460800
		if tama < 0:
			chd = -tama

		if chd > 0:
			self.db.set_cachesize(0, chd, 0)
		'''except:
			self.db.set_cachesize(0, 20480, 0)'''

		self.vx = {}
		self.tx = {}
		self.ch = cache
		self.inx = []
		self.huf = huf
		self.snc = None
		self._blq = ''
		if inx:
			for ncam, cp, col, tp in inx:
				ncam = ncam.lower()
				ls = [cp,
				      col,
				      '',
				      ncam,
				      tp]
				ls[2] = bkopen(_archivo + '.' + ncam, borrar=borrar, cache=cache, tmk=tmk)
				self.inx.append(ls)

		return

	def __del__(self):
		self.close()

	def __getattr__(self, name):
		return getattr(self.db, name)

	def __len__(self):
		if self.snc != '':
			return self.snc.__len__()
		return len(self.db)

	def __getitem__(self, idx):
		if self.snc:
			return self.snc.__getitem__(idx)
		if self.ch == 0:
			if self.huf == 0:
				return loads(self.db[idx])
			return loads(decompress(self.db[idx]))
		if idx in self.vx:
			return self.vx[idx]

		if self.huf == 0:
			dato = loads(self.db[idx])
		else:
			dato = loads(decompress(self.db[idx]))
		self.vx[idx] = dato
		return dato

	def __setitem__(self, idx, valor):
		if self.snc != '':
			return self.snc.__setitem__(idx, valor)
		valor_a = []
		if self.inx:
			if idx in self:
				valor_a = self[idx]

		if self.huf == 0:
			dato = dumps(valor, -1)
		else:
			dato = compress(dumps(valor, -1))

		self.db[idx] = dato
		if self.ch == 1:
			self.vx[idx] = valor

		if self.inx:
			for cps, cols, fi, ncam, tp in self.inx:
				if valor_a:
					for _k in range(len(cps)):
						cp = cps[_k]
						col = cols[_k]
						va = valor_a[cp]
						vn = valor[cp]
						if va == vn:
							continue
						if type(va) == list:
							for ln in va:
								if type(ln) == list:
									if Trae_Fila(vn, ln[col], clr=-2, clb=col) is not None:
										continue
									v = str(ln[col])
								else:
									if ln in vn:
										continue
									v = str(ln)
								if v == '':
									continue
								if tp == 'd':
									v = ajus0(v, 5)
									if v == '0None':
										v = '-9999'
								if tp == 'i':
									vls = Palabras_k(v)
									for v in vls:
										del fi[v + MX + idx]

								else:
									del fi[v + MX + idx]

						else:
							va = str(va)
							if va == '':
								continue
							if tp == 'd':
								va = ajus0(va, 5)
								if va == '0None':
									va = '-9999'
							if tp == 'i':
								vls = Palabras_k(va)
								for v in vls:
									del fi[v + MX + idx]

							else:
								del fi[va + MX + idx]

				for _k in range(len(cps)):
					cp = cps[_k]
					col = cols[_k]
					vn = valor[cp]
					va = ''
					if valor_a:
						va = valor_a[cp]
						if vn == va:
							continue
					if type(vn) == list:
						for ln in vn:
							if type(ln) == list:
								if va != '':
									if Trae_Fila(va, ln[col], clr=-2, clb=col) is not None:
										continue
								v = str(ln[col])
							else:
								if va != '':
									if ln in va:
										continue
								v = str(ln)
							if v == '':
								continue
							if tp == 'd':
								v = ajus0(v, 5)
								if v == '0None':
									v = '-9999'
							if tp == 'i':
								vls = Palabras_k(v)
								for v in vls:
									fi[v + MX + idx] = ''

							else:
								fi[v + MX + idx] = ''

					else:
						vn = str(vn)
						if vn == '':
							continue
						if tp == 'd':
							vn = ajus0(vn, 5)
							if vn == '0None':
								vn = '-9999'
						if tp == 'i':
							vls = Palabras_k(vn)
							for v in vls:
								fi[v + MX + idx] = ''

						else:
							fi[vn + MX + idx] = ''

		return

	def __delitem__(self, idx):
		if self.snc != '':
			return self.snc.__delitem__(idx)
		if self.inx:
			try:
				valor_a = self[idx]
				for cps, cols, fi, ncam, tp in self.inx:
					for _k in range(len(cps)):
						cp = cps[_k]
						col = cols[_k]
						va = valor_a[cp]
						if type(va) == list:
							for ln in va:
								if type(ln) == list:
									v = str(ln[col])
								else:
									v = str(ln)
								if v == '':
									continue
								if tp == 'd':
									v = ajus0(v, 5)
									if v == '0None':
										v = '-9999'
								if tp == 'i':
									vls = Palabras_k(v)
									for v in vls:
										del fi[v + MX + idx]

								else:
									del fi[v + MX + idx]

						else:
							va = str(va)
							if va == '':
								continue
							if tp == 'd':
								va = ajus0(va, 5)
								if va == '0None':
									va = '-9999'
							if tp == 'i':
								vls = Palabras_k(va)
								for v in vls:
									del fi[v + MX + idx]

							else:
								del fi[va + MX + idx]
			finally:
				pass

		if self.ch == 1:
			if idx in self.vx:
				del self.vx[idx]

		if idx in self.db:
			del self.db[idx]

	def cursor(self):
		if self.snc:
			return self.snc.cursor()
		return DBcpCursor(self.db.cursor(None, 0))

	def i_cursor(self, ncam):
		if self.snc:
			return self.snc.i_cursor(ncam)
		fi = Trae_Fila(self.inx, ncam.lower(), clb=3, clr=2)
		if fi is not None:
			return DBcpCursor(fi.db.cursor(None, 0))
		return

	def lee_b(self, idx):
		if self.snc:
			return self.snc.lee_b(idx)
		return self.db[idx]

	def graba_b(self, idx, valor):
		if self.snc:
			return self.snc.graba_b(idx, valor)
		self.db[idx] = valor

	# noinspection PyBroadException
	def i_selec(self, ncam, desde, hasta):
		if self.snc:
			return self.snc.i_selec(ncam, desde, hasta)
		tp = None
		for lnx in self.inx:
			if lnx[3] == ncam.lower():
				tp = lnx[4]
				break

		if tp is None:
			return
		res = []

		cursor = self.i_cursor(ncam)
		if cursor is None:
			return
		if tp == 'd':
			if desde == '' or desde is None:
				ds = -9999
			else:
				ds = int(desde)
			desde = ajus0(str(ds), 5)
			hasta = ajus0(str(hasta), 5)
		else:
			desde, hasta = str(desde), str(hasta)
		try:
			ids = cursor.set_range(desde, sl_cod=1)
			cds, idx = lista(ids, MX, 2)

			if desde <= cds <= hasta:
				res.append(idx)
				ers=0
			else:
				ers = 1
		except:
			ers = 1

		while ers == 0:
			try:
				ids = cursor.next(sl_cod=1)
				cds, idx = lista(ids, MX, 2)
				if desde <= cds <= hasta:
					res.append(idx)
				else:
					ers = 1
			except Exception as e:
				ers = 1

		del cursor

		return res

	def sincro(self):
		if self.snc != '':
			return self.snc.sincro()
		try:
			self.sync()
		finally:
			pass

		if self.inx:
			for ln in self.inx:
				try:
					ln[2].sync()
				finally:
					pass

	def cierra(self):
		del self.vx
		del self.tx
		self.close()
		if self.inx:
			for ln in self.inx:
				ln[2].close()


class DBcpCursor:
	__module__ = __name__

	def __init__(self, cursor):
		self.dbc = cursor

	def __del__(self):
		self.close()

	def __getattr__(self, nombre):
		return getattr(self.dbc, nombre)

	def dup(self, flags=0):
		return DBcpCursor(self.dbc.dup(flags))

	def get_1(self, flags, sl_cod=0):
		rec = self.dbc.get(flags)
		return self._extrae(rec, sl_cod)

	def current(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_CURRENT, sl_cod)

	def first(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_FIRST, sl_cod)

	def last(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_LAST, sl_cod)

	def next(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_NEXT, sl_cod)

	def prev(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_PREV, sl_cod)

	def consume(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_CONSUME, sl_cod)

	def next_dup(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_NEXT_DUP, sl_cod)

	def next_nodup(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_NEXT_NODUP, sl_cod)

	def prev_nodup(self, flags=0, sl_cod=0):
		return self.get_1(flags | DB_PREV_NODUP, sl_cod)

	def set(self, idx, flags=0, sl_cod=0):
		rec = self.dbc.set(idx, flags)
		return self._extrae(rec, sl_cod)

	def set_range(self, idx, flags=0, sl_cod=0):
		rec = self.dbc.set_range(idx, flags)
		return self._extrae(rec, sl_cod)


	def _extrae(rec, sl_cod):
		if rec is None:
			raise ValueError('No existe clave')
		else:
			if sl_cod == 1:
				return rec[0]
			return rec[0], loads(rec[1])



'''FVT = {apl: bkopen(GS_INS + '/Applications/' + apl + '/gsPgs', huf=1)}
DCN = {apl: list()}
FDC = {apl: bkopen(GS_INS + '/Applications/' + apl + '/Dcts', huf=1)}
FIM = {apl: bkopen(GS_INS + '/Applications/' + apl + '/gsImp', huf=1)}
ficheros = FDC[apl].keys()
DCN[apl] = ficheros

F = dict()
for archivo in FDC[gpx[1]].keys():
	ruta = GS_INS + "/Companies/jonathan/%s/%s/%s" % (gpx[1], gpx[2], archivo)
	k = ((gpx[0], gpx[1], gpx[2]), archivo)
	F[k] = bkopen(ruta, huf=0)'''
F, FDC, FAP, FVT, FIM, INF_U, ERS, FUS, BLQ_AP, FGAP = ({}, {}, {}, {}, {}, {}, {}, {}, {},
 None)
DCN, DCI = ({}, {})
HSI, MSP = ({}, {})
DCN[(0, 'emg')] = (
 'EMG_NM',
 1,
 1,
 '',
 0, (), ())
DCI[(0, 'emg')] = (
 'l',
 0,
 '',
 '')
DCN[(0, 'apg')] = (
 (
  'APG_NM',
  'APG_TP',
  'APG_VER'),
 3,
 3,
 '',
 0, (), ())
DCI[(0, 'apg')] = (
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''))
DCN[(0, 'apl')] = (
 (
  'APL_NM',
  'APL_EJER'),
 2,
 2,
 '',
 0, (), ())
DCI[(0, 'apl')] = (
 (
  'l',
  0,
  '',
  ''),
 (
  (
   'l',
   'l',
   'l',
   'l'),
  4,
  '',
  ''))
DCN[(0, 'dcs')] = (
 (
  'DC_NMA',
  'DC_TPO',
  'DC_CPS',
  'DC_INX',
  'DC_LCOD',
  'DC_ACTU',
  'DC_DPL',
  'DC_MID',
  'DC_GRABA',
  'DC_BORRA',
  'DC_HISTO',
  'DC_DGRABA',
  'DC_CACHE'),
 13,
 13,
 '',
 0, (), ())
DCI[(0, 'dcs')] = (
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  (
   'l',
   'l',
   'l',
   'i',
   'l',
   'l'),
  6,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'i',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''),
 (
  'l',
  0,
  '',
  ''))
DCN[('*', 'mensajes')] = [
 [
  'MJ_DE',
  'MJ_PA',
  'MJ_FEC',
  'MJ_HR',
  'MJ_ASU',
  'MJ_ES',
  'MJ_FE',
  'MJ_HE',
  'MJ_AC'],
 9,
 9,
 '',
 6, [],
 '',
 'MJ_',
 '',
 '',
 '',
 '',
 '',
 's']
DCI[('*', 'mensajes')] = [
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'd',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  [
   'l'],
  1,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'd',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'i',
  0,
  '',
  '',
  '']]
DCN[('*', 'tareas_ps')] = [[],
 11,
 11,
 '',
 3, [],
 '',
 'TP_',
 '',
 '',
 '',
 '',
 '',
 's']
DCN[('*', 'plgsb')] = [[],
 5,
 5,
 '',
 4, [],
 '',
 'PL_',
 '',
 '',
 '',
 '',
 [
  ''],
 '',
 's']
DCI[('*', 'plgsb')] = [
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'd',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  '']]
DCN[('*', 'psgsb')] = [[],
 13,
 13,
 '',
 4, [],
 '',
 'PS_',
 '',
 '',
 '',
 '',
 [
  ''],
 '',
 's']
DCI[('*', 'psgsb')] = [
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  [
   'l'],
  1,
  '',
  '',
  ''],
 [
  'l',
  0,
  '',
  '',
  ''],
 [
  'i',
  0,
  '',
  '',
  '']]
NL = '\x01\x02'
RAIZ, PRG, TRS, MVEN = ({}, {}, {}, {})
RAIZ[('MJ_', '*')] = [
 '*',
 'mensajes']
RAIZ[0] = ''

def Leex(fi, idx):
	try:
		return fi[idx]
	except:
		return 1


FGAP = bkopen(GS_INS + '/Applications/Apps', huf=1)
def Abre_Aplicacion(apl, FGAP=FGAP, er_fm=1):
	if DCN.has_key(apl):
		return (0, '')
	rg = Leex(FGAP, apl)
	if rg == 1:
		return (1,
		'No existe Aplicacion' + ' ' + apl)
	nmaplg, noeje, ver = rg
	FVT[apl] = bkopen(GS_INS + '/Applications/' + apl + '/gsPgs', huf=1)
	if noeje == 's':
		DCN[apl] = []
		return (0, 'ap_sd')
	FDC[apl] = bkopen(GS_INS + '/Applications/' + apl + '/Dcts', huf=1)
	FIM[apl] = bkopen(GS_INS + '/Applications/' + apl + '/gsImp', huf=1)
	ficheros = FDC[apl].keys()
	DCN[apl] = ficheros
	for fi in ficheros:
		rg = Leex(FDC[apl], fi)
		nmfi, tipo, campos, inx, lcod, actu, a_dpl, a_mid, a_graba, a_borra, a_histo, a_dp_grabar, cache = rg
		lt, lr = (0, 0)
		cpx, fmx = [], []
		inx = inx.split()
		a_mid = a_mid.split()
		actu = actu.split()
		a_histo = lista(a_histo, ',', 2)
		raiz = ''
		if len(campos) > 0:
			c1 = campos[0][0]
			q = c1.find('_')
			if q > 0:
				raiz = c1[:q + 1]
			RAIZ[(raiz, apl)] = (apl, fi)
		for cps in campos:
			ncam, dscam, fmt, ncols, rel, formu = cps
			malo = 0
			if ncols > 0:
				if rel != '':
					if rel.find('/') < 0:
						malo = 2
				fmt = fmt.split()
				if len(fmt) != ncols:
					malo = 1
				for f in fmt:
					if len(f) != 1:
						malo = 1
					elif f not in '012345li%dh':
						malo = 1

				dscam = lista(dscam, '|', ncols)
			else:
				if rel != '':
					q = rel.find('/')
					if q == 1:
						malo = 2
				if len(fmt) != 1:
					malo = 1
				else:
					if fmt not in '012345li%dh':
						malo = 1
			if malo > 0 and er_fm == 1:
				FDC[apl].cierra()
				del FDC[apl]
				del DCN[apl]
				if malo == 1:
					return (1,
					 'Formato (' + str(fmt) + ') para' + ' ' + ncam + ': ' + fi + ' > ' + apl + ' ' + ' ' + 'erroneo')
				return (
				 1,
				 '!Error en relacion' + ' ' + ncam + ': ' + fi)
			cpx.append(ncam)
			fmx.append([fmt,
			 ncols,
			 rel,
			 formu,
			 formu,
			 dscam])
			if formu == '':
				lr = lr + 1
			lt = lt + 1

		if cache.lower() != 's':
			cache = 0
		else:
			cache = 1
		inz = []
		for ind in inx:
			cps, nm = lista(ind, '=', 2)
			if nm == '':
				nm = cps
			nm, tpo = lista(nm, '/', 2)
			cps = cps.split('+')
			cpi, cols = [], []
			for cp in cps:
				cpr, col = lista(cp, '/', 2, '1')
				try:
					nc = cpx.index(cpr)
				except:
					nc = -1

				if nc == -1:
					continue
				cpi.append(nc)
				cols.append(col)
				if tpo == '':
					try:
						fmt = fmx[nc][0]
						if type(fmt) == ListType:
							fm = fmt[col]
						else:
							fm = fmt
						if fm == 'i' or fm == 'd':
							tpo = 'd'
					except:
						pass

			if cpi != []:
				inz.append([nm,
				 cpi,
				 cols,
				 tpo])

		DCN[(apl, fi)] = [cpx,
		 lr,
		 lt,
		 tipo,
		 lcod,
		 inz,
		 actu,
		 a_dpl,
		 raiz,
		 a_mid,
		 a_graba,
		 a_borra,
		 a_histo,
		 a_dp_grabar,
		 cache]
		DCI[(apl, fi)] = fmx

	# Traduce_Formulas(apl)
	return (0, '')


def abre_flx(emges, apl, emp, archi, inx, cache, borrar='', sincro=1):
	raiz = GS_INS + '/Companies/' + emges + '/' + apl + '/' + emp + '/'
	F[(emges, apl, emp, archi)] = bkopen(raiz + archi, borrar=borrar, inx=inx, cache=cache)

	rg = Leex(FAP[emges], apl)
	if rg == 1:
		return 'Imposible abrir archivo de Aplicaciones' + ' ' + str((emges,
		 apl))
	if sincro == 0:
		return 0
	sincro = rg[2]
	for ejo, ars, ejd in sincro:
		if ejd == emp:
			continue
		if ejo == '*':
			ejo = emp
		ars = ars.split()
		ejo = ejo.split()
		if emp in ejo and archi in ars:
			if not F.has_key((emges,
			 apl,
			 ejd,
			 archi)):
				ok = Abre_Empresa(emges, apl, ejd, no_dc=1)
				if ok == 1:
					return 'Imposible abrir archivo sincronizado' + ' ' + str((emges,
					 apl,
					 ejd,
					 archi))
			F[(emges, apl, emp, archi)].snc = F[(emges,
			 apl,
			 ejd,
			 archi)]
			break

	return 0

def Abre_Empresa(emges, apl, emp, no_dc=0, clx=None, ars=[], sincro=1):
	emges = emges.lower()
	apl = apl.lower()
	emp = emp.lower()
	dcts = []
	if ars != []:
		ficheros = ars
	else:
		try:
			ficheros = DCN[apl]
		except:
			return 1

	if not INF_U.has_key((emges,
	 apl)) and ficheros != []:
		INF_U[(emges, apl)] = bkopen(GS_INS + '/Companies/' + emges + '/' + apl + '/info_usu')
	if emp == '':
		return 1

	FAP[emges] = bkopen(GS_INS + '/Companies/' + emges + '/apls')
	FUS[emges] = bkopen(GS_INS + '/Companies/' + emges + '/usGsb')
	F[(emges, '*', '*', 'plgsb')] = bkopen(GS_INS + '/Companies/' + emges + '/plGsb')
	F[(emges, '*', '*', 'psgsb')] = bkopen(GS_INS + '/Companies/' + emges + '/psGsb')

	for fi in ficheros:
		tipo, inx, cache, lcod = (
		 DCN[(apl, fi)][3],
		 DCN[(apl, fi)][5],
		 DCN[(apl, fi)][14],
		 DCN[(apl, fi)][4])
		if not F.has_key((emges, apl, emp, fi)):

			ok = abre_flx(emges, apl, emp, fi, inx, cache, sincro=sincro)
			if ok != 0:
				if clx is not None:
					a = 1 / 0
				return 1


		if no_dc == 0:
			dccl = []
			raiz = DCN[(apl,
			 fi)][8]
			lr = DCN[(apl,
			 fi)][1]
			for nc in range(lr):
				cp = DCN[(apl,
				 fi)][0][nc]
				fmt, ncols, rel = DCI[(apl,
				 fi)][nc][:3]
				dscam = DCI[(apl,
				 fi)][nc][5]
				dccl.append((cp,
				 fmt,
				 ncols,
				 rel,
				 dscam))

			dcts.append((fi,
			 dccl,
			 raiz))

	return dcts


MESES = [
	'Ene',
	'Feb',
	'Mar',
	'Abr',
	'May',
	'Jun',
	'Jul',
	'Ago',
	'Set',
	'Oct',
	'Nov',
	'Dic']
MESES_C = [
	'Enero',
	'Febrero',
	'Marzo',
	'Abril',
	'Mayo',
	'Junio',
	'Julio',
	'Agosto',
	'Septiembre',
	'Octubre',
	'Noviembre',
	'Diciembre']


def Pr_char(n):
	i = ord(n)
	i = i + 1
	n = chr(i)
	if n > '9' and n < 'A':
		n = 'A'
	if n > 'Z':
		n = '0'
	return n


def Busca_Prox(n, lfija='n'):
	l = len(n)
	if n == '':
		return '0'
	if n.isdigit():
		if n == '9' * l:
			return '9' * (l - 1) + 'A'
		n = long(n)
		n = n + 1
		return ajus0(str(n), l)
	for k in range(l - 1, -1, -1):
		if not n[k].isdigit():
			break

	letra, num = n[:k + 1], n[k + 1:]
	if num != '':
		num = Busca_Prox(num)
		return letra + num
	if n == 'Z' * l:
		if lfija == 'n':
			return n + '0'
	if not n.isalnum() or n == 'Z' * l or not n.isupper():
		return 0
	i = l - 1
	q = Pr_char(n[i])
	while q == '0' and i >= 0:
		n = n[:i] + q + n[i + 1:]
		i = i - 1
		q = Pr_char(n[i])

	n = n[:i] + q + n[i + 1:]
	return n


def u_libre(gpx, fi, serie=''):
	emges, apl, emp = gpx
	lcod = DCN[(apl, fi)][4]
	fz = F[(emges, apl, emp, fi)]
	cursor = fz.cursor()
	if serie == '':
		try:
			urg = cursor.last(sl_cod=1)
		except:
			urg = ''

		idu = urg
		if lcod == 0:
			idz = Busca_Prox(idu)
		else:
			if idu == '':
				idz = '0' * lcod
			else:
				idz = Busca_Prox(idu, lfija='s')
				if idz != 0:
					if len(idz) != lcod:
						idz = 0
		idx = idz
	else:
		if lcod == 0:
			serie, lcod = lista(serie, ',', 2, '1')
			if lcod == 0:
				lcod = 8
		ls = len(serie)
		lnum = lcod - ls
		try:
			ids = cursor.set_range(serie, sl_cod=1)
			ida = ids
			if len(ids) != lcod:
				ids = ''
			if ids[:ls] == serie:
				serie_p = Busca_Prox(serie)
				if serie_p != 0:
					try:
						ids = cursor.set_range(serie_p, sl_cod=1)
						ids = cursor.prev(sl_cod=1)
					except:
						try:
							ids = cursor.last(sl_cod=1)
						except:
							ids = ''

					if ids[:ls] == serie:
						numero = ids[ls:]
						if lcod == 0:
							numero = Busca_Prox(numero)
						else:
							numero = Busca_Prox(numero, lfija='s')
						if numero != 0:
							idz = serie + numero
							if len(idz) != lcod:
								idz = 0
						else:
							idz = 0
						if idz != 0:
							if idz[:ls] != serie:
								idz = 0
					else:
						numero = ida[ls:]
						if lcod == 0:
							numero = Busca_Prox(numero)
						else:
							numero = Busca_Prox(numero, lfija='s')
						if numero != 0:
							idz = serie + numero
							if len(idz) != lcod:
								idz = 0
						else:
							idz = 0
				else:
					idz = 0
			else:
				if lcod == 0:
					idz = serie + '0'
				else:
					idz = serie + ajus0('0', lnum)
					if len(idz) != lcod:
						idz = 0
		except:
			if lcod == 0:
				idz = serie + '0'
			else:
				idz = serie + ajus0('0', lnum)
				if len(idz) != lcod:
					idz = 0

		idx = idz
	del cursor
	return idx


def lista(cad, sep, _min=0, nms='', mx=''):
	ax = cad.split(sep)
	if _min > 0:
		while len(ax) < _min:
			ax.append('')

	if nms != '':
		_l = len(ax)
		if nms == '*':
			bx, _k = [], 0
			while len(bx) < _l:
				bx.append(str(_k))
				_k = _k + 1

		else:
			bx = nms.split('-')
		for _k in bx:
			entero = 1
			if _k.find('.') >= 0:
				entero = 0
				_k = _k.replace('.', '')
			_k = int(_k)
			if ax[_k] == '':
				ax[_k] = '0'
			if entero == 0:
				ax[_k] = 0.0
				try:
					ax[_k] = float(ax[_k])
				finally:
					pass

			else:
				ax[_k] = 0
				try:
					ax[_k] = int(ax[_k])
				finally:
					pass

	if mx != '':
		return ax[:mx]
	return ax


def lista(cad, sep, min=0, nms='', mx=''):
	ax = cad.split(sep)
	if min > 0:
		while len(ax) < min:
			ax.append('')

	if nms != '':
		l = len(ax)
		if nms == '*':
			bx, k = [], 0
			while len(bx) < l:
				bx.append(str(k))
				k = k + 1

		else:
			bx = nms.split('-')
		for k in bx:
			entero = 1
			if k.find('.') >= 0:
				entero = 0
				k = k.replace('.', '')
			k = int(k)
			if ax[k] == '':
				ax[k] = '0'
			if entero == 0:
				try:
					ax[k] = float(ax[k])
				except:
					ax[k] = 0.0

			else:
				try:
					ax[k] = int(ax[k])
				except:
					ax[k] = 0

	if mx != '':
		return ax[:mx]
	return ax


# noinspection PyPep8Naming
def Trae_Fila(tabla, valor, clr=0, clb=0, match_ini=''):
	lmax = len(tabla)
	vc = valor
	for _k in range(lmax):
		if match_ini != '':
			ldt = len(tabla[_k][clb])
			vc = valor[:ldt]
		if tabla[_k][clb] == vc:
			if clr == -1:
				return tabla[_k]
			if clr == -2:
				return _k
			return tabla[_k][clr]

	return


def ajus0(s, longitud):
	res = s.rjust(longitud).replace(' ', '0')
	la = len(res)
	if la > longitud:
		return res[la - longitud:]
	return res


# noinspection PyPep8Naming
def Palabras_k(v):
	v = v.replace('.', ' ')
	v = v.replace(',', ' ')
	vls = v.split()
	res = []
	for v in vls:
		if len(v) < 2:
			continue
		res.append(v.upper())

	return res


def dump(obj, _file, protocol=None):
	Pickler(_file, protocol).dump(obj)


def dumps(obj, protocol=-1):
	_file = StringIO()
	Pickler(_file, protocol).dump(obj)
	return _file.getvalue()


def load(_file):
	return Unpickler(_file).load()


def loads(_str):
	_file = StringIO(_str)
	return Unpickler(_file).load()


def lee_dc(lee_dc, _gpx, fi, idx, mode='value', rels='s', respu='n'):
	if type(idx) == unicode:
		idx = idx.encode('latin-1')

	rg = {}

	if type(_gpx[2]) == unicode:
		_gpx = (_gpx[0], _gpx[1], _gpx[2].encode('latin-1'))
	_k = ((_gpx[0], _gpx[1], _gpx[2]), fi)


	try:aux = lee(None, gpx, fi, idx)
	except:
		raise ValueError(_gpx, fi)
	campos = FDC[_gpx[1]][fi][2]
	if aux == 1:
		rg = aux
		if respu == 'n':
			rg = {}
			for _k in range(len(campos)):
				name, deno, fmt, cols, relations, calculate = eval(repr(campos[_k][:6]))
				val = ''
				if cols:
					val=[]
				elif fmt in ['12345']:
					val=0.
				elif fmt == 'i':
					val=0
				elif fmt=='d':
					val = None
				rg[name] = val

		return rg

	for _k in range(len(campos)):
		name, deno, fmt, cols, relations, calculate = eval(repr(campos[_k][:6]))
		if calculate:
			continue
		if type(name) == unicode:
			name = name.encode('latin-1')
		if type(aux[_k]) == unicode:
			aux[_k] = aux[_k].encode('latin-1')
		if not mode or mode == 'value':
			rg[name] = aux[_k]
		elif mode == 'deno':
			rg[name] = deno
		elif mode == 'fmt':
			if cols:
				fmt = fmt.split(' ')
			rg[name] = fmt
		elif mode == 'rels':
			if cols:
				_relations = relations.split('|')
				relations = list()
				for ln in _relations:
					relations.append(ln.split('/')[:2])
			rg[name] = relations
		if rels == 's':
			if not cols and relations:
				if '/' in relations:
					relations = relations.split('/')[0]
				rg[name] = lee_dc(lee_dc, gpx, relations, aux[_k], mode, rels='')
				rg[name]['IDX'] = aux[_k]

	return rg


def lee(cl, gpx, fi, idx, ix_s=''):
	emges, apl, emp = gpx
	fz = F[(emges,
	 apl,
	 emp,
	 fi)]
	tr_cl = TRS.get(cl, [])
	rg_t = 0
	if fz in tr_cl:
		rg_t = fz.tx.get(idx, 1)
		if rg_t != 1:
			return rg_t
		if fz.ch == 0:
			rg_t = 0
	rs = _Lee(fz, idx)
	if rs != 1:
		if rg_t == 0:
			return rs
		return copia_rg(rs)
	if ix_s != '':
		for cp_ix in ix_s:
			idxs = fz.i_selec(cp_ix, idx, idx)
			if idxs is None:
				error(cl, 'Imposible leer por indice' + ' ' + cp_ix + '\ncompacte archivo' + ' ' + fi + ' ' + 'o vea si est\xe1 estropeado indice')
			if idxs:
				idx = idxs[0]
				rs = _Lee(fz, idx)
				break

	if rg_t == 0 or rs == 1:
		return rs
	return copia_rg(rs)



def Veri_Fecha(d, m, y):
	malo = 1
	if m < 1 or m > 12 or d > 31 or d < 1:
		malo = 0
	if m in (4, 6, 9, 11) and d > 30:
		malo = 0
	if y < 1 or y > 9999:
		malo = 0
	if m == 2:
		if (y / 4.0 != int(y / 4.0) or y / 100.0 == int(y / 100.0) and y / 400.0 != int(y / 400.0)) and d > 28:
			malo = 0
		if d > 29:
			malo = 0
	return malo


def Fecha_aNum(v):
	if v == '':
		return
	if len(v) < 8:
		d, m, y = v[:2], v[2:4], v[4:]
		dh, mh, yh = Fecha('dmy')
		d = int(d)
		if m != '':
			m = int(m)
		else:
			m = mh
		if y != '':
			y = int(y)
		else:
			y = yh
	else:
		try:
			d, m, y = int(v[:2]), int(v[2:4]), int(v[4:])
		except:
			return

	if Veri_Fecha(d, m, y):
		return d - 1 + (m - 1) * 31 + (y - 2000) * 372
	return


def Fecha(fmt='', t='', fmti=0):
	if t == '':
		t = localtime(time())
	else:
		if t > 2147483600.0:
			t = 2147483600.0
		t = localtime(t)
	if fmt == 't':
		return time()
	d, m, y, hora, minu, sec, dia_s = (
		t[2],
		t[1],
		t[0],
		t[3],
		t[4],
		t[5],
		t[6])
	if fmt == 'd':
		return d
	if fmt == 'ds':
		return ajus0(str(d), 2)
	if fmt == 'm':
		return m
	if fmt == 'ms':
		return ajus0(str(m), 2)
	if fmt == 'a':
		return y
	if fmt == 'as':
		return ajus0(str(y), 4)
	if fmt == 'hm':
		return ajus0(str(hora), 2) + ':' + ajus0(str(minu), 2)
	if fmt == 'hms':
		return ajus0(str(hora), 2) + ':' + ajus0(str(minu), 2) + ':' + ajus0(str(sec), 2)
	if fmt in ('X', 'x'):
		da = ajus0(str(d), 2)
		ma = ajus0(str(m), 2)
		ya = ajus0(str(y), 4)
		sep = '/'
		if fmt == 'x':
			sep = ''
		if fmti == 0:
			return da + sep + ma + sep + ya
		if fmti == 1:
			return ma + sep + da + sep + ya
		return ya + sep + ma + sep + da
	else:
		if fmt == 'dmy':
			return (d,
			        m,
			        y)
		if fmt in ('1e', 'fa', '1m'):
			ya = ajus0(str(y), 4)
			da = '01'
			if fmt == 'fa':
				da = '31'
			ma = ajus0(str(m), 2)
			if fmt == '1e':
				ma = '01'
			if fmt == 'fa':
				ma = '12'
			return Fecha_aNum(da + ma + ya)
		if fmt == 'dsm':
			return dia_s
	if fmt == 'L':
		da = str(d)
		ma = MESES_C[m - 1]
		ya = str(y)
		sep = ', '
		if fmti == 0:
			return da + sep + ma + sep + ya
		if fmti == 1:
			return ma + sep + da + sep + ya
		return ya + sep + ma + sep + da
	da = ajus0(str(d), 2)
	ma = ajus0(str(m), 2)
	ya = ajus0(str(y), 4)
	return Fecha_aNum(da + ma + ya)


def Num_aFecha(n, fmt='', fmti=0):
	if (not n) == 1 and n != 0:
		return ''
	n = int(n)
	y = int(n / 372)
	v = n - 372 * y
	m = int(v / 31)
	d = v - 31 * m + 1
	m = m + 1
	y = y + 2000
	if Veri_Fecha(d, m, y) == 0:
		return
	if fmt in ('', 'x', 'x2'):
		da = ajus0(str(d), 2)
		ma = ajus0(str(m), 2)
		ya = ajus0(str(y), 4)
		sep = '/'
		if fmt in ('x', 'x2'):
			sep = ''
		if fmt == 'x2':
			ya = ya[2:]
		if fmti == 0:
			return da + sep + ma + sep + ya
		if fmti == 1:
			return ma + sep + da + sep + ya
		return ya + sep + ma + sep + da
	else:
		if fmt == 'L':
			da = str(d)
			ma = MESES_C[m - 1]
			ya = str(y)
			sep = ', '
			if fmti == 0:
				return da + sep + ma + sep + ya
			if fmti == 1:
				return ma + sep + da + sep + ya
			return ya + sep + ma + sep + da
		else:
			if fmt == 'd':
				return d
			if fmt == 'ds':
				return ajus0(str(d), 2)
			if fmt == 'm':
				return m
			if fmt == 'm-':
				return m - 1
			if fmt == 'ms':
				return ajus0(str(m), 2)
			if fmt == 'dmy':
				return (d,
				        m,
				        y)
			if fmt == 'a':
				return y
			if fmt == 'as':
				return ajus0(str(y), 4)
			if fmt == 'dsm':
				f1 = mktime((y,
				             m,
				             d,
				             2,
				             0,
				             0,
				             0,
				             0,
				             -1))
				return Fecha('dsm', f1)
	return


def rg_vacio(_gpx, fi, gpx=gpx):
	return lee_dc(None, _gpx, fi, '', mode='value', rels='', respu='n')


# noinspection PyPep8Naming
def Cb_Grupo(m):
	v = m.group()
	if v == '<>':
		return v
	if ' ' in v:
		return v
	v = v[1:len(v) - 1]
	if v[:1] == ':':
		return 'pb._vps["' + v[1:] + '"]'
	if ',' in v:
		v = v.split(',')
		rs = 'Vx'
		for k in v:
			if k != '':
				rs = rs + '[' + k + ']'

		return rs
	return 'ct["' + v + '"]'


# noinspection PyPep8Naming
def Traduce(prg):
	prex = re.compile('<.*?>')
	prg = prex.sub(Cb_Grupo, prg)
	prg = prg.replace(chr(171), '<')
	prg = prg.replace(chr(187), '>')
	return prg


# return compile(prg, '', 'exec')

def Int(value):
	try:
		return int(value)
	except:
		return 0


def Num(value):
	try:
		return float(value)
	except:
		return 0.


def Busca_Error(cl):
	raise ValueError('')


def copia_rg(obj):
	return eval(repr(obj))


def i_selec(cl, gpx, fi, ixs, desde, hasta):
	res = F[(gpx[0],
	         gpx[1],
	         gpx[2],
	         fi)].i_selec(ixs, desde, hasta)
	if res == None:
		error(cl, 'Hay algun error al seleccionar por' + ' ' + fi + ' > ' + ixs)
	return res


def Campo_a_Num(kdc, cp):
	if cp == '?':
		return cp
	try:
		return DCN[kdc][0].index(cp)
	except:
		return -1


def cons_preguntas(emges, apl, emp, archi, preguntas, bus_inx=''):
	inz = DCN[(apl,
	           archi)][5]
	inx = []
	for lnx in inz:
		inx.append(lnx[0])

	preg = []
	preg_cd = []
	busca_inx = bus_inx
	preg_bi = {}
	inx_dh = []
	for cp, ope, v in preguntas:
		colu = ''
		if cp.find('/') >= 0:
			cp, colu = lista(cp, '/', 2)
		if cp == '0':
			preg.append((-1,
			             ope,
			             v,
			             '',
			             0))
			preg_cd.append((ope,
			                v))
		else:
			nc = Campo_a_Num((apl,
			                  archi), cp)
			if nc == -1:
				continue
			fmt, formu = DCI[(apl,
			                  archi)][nc][0], DCI[(apl,
			                                       archi)][nc][4]
			if colu != '':
				if colu[:1] == 'f':
					fmt = fmt[0]
					colu = Int('-' + colu[1:])
					colu = colu - 1
				else:
					colu = Int(colu)
					fmt = fmt[colu]
			else:
				if type(fmt) == ListType:
					colu = 0
					fmt = fmt[0]
			num = 0
			if fmt in '012345':
				v, num = Num(v), 1
			else:
				if fmt in 'id':
					v, num = Int(v), 1
			if busca_inx == '':
				if formu[:5] == 'Leev(' and fmt == 'l':
					arg = formu[6:len(formu) - 2]
					if ope == '=' and v.find('[') < 0 and v.find('^') < 0:
						for p in ('<', '>', '&', '%', 'rg'):
							arg = arg.replace(p, '')

						fi, cpa, cpr = lista(arg, ',', 3)
						if cpa in inx:
							inxs = Trae_Fila(DCN[(apl,
							                      fi)][5], 'inv', clr=-1)
							if inxs != None:
								ncr = Campo_a_Num((apl,
								                   fi), cpr)
								if ncr in inxs[1]:
									vx = v.replace(']', '')
									vx = vx.replace(' ', '?')
									vx = vx.replace('&', '?')
									vx = vx.replace('.', '')
									vx = vx.replace(',', '')
									vx = vx.upper()
									vx = vx.split('?')
									if v.find(' ') >= 0:
										vx = vx[:1]
									fz0 = F[(emges,
									         apl,
									         emp,
									         fi)]
									vt = []
									for p in vx:
										try:
											p = fz0.i_selec('inv', p, p + 'z')
											if p != []:
												vt.extend(p)
										except:
											pass

									if vt != []:
										busca_inx = (
											cpa,
											vt)
										if v.find(' ') >= 0:
											preg.append((nc,
											             ope,
											             v,
											             colu,
											             num))
				vz = v
				if num == 1:
					v = str(v)
				if fmt == 'd':
					v = ajus0(v, 5)
				if cp in inx and (colu == '' or colu == 0):
					if ope == '=' and v.find('[') < 0 and v.find('&') < 0 and v.find('^') < 0:
						vx = v.replace(']', '')
						vx = vx.split('?')
						busca_inx = (cp,
						             vx)
					elif ope in ('>=', '<=', '>', '<') and v.find('[') < 0 and v.find('&') < 0 and v.find('^') < 0:
						if inx_dh == []:
							inx_dh = [
								cp,
								'',
								'']
						if cp == inx_dh[0]:
							if ope == '>=' or ope == '>':
								inx_dh[1] = v
							if ope == '<=' or ope == '<':
								inx_dh[2] = v
				if 'inv' in inx:
					inxs = Trae_Fila(inz, 'inv', clr=-1)
					nc_campos_inv = inxs[1]
					if nc in nc_campos_inv:
						if ope == '=' and v.find('[') < 0 and v.find('^') < 0:
							vx = v.replace(']', '')
							vx = vx.replace(' ', '?')
							vx = vx.replace('&', '?')
							vx = vx.replace('.', '')
							vx = vx.replace(',', '')
							vx = vx.upper()
							vx = vx.split('?')
							if v.find(' ') >= 0:
								vx = vx[:1]
							busca_inx = (
								'inv',
								vx)
							if v.find(' ') >= 0:
								preg.append((nc,
								             ope,
								             vz,
								             colu,
								             num))
				if busca_inx == '' or busca_inx != '' and v.find('&') >= 0:
					preg.append((nc,
					             ope,
					             vz,
					             colu,
					             num))
			else:
				preg.append((nc,
				             ope,
				             v,
				             colu,
				             num))
		for nc, ope, v, colu, num in preg:
			if colu != '' and colu >= 0:
				prb = preg_bi.get(nc, [])
				prb.append([ope,
				            v,
				            colu,
				            num])
				preg_bi[nc] = prb

	return (inx,
	        preg,
	        preg_cd,
	        busca_inx,
	        preg_bi,
	        inx_dh)


def Tradu_Campos(kdc, campos, noexis=0):
	res = []
	for ncam in campos:
		if ncam == '0':
			res.append(-1)
		else:
			nc = Campo_a_Num(kdc, ncam)
			if nc != -1:
				res.append(nc)
			elif noexis == 1:
				res.append(-2)

	return res


def _Lee(fi, idx):
	try:
		return fi[idx]
	except:
		return 1


def Rever_Cd(cd, l=0):
	cd = cd.upper()
	if l != 0:
		cd = ajus0(cd, l)
	res = ''
	for k in cd:
		res = res + chr(256 - ord(k))

	return res

def Nulo(fmt):
	v = ''
	if fmt in '012345':
		v = 0.0
	else:
		if fmt == 'i':
			v = 0
		else:
			if fmt == 'd':
				v = None
	return v


def Calcu_Campo(emges, apl, emp, fi, idx, rg, nc):
	formu = DCI[(apl,
	 fi)][nc][3]
	try:
		return eval(formu)
	except:
		fmt = DCI[(apl,
		 fi)][nc][0]
		return Nulo(fmt)

def concuerda(dato, fmt):
	if fmt.find('?') >= 0:
		fmt = fmt.split('?')
	else:
		fmt = [
		 fmt]
	for v in fmt:
		l = len(v)
		if l == 0:
			continue
		fm = ''
		if v.find('^') >= 0:
			for k in range(l):
				if v[k] == '^':
					ca = dato[k:k + 1]
					if ca == '':
						ca = ' '
					fm = fm + ca
				else:
					fm = fm + v[k]

			v = fm
		l = len(v)
		ini, fin = v[:1], v[l - 1]
		if ini == '[' and fin == ']':
			if dato.find(v[1:l - 1]) >= 0:
				return 1
		if fin == ']':
			if dato[:l - 1] == v[:l - 1]:
				return 1
		elif ini == '[':
			v, l = v[1:], l - 1
			lx = len(dato)
			if dato[lx - l:] == v:
				return 1
		elif fm != '':
			if dato == v:
				return 1
		else:
			if dato.find(v) >= 0:
				return 1

	return -1

def Forma_Sele(emges, apl, emp, fi, idx, rg, preg, preg_bi, s_campos, s_orden, s_descen):
	cpc = {}
	for cp, ope, v, colu, num in preg:
		if colu != '':
			if colu >= 0:
				continue
		if ope == '=':
			ope = '=='
		if cp == -1:
			vreg = idx
		else:
			try:
				vreg = rg[cp]
			except:
				vreg = Calcu_Campo(emges, apl, emp, fi, idx, rg, cp)
				cpc[cp] = vreg

		if colu < 0:
			fila = -(colu + 1)
			try:
				vreg = vreg[fila]
			except:
				if num == 0:
					vreg = ''
				else:
					vreg = 0

			colu = ''
		if num == 1:
			if vreg == None:
				vreg = -9999
			if eval(`vreg` + ope + `v`) != 1:
				return ([], [])
		else:
			vl, vreg = v.upper(), vreg.upper()
			if vl.find('&') >= 0:
				vl = vl.split('&')
			else:
				vl = [
					vl]
			for v in vl:
				if ope == '==':
					if concuerda(vreg, v) == -1:
						return ([], [])
				elif ope == '<>':
					if concuerda(vreg, v) == 1:
						return ([], [])
				else:
					if eval(`vreg` + ope + `v`) != 1:
						return ([], [])

	for nc in preg_bi.keys():
		sel_por_camp = 0
		try:
			reg = rg[nc]
		except:
			reg = []

		if len(reg) > 0:
			try:
				if type(reg[0]) != ListType:
					reg1 = reg
					reg = []
					for ln in reg1:
						reg.append([ln])

			except:
				reg = []

		selecciono = 0
		for ln in reg:
			for ope, v, colu, num in preg_bi[nc]:
				try:
					vreg = ln[colu]
				except:
					vreg = ''

				if ope == '=':
					ope = '=='
				if num > 0:
					if vreg == None:
						vreg = -9999
					if vreg == '':
						vreg = 0
					if eval(`vreg` + ope + `v`) != 1:
						break
					selecciono = selecciono + 1
				else:
					vl, vreg = v.upper(), vreg.upper()
					if vl.find('&') >= 0:
						vl = vl.split('&')
					else:
						vl = [
							vl]
					slx = 0
					for v in vl:
						if ope == '==':
							if concuerda(vreg, v) == -1:
								continue
						else:
							if ope == '<>':
								if concuerda(vreg, v) == 1:
									continue
							else:
								if eval(`vreg` + ope + `v`) != 1:
									continue
						slx = slx + 1
						break

					if slx == 0:
						break
					selecciono = selecciono + 1

			if selecciono == len(preg_bi[nc]):
				sel_por_camp = 1
				break
			else:
				selecciono = 0

		if sel_por_camp == 0:
			return ([], [])

	res, orden = [], []
	j = 0
	for k in s_campos:
		if k < 0:
			if k == -1:
				res.append(idx)
			else:
				res.append('')
		else:
			try:
				res.append(rg[k])
			except:
				if cpc.has_key(k):
					res.append(cpc[k])
				else:
					cpc[k] = Calcu_Campo(emges, apl, emp, fi, idx, rg, k)
					res.append(cpc[k])

		j = j + 1

	if s_orden != []:
		l = len(s_orden)
		for s in range(l):
			k = s_orden[s]
			if k == -1:
				if s_descen[s] != 0:
					orden.append(Rever_Cd(idx))
				else:
					orden.append(idx)
			else:
				try:
					v = rg[k]
				except:
					if cpc.has_key(k):
						v = cpc[k]
					else:
						v = Calcu_Campo(emges, apl, emp, fi, idx, rg, k)

				if s_descen[s] != 0:
					if type(v) == StringType:
						orden.append(Rever_Cd(v))
					else:
						try:
							orden.append(-v)
						except:
							orden.append(999999)

				else:
					orden.append(v)

	return (
		res,
		orden)


def selec(gpx, archi, campos=[], preguntas=[], orden=[], cabe=[], informa='', cl=''):
	emges, apl, emp = gpx
	if cl != '':
		try:
			racs = cl.rs_ar[apl][archi]
		except:
			racs = ''

		if racs != '':
			preg = copia_rg(racs)
			preg.extend(preguntas)
			preguntas = preg
	fl_tipo = DCN[(apl,
	               archi)][3]
	fz = F[(emges,
	        apl,
	        emp,
	        archi)]
	if fz.snc != '':
		fz = fz.snc
	if fl_tipo == 'i':
		res = fz.keys()
		res.sort()
		return res
	if type(campos) != ListType:
		campos = campos.split()
	if type(orden) != ListType:
		orden = orden.split()
	inx, preg, preg_cd, busca_inx, preg_bi, inx_dh = cons_preguntas(emges, apl, emp, archi, preguntas)
	s_campos = Tradu_Campos((apl,
	                         archi), campos, noexis=1)
	s_orden, s_descen = [], []
	for cp in orden:
		desc = 0
		if cp[:1] == '-':
			desc = 1
			cp = cp[1:]
		if cp == '0':
			s_orden.append(-1)
			s_descen.append(desc)
		else:
			nc = Campo_a_Num((apl,
			                  archi), cp)
			if nc == -1:
				continue
			s_orden.append(nc)
			s_descen.append(desc)

	ers = 0
	res, lorden = [], []
	j = 0
	n_cps = len(s_campos)
	if s_campos != []:
		if busca_inx != '':
			cpx, cdss = busca_inx
			cursor = fz.i_cursor(cpx)
			procesados = {}
			for cda in cdss:
				lx, ers = len(cda), 0
				try:
					ids = cursor.set_range(cda, sl_cod=1)
					cds, idx = lista(ids, MX, 2)
					if cds[:lx] == cda[:lx]:
						if not procesados.has_key(idx):
							rg = _Lee(fz, idx)
							if rg != 1:
								v, orden = Forma_Sele(emges, apl, emp, archi, idx, rg, preg, preg_bi, s_campos, s_orden,
								                      s_descen)
								procesados[idx] = ''
								if v != []:
									if n_cps == 1:
										res.append(v[0])
									else:
										res.append(v)
									if orden != []:
										orden.append(j)
										lorden.append(orden)
										j = j + 1
					else:
						ers = 1
				except:
					ers = 1

				while ers == 0:
					try:
						ids = cursor.next(sl_cod=1)
						cds, idx = lista(ids, MX, 2)
						if cds[:lx] == cda[:lx]:
							if not procesados.has_key(idx):
								rg = _Lee(fz, idx)
								if rg != 1:
									v, orden = Forma_Sele(emges, apl, emp, archi, idx, rg, preg, preg_bi, s_campos,
									                      s_orden, s_descen)
									procesados[idx] = ''
									if v != []:
										if n_cps == 1:
											res.append(v[0])
										else:
											res.append(v)
										if orden != []:
											orden.append(j)
											lorden.append(orden)
											j = j + 1
						else:
							ers = 1
					except:
						ers = 1

			del cursor
		else:
			info = 0
			if inx_dh == []:
				if preg_cd == []:
					rgzs = fz.keys()
					maxreg = len(rgzs)
				else:
					desde = hasta = igual = ''
					for ope, v in preg_cd:
						if ope == '>=' or ope == '>':
							desde = v
						if ope == '<=' or ope == '<':
							hasta = v
						if ope == '=':
							igual = v

					if igual.find('^') >= 0 or igual.find('[') >= 0 or igual == ']':
						igual = ''
					cursor = fz.cursor()
					rgzs = []
					if igual != '':
						igual = igual.replace('?', '&')
						v = igual.split('&')
						for vl in v:
							inicio = 0
							if vl[-1:] == ']':
								vl = vl[:-1]
								inicio = len(vl)
							try:
								ids = cursor.set_range(vl, sl_cod=1)
								rgzs.append(ids)
								ers = 0
								if inicio > 0:
									while ers == 0:
										try:
											ids = cursor.next(sl_cod=1)
											if ids[:inicio] != vl:
												ers = 1
											else:
												rgzs.append(ids)
										except:
											ers = 1

							except:
								pass

					else:
						if desde != '' or hasta != '':
							if hasta == ']':
								hasta = ''
							if desde == ']':
								desde = ''
							try:
								ids = cursor.set_range(desde, sl_cod=1)
								rgzs.append(ids)
								inicio = 0
								if hasta[-1:] == ']':
									hasta = hasta[:-1]
									inicio = len(hasta)
								ers = 0
								while ers == 0:
									try:
										ids = cursor.next(sl_cod=1)
										if hasta != '':
											if inicio == 0:
												if ids > hasta:
													ers = 1
												else:
													rgzs.append(ids)
											elif ids[:inicio] > hasta:
												ers = 1
											else:
												rgzs.append(ids)
										else:
											rgzs.append(ids)
									except:
										ers = 1

							except:
								pass

						else:
							rgzs = fz.keys()
					maxreg = len(rgzs)
					del cursor
			else:
				cpi, desde, hasta = inx_dh
				nc = Campo_a_Num((apl,
				                  archi), cpi)
				if nc >= 0:
					fmt = DCI[(apl,
					           archi)][nc][0]
					if type(fmt) == type([]):
						fmt = fmt[0]
					if fmt == 'i' or fmt == 'd':
						if desde == '':
							desde = '-9999'
						if hasta == '':
							hasta = '99999'
					else:
						if hasta == '':
							hasta = chr(255)
						if desde == '':
							desde = chr(0)
				rgzs = fz.i_selec(cpi, desde, hasta)
				if rgzs == None:
					rgzs = []
				maxreg = len(rgzs)
			solo_pcd = 1
			if s_campos != [-1]:
				solo_pcd = 0
			if s_orden != []:
				if s_orden != [-1]:
					solo_pcd = 0
			if s_descen != []:
				solo_pcd = 0
			if solo_pcd == 1:
				for psx in preg:
					if psx[0] != -1:
						solo_pcd = 0
						break

			if solo_pcd == 1:
				s_orden = []
			if maxreg > 10000 and solo_pcd == 0:
				if informa != '' and len(fz.vx) < 10000:
					informa(cl, 0, maxreg, 'Seleccionando registros de ' + archi, 's')
					info = 1
			nrp = 0
			for idx in rgzs:
				if info == 1:
					nrp = nrp + 1
					if nrp % 1000.0 == 0 or maxreg == nrp:
						informa(cl, nrp)
				if preg_cd != []:
					malo = 0
					for k in preg_cd:
						ope, v = k
						if ope == '=':
							ope = '=='
						vl, vreg = v.upper(), idx
						if vl.find('&') >= 0:
							vl = vl.split('&')
						else:
							vl = [
								vl]
						for v in vl:
							if ope == '==':
								if concuerda(vreg, v) == -1:
									malo = 1
							else:
								if ope == '<>':
									if concuerda(vreg, v) == 1:
										malo = 1
								else:
									if eval(`vreg` + ope + `v`) != 1:
										malo = 1
							if malo == 1:
								break

						if malo == 1:
							break

					if malo == 1:
						continue
				if solo_pcd == 0:
					rg = _Lee(fz, idx)
					v, orden = Forma_Sele(emges, apl, emp, archi, idx, rg, preg, preg_bi, s_campos, s_orden, s_descen)
					if v != []:
						if n_cps == 1:
							res.append(v[0])
						else:
							res.append(v)
						if orden != []:
							orden.append(j)
							lorden.append(orden)
							j = j + 1
				else:
					res.append(idx)

	if s_orden != []:
		rs = []
		if lorden != []:
			lorden.sort()
			cn = len(lorden[0]) - 1
			for k in lorden:
				j = k[cn]
				rs.append(res[j])

		res = rs

	return res


if __name__ == '__main__':
	rg = lee_dc(lee_dc, gpx, 'alb-venta', '0000000001', '', rels='s')
