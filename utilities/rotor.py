# uncompyle6 version 3.2.3
# Python bytecode 2.7 (62211)
# Decompiled from: Python 2.7.13 (v2.7.13:a06454b1afa1, Dec 17 2016, 20:42:59) [MSC v.1500 32 bit (Intel)]
# Embedded file name: rotor.py
"""This module was written by Gerd Woetzel:
   http://mail.python.org/pipermail/python-list/2005-January/261304.html


   This module is derived from Modules/rotormodule.c and translated
   into Python. I have appended the Copyright by Lance Ellinghouse
   below. The rotor module has been removed from the Python 2.4
   distribution because

      the rotor module uses an insecure algorithm and is deprecated.
      ==============================================================

   Of course, this does still hold. However, I think this module might
   be used and adapted for demonstration purposes and might help some
   poor users who have encrypted (or obfuscated) some old stuff with
   the rotor module and have no access to older Python versions any
   more.

   Documentation can be found in

   Python Library Reference,  Guido van Rossum, Fred L. Drake, Jr., editor,
   PythonLabs, Release 2.3.4 May 20, 2004
   <http://www.python.org/doc/2.3.4/lib/module-rotor.html>

   #####################################################################
   Copyright 1994 by Lance Ellinghouse,
   Cathedral City, California Republic, United States of America.

                        All Rights Reserved

   Permission to use, copy, modify, and distribute this software and its
   documentation for any purpose and without fee is hereby granted,
   provided that the above copyright notice appear in all copies and that
   both that copyright notice and this permission notice appear in
   supporting documentation, and that the name of Lance Ellinghouse
   not be used in advertising or publicity pertaining to distribution
   of the software without specific, written prior permission.

   LANCE ELLINGHOUSE DISCLAIMS ALL WARRANTIES WITH REGARD TO
   THIS SOFTWARE, INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND
   FITNESS, IN NO EVENT SHALL LANCE ELLINGHOUSE BE LIABLE FOR ANY SPECIAL,
   INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING
   FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT,
   NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION
   WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
   #####################################################################
"""


# noinspection PyPep8Naming
class newrotor(object):
	def __init__(self, key, n_rotors=6):
		self.n_rotors = n_rotors
		self.setkey(key)

	def setkey(self, key):
		self.key = key
		self.rotors = None
		self.positions = [None, None]
		return

	def encrypt(self, buf):
		self.positions[0] = None
		return self.cryptmore(buf, 0)

	def encryptmore(self, buf):
		return self.cryptmore(buf, 0)

	def decrypt(self, buf):
		self.positions[1] = None
		return self.cryptmore(buf, 1)

	def decryptmore(self, buf):
		return self.cryptmore(buf, 1)

	def cryptmore(self, buf, do_decrypt):
		size, nr, rotors, pos = self.get_rotors(do_decrypt)
		outbuf = []
		append = outbuf.append
		for c in map(ord, buf):
			if do_decrypt:
				for i in xrange(nr - 1, -1, -1):
					c = pos[i] ^ rotors[i][c]

			else:
				for i in xrange(nr):
					c = rotors[i][c ^ pos[i]]

			append(c)
			pnew = 0
			for i in xrange(nr):
				pnew = (pos[i] + (pnew >= size) & 255) + rotors[i][size]
				pos[i] = pnew % size

		return ('').join(map(chr, outbuf))

	def get_rotors(self, do_decrypt):
		nr = self.n_rotors
		rotors = self.rotors
		positions = self.positions[do_decrypt]
		if positions is None:
			if rotors:
				positions = list(rotors[3])
			else:
				self.size = size = 256
				id_rotor = range(size + 1)
				rand = random_func(self.key)
				E = []
				D = []
				positions = []
				for i in xrange(nr):
					i = size
					positions.append(rand(i))
					erotor = id_rotor[:]
					drotor = id_rotor[:]
					drotor[i] = erotor[i] = 1 + 2 * rand(i / 2)
					while i > 1:
						r = rand(i)
						i -= 1
						er = erotor[r]
						erotor[r] = erotor[i]
						erotor[i] = er
						drotor[er] = i

					drotor[erotor[0]] = 0
					E.append(tuple(erotor))
					D.append(tuple(drotor))

				self.rotors = rotors = (tuple(E), tuple(D), size, tuple(positions))
			self.positions[do_decrypt] = positions
		return (rotors[2], nr, rotors[do_decrypt], positions)


def random_func(key):
	mask = 65535
	x = 995
	y = 576
	z = 767
	for c in map(ord, key):
		x = (x << 3 | x >> 13) + c & mask
		y = ((y << 3 | y >> 13) ^ c) & mask
		z = (z << 3 | z >> 13) - c & mask

	maxpos = mask >> 1
	mask += 1
	if x > maxpos:
		x -= mask
	if y > maxpos:
		y -= mask
	if z > maxpos:
		z -= mask
	y |= 1
	x = 171 * (x % 177) - 2 * (x / 177)
	y = 172 * (y % 176) - 35 * (y / 176)
	z = 170 * (z % 178) - 63 * (z / 178)
	if x < 0:
		x += 30269
	if y < 0:
		y += 30307
	if z < 0:
		z += 30323

	def rand(n, seed=((x, y, z))):
		x, y, z = seed[0]
		seed[0] = (171 * x % 30269, 172 * y % 30307, 170 * z % 30323)
		return int((x / 30269.0 + y / 30307.0 + z / 30323.0) * n) % n

	return rand
