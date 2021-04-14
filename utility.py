if __name__ == '__main__':
	txt = 'id="Text_box" o:spid="_x0000_s1026" type="#_x0000_t202" stroked="f"'
	#txt += ' ' \
		 # '
	txt+= 'style="position:absolute;margin-left:-355.05pt;margin-top:378.1pt;width:592.65pt;height:20.25pt;rotation:-90;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top"'

	txt = txt.replace(':', '_').split(' ')

	ar = ''
	_vars = list()
	for ln in txt:
		try:var, value = ln.split('=')
		except:print ln
		new_var = ''
		for i in var:
			if i == i.upper() and not i.isdigit() and i != '_':
				new_var += '_' + i.lower()
			else:
				new_var += i
		if ar:
			ar += ' '
		ar += new_var + '=' + value + ','
		_vars.append(new_var)
		res = '\t\tself.' + new_var + ' = ' + new_var
		print res

	print '\n' + ar + '\n'
	x = 0
	xml = ''
	for var in _vars:
		print '\tdef get_%s(self):\n\t\treturn self.%s\n' % (var, var)

		print '\tdef set_%s(self, %s):\n\t\tself.%s = %s\n' % (var, var, var, var)

		t = txt[x].split('=')[0]
		xml += '\t\tif self.get_%s():\n\t\t\targs.append(%s%s=%s%s+self.get_%s()+%s%s%s)\n\n' % \
			   (var, "'", t, '"', "'", var, "'", '"', "'")
		x += 1

	print '\n'
	print xml
