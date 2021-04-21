# coding=utf-8
from db import error, Trae_Fila, lee, rg_vacio, gpx, cl, copia_rg, FDC, lista, Num_aFecha, lee_dc, selec

import json

# noinspection PyPep8Naming
def DameRedondeos(rg='', cl=cl, gpx=gpx):
	if rg == '':
		rg = lee(cl, gpx, 'parametros', '0')
	if rg == 1:
		rg = rg_vacio(gpx[1], 'parametros')

	if rg["P_REDONDEO"] == []:
		ls = [0] * 6
	else:
		ls = rg["P_REDONDEO"][0]
	dc = {}  # uds|pvp|dto|neto|importe|pcoste
	dc['unidades'] = ls[0]
	dc['precio_venta'] = ls[1]
	dc['descuento'] = ls[2]
	dc['porcentaje'] = ls[2]
	dc['neto'] = ls[3]
	dc['importe'] = ls[4]
	dc['precio_coste'] = ls[5]
	return dc


# noinspection PyPep8Naming
def GetColumnasPT_ALB(arg=''):
	columnas = []
	columnas.append('albaran')
	columnas.append('fecha_generacion')
	if arg == 'v':
		columnas.append('fecha_albaran')
		columnas.append('tipo')
		columnas.append('deno_tipo')
		columnas.append('importe')
		columnas.append('factura')
		columnas.append('fecha_factura')
		columnas.append('forma_pago')
		columnas.append('deno_forma_pago')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasACC_CK(arg=''):
	columnas = []
	columnas.append('numero')
	columnas.append('padre')
	columnas.append('descripcion')
	columnas.append('tipo')
	columnas.append('posibles_valores')
	columnas.append('obligatoria')
	columnas.append('valor_defecto')
	columnas.append('respuesta')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCK_LNA(arg=''):
	columnas = []
	columnas.append('numero')
	columnas.append('padre')
	columnas.append('descripcion')
	columnas.append('tipo')
	columnas.append('posibles_valores')
	columnas.append('obligatoria')
	columnas.append('valor_defecto')
	columnas.append('periodicidad')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCCK_LNA(arg=''):
	columnas = []
	columnas.append('numero')
	columnas.append('padre')
	columnas.append('descripcion')
	columnas.append('tipo')
	columnas.append('posibles_valores')
	columnas.append('obligatoria')
	columnas.append('valor_defecto')
	columnas.append('respuesta')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasACC_CAR(arg=''):
	columnas = []
	columnas.append('caracteristica')
	if arg == 'v':
		columnas.append('deno_caracteristica')
	columnas.append('tipo_dato')
	columnas.append('valor_texto')
	columnas.append('valor_numero')
	columnas.append('valor_fecha')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


##

def GetColumnasNS_CAR(arg=''):
	columnas = []
	columnas.append('caracteristica')
	if arg == 'v':
		columnas.append('deno_caracteristica')
	columnas.append('tipo_dato')
	columnas.append('valor_texto')
	columnas.append('valor_numero')
	columnas.append('valor_fecha')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


##

def GetColumnasNS_FECHAS(arg=''):
	columnas = []
	columnas.append('tipo_fecha')
	if arg == 'v':
		columnas.append('deno_tipo_fecha')
	columnas.append('ultima_fecha')
	columnas.append('periodicidad')
	columnas.append('tipo_periodicidad')
	columnas.append('proxima_fecha')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


##

def GetColumnasPT_MATNS(arg=''):
	columnas = []
	columnas.append('id_linea_materiales')
	if arg == 'v':
		columnas.append('cdar')
		columnas.append('deno_cdar')
	columnas.append('marca')
	if arg == 'v':
		columnas.append('deno_marca')
	columnas.append('numero_serie')
	columnas.append('ubicacion')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasACC_FECHAS(arg=''):
	columnas = []
	columnas.append('tipo_fecha')
	if arg == 'v':
		columnas.append('deno_tipo_fecha')
	columnas.append('ultima_fecha')
	columnas.append('periodicidad')
	columnas.append('tipo_periodicidad')
	columnas.append('proxima_fecha')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasXX_FECHAS(arg=''):
	columnas = []
	columnas.append('tipo_fecha')
	if arg == 'v':
		columnas.append('deno_tipo_fecha')
	columnas.append('periodicidad')
	columnas.append('tipo_periodicidad')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCL_TELS(arg=''):
	columnas = []
	columnas.append('telefono')
	columnas.append('contacto')
	columnas.append('cif')
	columnas.append('mail')
	columnas.append('fax')
	columnas.append('defecto')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasPTR_REPRV(arg=''):
	columnas = []
	columnas.append('cdar')
	if arg == 'v':
		columnas.append('deno')
	columnas.append('unidades')
	columnas.append('horas')
	columnas.append('coste_hora')
	columnas.append('precio_coste')
	columnas.append('precio_venta')
	columnas.append('importe')
	columnas.append('almacen')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCL_PVES(arg=''):
	columnas = []
	columnas.append('cdar')
	if arg == 'v':
		columnas.append('deno')
	columnas.append('descuento')
	columnas.append('descuento_2')
	columnas.append('descuento_3')
	columnas.append('descuento_4')
	columnas.append('precio_venta')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasPAM_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	if arg == 'v':
		columnas.append('stock')
		columnas.append('stock_minimo')
		columnas.append('stock_maximo')
		columnas.append('pedir_de')
	columnas.append('unidades')
	columnas.append('unidades_servidas')
	columnas.append('diccionario')
	if arg == 'v':
		columnas.append('deno_diccionario')
	columnas.append('documento')
	columnas.append('id_linea')
	columnas.append('info')
	columnas.append('estado_linea')
	if arg == 'v':
		columnas.append('deno_estado_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasDL_TELS(arg=''):
	columnas = []
	columnas.append('telefono')
	columnas.append('contacto')
	columnas.append('cif')
	columnas.append('mail')
	columnas.append('fax')
	columnas.append('defecto')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCL_DIRFAC(arg=''):
	columnas = []
	columnas.append('direccion')
	columnas.append('codigo_postal')
	columnas.append('poblacion')
	columnas.append('provincia')
	columnas.append('pais')
	if arg == 'v':
		columnas.append('deno_pais')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasDIRECCIONES(arg=''):
	columnas = []
	columnas.append('direccion')
	columnas.append('codigo_postal')
	columnas.append('poblacion')
	columnas.append('provincia')
	columnas.append('pais')
	if arg == 'v':
		columnas.append('deno_pais')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasDescompuesto(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('precio_venta')
	if arg == 'v':
		columnas.append('importe')
	columnas.append('precio_coste')
	if arg == 'v':
		columnas.append('coste_total')
	columnas.append('subdescompuesto')
	columnas.append('certificado')
	columnas.append('servido')
	columnas.append('ejecutado')
	if arg == 'v':
		columnas.append('unidades_totales')
		columnas.append('unidades_certificadas')
		columnas.append('unidades_servidas')
		columnas.append('unidades_ejecutadas')
	columnas.append('id')
	columnas.append('unidades_albaran')

	nulo = {}
	for name in columnas:
		value = ''
		if name in ['precio_venta', 'precio_coste', 'unidades', 'unidades_albaran']:
			value = 0.
		nulo[name] = value

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k

	return columnas, nulo, dc


# noinspection PyPep8Naming
def GetColumnasD_LNA(arg='', sub=False):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades_presupuestos')
	columnas.append('unidades_pedidos_cl')
	columnas.append('unidades_partes')
	columnas.append('unidades_alb-venta')
	if arg == 'v':
		columnas.append('total_alb-venta')
	columnas.append('unidades_obras')
	columnas.append('unidades_ejecucion')
	if arg == 'v':
		columnas.append('total_ejecucion')
	columnas.append('unidades_obras_certificacion')
	if arg == 'v':
		columnas.append('total_obras_certificacion')
		columnas.append('total_documento')
	columnas.append('precio_venta')
	if arg == 'v':
		columnas.append('importe')
	columnas.append('precio_coste')
	if arg == 'v':
		columnas.append('coste_total')
	columnas.append('id_linea')
	if sub:
		columnas.append('id_linea_padre')
	else:
		columnas.append('subdescompuesto')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k

	return dc


# noinspection PyPep8Naming
def GetColumnasD_SUBLNA(arg='', GetColumnasD_LNA=GetColumnasD_LNA):
	return GetColumnasD_LNA(arg, sub=True)


# noinspection PyPep8Naming
def GetColumnasPT_MAT(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades_previstas')
	columnas.append('unidades')
	columnas.append('unidades_servidas')
	columnas.append('precio_venta')
	columnas.append('descuento')
	columnas.append('descuento_2')
	columnas.append('descuento_3')
	columnas.append('descuento_4')
	if arg == 'v':
		columnas.append('neto')
		columnas.append('importe')
	columnas.append('precio_coste')
	columnas.append('facturable')
	columnas.append('almacen')
	if arg == 'v':
		columnas.append('deno_almacen')
	columnas.append('partida')
	columnas.append('capitulo')
	columnas.append('referencia_capitulo')
	columnas.append('subcapitulo')
	columnas.append('referencia_subcapitulo')
	columnas.append('numeros_serie')
	columnas.append('ubicacion')
	columnas.append('observaciones')
	columnas.append('descripcion_extendida_articulo')
	columnas.append('descripcion_extendida_capitulo')
	columnas.append('descripcion_extendida_subcapitulo')
	columnas.append('id_linea')
	columnas.append('id_linea_pedido')
	columnas.append('descompuesto')
	columnas.append('id_descompuesto')
	columnas.append('info')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasFAC_COMP(arg=''):
	columnas = []
	columnas.append('cdar')
	if arg == 'v':
		columnas.append('deno')

	columnas.append('unidades_a_realizar')
	columnas.append('cantidad_unitaria')
	columnas.append('unidades_realizadas')
	columnas.append('unidades_roturas')
	columnas.append('id_linea_parte')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasTV_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	if arg == 'v':
		columnas.append('deno')
	columnas.append('unidades')
	columnas.append('precio_coste')
	if arg == 'v':
		columnas.append('valor')
	columnas.append('almacen')
	if arg == 'v':
		columnas.append('deno_alm')
		columnas.append('stock_alm')
	columnas.append('nserie')
	columnas.append('lotes')
	columnas.append('descompuesto')
	columnas.append('vacio')
	columnas.append('vacio2')
	columnas.append('id_linea')
	columnas.append('info')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasDL_BANS(arg=''):
	columnas = []
	columnas.append('iban')
	columnas.append('swift')
	columnas.append('defecto')
	if arg == 'v':
		columnas.append('banco')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_CKG(arg=''):
	columnas = []
	columnas.append('checklist')
	if arg == 'v':
		columnas.append('deno_cchecklist')
	columnas.append('tipo_accion')
	if arg == 'v':
		columnas.append('deno_tipo_accion')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCL_VD(arg=''):
	columnas = []
	columnas.append('vendedor')
	if arg == 'v':
		columnas.append('deno')

	columnas.append('defecto')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCL_BANS(arg=''):
	columnas = []
	columnas.append('iban')
	columnas.append('swift')
	columnas.append('defecto')
	if arg == 'v':
		columnas.append('banco')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasDL_VD(arg=''):
	columnas = []
	columnas.append('vendedor')
	if arg == 'v':
		columnas.append('deno')

	columnas.append('defecto')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasIN_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	if arg == 'v':
		columnas.append('deno')
	columnas.append('unidades')
	columnas.append('pmc')
	if arg == 'v':
		columnas.append('importe')
	columnas.append('unidades_anteriores')
	if arg == 'v':
		columnas.append('diferencia')

	columnas.append('libre1')
	columnas.append('libre2')
	columnas.append('libre3')
	columnas.append('usuario')
	columnas.append('fechayhora')
	columnas.append('rnd')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasFPA_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades_previstas')
	columnas.append('unidades_utilizadas')
	columnas.append('precio_coste')
	columnas.append('tolerancia_superior')
	columnas.append('tolerancia_inferior')
	columnas.append('almacen')
	if arg == 'v':
		columnas.append('deno_almacen')
	columnas.append('id_linea')
	columnas.append('info')
	columnas.append('fase')
	if arg == 'v':
		columnas.append('deno_fase')
	columnas.append('unidades_roturas')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_ESC(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('precio_coste')
	if arg == 'v':
		columnas.append('total_coste')
	columnas.append('tolerancia_inferior')
	columnas.append('tolerancia_superior')
	columnas.append('tipo')
	if arg == 'v':
		columnas.append('stock')
	columnas.append('fase')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCUO_LNA(arg=''):
	columnas = []
	if arg == 'v':
		columnas.append('numero')

	columnas.append('contrato')
	columnas.append('cliente')
	if arg == 'v':
		columnas.append('deno_cliente')
	columnas.append('delegacion')
	columnas.append('codigo_delegacion')
	if arg == 'v':
		columnas.append('deno_delegacion')

	columnas.append('forma_pago')
	if arg == 'v':
		columnas.append('deno_forma_pago')
	columnas.append('importe')
	columnas.append('serie')
	columnas.append('fecha_albaran')
	columnas.append('cd_albaran')
	columnas.append('facturar')
	columnas.append('rg')
	columnas.append('cd_factura')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_DESC(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('precio_coste')
	if arg == 'v':
		columnas.append('total_coste')
		columnas.append('precio_venta')
		columnas.append('importe')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_OFV(arg=''):
	columnas = []
	columnas.append('desde_fecha')
	columnas.append('hasta_fecha')
	columnas.append('precio_venta')
	columnas.append('descuento')
	columnas.append('tarifa')
	columnas.append('unidades_minimas')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasOBM_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('unidades_enviadas')
	columnas.append('precio_coste')
	columnas.append('descuento')
	if arg == 'v':
		columnas.append('valor')
	columnas.append('tipo')
	if arg == 'v':
		columnas.append('deno_tipo')
	columnas.append('estado_linea')
	if arg == 'v':
		columnas.append('deno_estado')

	columnas.append('capitulo')
	columnas.append('referencia_capitulo')
	columnas.append('subcapitulo')
	columnas.append('referencia_subcapitulo')
	columnas.append('descripcion_extendida_articulo')
	columnas.append('descompuesto')
	columnas.append('info')
	if arg == 'v':
		columnas.append('stock_obra')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasOB_LNA(arg='', ventana=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('unidades_ejecutadas')
	columnas.append('unidades_certificadas')
	if ventana == 'obras_certifica':
		columnas.append('unidades_albaran')
	columnas.append('precio_venta')
	columnas.append('descuento')
	columnas.append('iva')
	if arg == 'v':
		columnas.append('neto')
		columnas.append('importe')
	columnas.append('precio_coste')
	if arg == 'v':
		columnas.append('margen_beneficio')
		columnas.append('porcentaje_certificado')
		columnas.append('porcentaje_ejecutado')
	columnas.append('partida')
	columnas.append('capitulo')
	columnas.append('referencia_capitulo')
	columnas.append('subcapitulo')
	columnas.append('referencia_subcapitulo')
	columnas.append('libre')
	columnas.append('presupuesto')
	columnas.append('descompuesto')
	columnas.append('descripcion_extendida_articulo')
	columnas.append('descripcion_extendida_capitulo')
	columnas.append('descripcion_extendida_subcapitulo')
	columnas.append('info')
	columnas.append('id_descompuesto')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_COM(arg=''):
	columnas = []
	columnas.append('vendedor')
	if arg == 'v':
		columnas.append('deno')
	columnas.append('comision_sin_dto')
	columnas.append('comision_con_dto')
	columnas.append('porcentaje_minimo')
	columnas.append('porcentaje_maximo')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCN_MAT(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('marca')
	if arg == 'v':
		columnas.append('deno_marca')
	columnas.append('numero_serie')
	columnas.append('ubicacion')
	columnas.append('periodicidad_revision')
	columnas.append('tipo_periodicidad_revision')
	columnas.append('fecha_revision')
	columnas.append('periodicidad_certificacion')
	columnas.append('tipo_periodicidad_certificacion')
	columnas.append('fecha_certificacion')
	columnas.append('fecha_baja')
	columnas.append('id_linea')
	columnas.append('info')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasPP_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('referencia')
	columnas.append('unidades')
	columnas.append('unidades_servidas')
	columnas.append('precio_coste')
	columnas.append('descuento')
	columnas.append('iva')
	if arg == 'v':
		columnas.append('neto')
		columnas.append('importe')
	columnas.append('fecha_servir')
	columnas.append('estado_linea')
	if arg == 'v':
		columnas.append('deno_estado')
	columnas.append('descripcion_extendida')
	columnas.append('id_linea_pedido_cl')
	columnas.append('info')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasPPP_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('referencia')
	columnas.append('unidades')
	columnas.append('precio_coste')
	columnas.append('unidades_disponibles')
	columnas.append('precio_proveedor')
	columnas.append('fecha_servir')
	columnas.append('unidades_pedidas')
	columnas.append('estado_linea')
	if arg == 'v':
		columnas.append('deno_estado')
	columnas.append('info')
	columnas.append('observaciones')
	if arg == 'v':
		columnas.append('tiene_observaciones')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAC_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('referencia')
	columnas.append('unidades')
	columnas.append('precio_coste')
	columnas.append('descuento')
	columnas.append('iva')
	if arg == 'v':
		columnas.append('neto')
		columnas.append('importe')
	columnas.append('almacen')
	if arg == 'v':
		columnas.append('nombre_almacen')
	columnas.append('id_linea_pedido')
	columnas.append('numeros_serie')
	columnas.append('info')
	columnas.append('descripcion_extendida')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_POP(arg=''):
	columnas = []
	columnas.append('proveedor')
	if arg == 'v':
		columnas.append('deno_proveedor')
	columnas.append('precio_coste')
	columnas.append('descuento')
	if arg == 'v':
		columnas.append('coste_total')
	columnas.append('ultima_compra')
	columnas.append('albaran_compra')
	columnas.append('referencia')
	columnas.append('precio_acordado')
	columnas.append('defecto')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_TARI(arg=''):
	columnas = []
	columnas.append('tarifa')
	if arg == 'v':
		columnas.append('deno')
	columnas.append('margen')
	columnas.append('precio_venta')
	columnas.append('comision')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAR_STOK(arg=''):
	columnas = []
	columnas.append('almacen')
	if arg == 'v':
		columnas.append('deno_almacen')
	columnas.append('ubicacion')
	columnas.append('stock')
	columnas.append('stock_imputado')
	if arg == 'v':
		columnas.append('por_recibir')
		columnas.append('por_utilizar')
		columnas.append('stock_real')
		columnas.append('valor')
	columnas.append('minimo')
	columnas.append('maximo')
	columnas.append('pedir_de')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCN_SER(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('precio_venta')
	columnas.append('fecha_alta')
	columnas.append('fecha_baja')
	columnas.append('periodicidad')
	columnas.append('tipo_periodicidad')
	columnas.append('enero')
	columnas.append('febrero')
	columnas.append('marzo')
	columnas.append('abril')
	columnas.append('mayo')
	columnas.append('junio')
	columnas.append('julio')
	columnas.append('agosto')
	columnas.append('septiembre')
	columnas.append('octubre')
	columnas.append('novienbre')
	columnas.append('diciembre')
	columnas.append('descripcion_extendida')
	columnas.append('id_linea')
	columnas.append('info')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasPC_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('precio_venta')
	columnas.append('descuento')
	columnas.append('descuento_2')
	columnas.append('descuento_3')
	columnas.append('descuento_4')
	if arg == 'v':
		columnas.append('neto')
		columnas.append('importe')
	columnas.append('precio_coste')
	if arg == 'v':
		columnas.append('beneficio')
	columnas.append('iva')
	columnas.append('opcion')
	columnas.append('partida')
	columnas.append('capitulo')
	columnas.append('referencia_capitulo')
	columnas.append('subcapitulo')
	columnas.append('referencia_subcapitulo')
	columnas.append('descompuesto')
	columnas.append('descripcion_extendida_articulo')
	columnas.append('descripcion_extendida_capitulo')
	columnas.append('descripcion_extendida_subcapitulo')
	columnas.append('info')
	columnas.append('orden')
	columnas.append('id_descompuesto')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasAV_LNA(arg='', albaranar_partes=False):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	if albaranar_partes:
		columnas.append('unidades_pendientes')
	columnas.append('unidades')
	columnas.append('precio_venta')
	columnas.append('descuento')
	columnas.append('descuento_2')
	columnas.append('descuento_3')
	columnas.append('descuento_4')
	if arg == 'v':
		columnas.append('neto')
		columnas.append('importe')
	if albaranar_partes:
		columnas.append('facturable')
	columnas.append('precio_coste')
	columnas.append('iva')
	columnas.append('almacen')
	if arg == 'v':
		columnas.append('nombre_almacen')
	columnas.append('partida')
	columnas.append('capitulo')
	columnas.append('referencia_capitulo')
	columnas.append('subcapitulo')
	columnas.append('referencia_subcapitulo')
	columnas.append('numeros_serie')
	columnas.append('descripcion_extendida_articulo')
	columnas.append('descripcion_extendida_capitulo')
	columnas.append('descripcion_extendida_subcapitulo')
	columnas.append('info')
	columnas.append('lotes')
	columnas.append('unidades_abonadas')
	columnas.append('descompuesto')
	columnas.append('id_descompuesto')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasCOB_LNA(arg='', GetColumnasAV_LNA=GetColumnasAV_LNA):
	# Las lineas tienen que ser iguales que la de los albaranes
	return GetColumnasAV_LNA(arg)


# noinspection PyPep8Naming
def GetColumnasPD_LNA(arg=''):
	columnas = []
	columnas.append('cdar')
	columnas.append('deno')
	columnas.append('unidades')
	columnas.append('unidades_servidas')
	columnas.append('precio_venta')
	columnas.append('descuento')
	columnas.append('descuento_2')
	columnas.append('descuento_3')
	columnas.append('descuento_4')
	if arg == 'v':
		columnas.append('stock')
		columnas.append('neto')
		columnas.append('importe')
	columnas.append('precio_coste')
	columnas.append('estado_linea')
	if arg == 'v':
		columnas.append('deno_estado')
	columnas.append('iva')
	columnas.append('partida')
	columnas.append('capitulo')
	columnas.append('referencia_capitulo')
	columnas.append('subcapitulo')
	columnas.append('referencia_subcapitulo')
	columnas.append('descompuesto')
	columnas.append('descripcion_extendida_articulo')
	columnas.append('descripcion_extendida_capitulo')
	columnas.append('descripcion_extendida_subcapitulo')
	columnas.append('info')
	columnas.append('id_descompuesto')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def GetColumnasTIE_LNA(arg=''):
	columnas = []
	columnas.append('tecnico')
	if arg in ['v', 'a']:
		columnas.append('deno')
	if arg in ['a']:
		columnas.append('fecha')
	columnas.append('hora_entrada')
	columnas.append('hora_salida')
	columnas.append('numero_horas')

	columnas.append('parte')
	columnas.append('accion')
	columnas.append('departamento')
	if arg in ['v', 'a']:
		columnas.append('deno_departamento')

	columnas.append('rol_salarial')
	if arg in ['v', 'a']:
		columnas.append('deno_rol')
	columnas.append('horas_normales')
	columnas.append('coste_hora_normal')
	columnas.append('horas_extra')
	columnas.append('coste_hora_extra')
	columnas.append('horas_festivos')
	columnas.append('coste_hora_festivo')
	columnas.append('horas_nocturnas')
	columnas.append('coste_hora_nocturno')
	columnas.append('horas_fin_semana')
	columnas.append('coste_hora_fin_sem')
	columnas.append('usuario_sync')
	columnas.append('fecha_sync')
	columnas.append('hora_sync')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc

# noinspection PyPep8Naming
def GetColumnasDIE_LNA(arg=''):
	columnas = []
	columnas.append('cd_tecnico')
	if arg == 'v':
		columnas.append('deno')
	columnas.append('kilometros')
	columnas.append('tiempo_desplazamiento')
	columnas.append('dietas')
	columnas.append('otros_gastos')
	columnas.append('parte')
	columnas.append('accion')
	columnas.append('departamento')
	if arg == 'v':
		columnas.append('deno_departamento')
	columnas.append('usuario_sync')
	columnas.append('fecha_sync')
	columnas.append('hora_sync')
	columnas.append('id_linea')

	dc = {}
	for k in range(len(columnas)):
		dc[columnas[k]] = k
	return dc


# noinspection PyPep8Naming
def LineaToDc(linea,columnas):
	dc={}
	for key,posicion in columnas.items():
		dc[key]=copia_rg(linea[posicion])
	return dc
##

# noinspection PyPep8Naming
def Serie(serie, fecha, cl=cl, gpx=gpx):
	rg = lee_dc(cl, gpx, 'series', serie, rels='')
	if rg == 1:
		error(cl, "No existe la serie '%s'." % serie)

	if rg["SER_POST"] == 'A':
		serie += Num_aFecha(fecha, 'as')[-2:]
	elif rg["SER_POST"] == 'M':
		serie += Num_aFecha(fecha, 'ms')
	return serie

# noinspection PyPep8Naming
def SerieCertificado(tipo_contrato='', parametros=None, cl=cl, gpx=gpx):
	if parametros is None:
		parametros = lee(cl, gpx, 'parametros', '0')

	## La segunda que mas manda es la serie del tipo de contrato
	## Solo se utiliza en los partes y contratos?
	if tipo_contrato:
		if type(tipo_contrato) == str:
			rg_tipo_contrato = lee(cl, gpx, 'contratos_clases', tipo_contrato)
			if rg_tipo_contrato == 1:
				error(cl, "No existe el tipo de contrato '%s'." % tipo_contrato)

		else:
			rg_tipo_contrato = tipo_contrato
		if rg_tipo_contrato["CLC_SERC"]:
			return rg_tipo_contrato["CLC_SERC"]

	## Si no hay ninguna de las anteriores se busca la serie del documento
	series = parametros["P_SER"]
	if series == []:
		error(cl, 'Dede definir las series en parametros de la aplicación.')

	archivo = 'partes_certificados'
	tipos_ = selec(gpx, 'variables', 'V_CLAVE V_SFILTRO', preguntas=[['V_FILTRO', '==', 'DOCUMENTOS'],
	                                                                 ['V_SFILTRO', '==', archivo]])
	tipos = []
	for tipo in tipos_:
		if tipo[1] == archivo:
			tipos.append(tipo[0])

	if tipos == []:
		error(cl, "No definido en variables el archivo: '%s'." % archivo)
	if len(tipos) > 1:
		error(cl, "Hay definidas varias variables del archivo: '%s'.\12Revise solo puede haber una." % archivo)
	tipo = tipos[0]

	serie = Trae_Fila(series, tipo, clb=0, clr=1)
	## Si no hay serie definida para el documento se devuelve la serie por defecto
	if serie == None:
		serie = Trae_Fila(series, 'DEFECTO', clb=0, clr=1)
	if serie == None:
		error(cl, "Defina como mínimo la serie DEFECTO.")
	return serie


##
