import sys
import os
import glob
import csv
import string
import xlwt
import xlrd
import datetime
from xlutils.copy import copy
from optparse import OptionParser
from xml.dom import minidom


class XmlCfdi(object):

	cfdi_list = []


	def __init__(self, xml_file):
		""" Initialize instance. """
		self.xml_file = xml_file


	def get_cfdi_data(self, action):
		cfdi = {}
		xml = minidom.parse(self.xml_file)
		nodos = xml.childNodes
		comprobante = nodos[0]
		comp_atributos = dict(comprobante.attributes.items())
		
		xml_file = str(self.xml_file)
		if "/" in xml_file:
			xml_file = str(self.xml_file).split("/")[2]

		try:
			version = float(comp_atributos['version']) if 'version' in comp_atributos else float(comp_atributos['Version'])
			descuento = None
			subtotal = float(comp_atributos['subTotal']) if 'subTotal' in comp_atributos else float(comp_atributos['SubTotal'])
			total = float(comp_atributos['total']) if 'total' in comp_atributos else float(comp_atributos['Total'])
			
			if 'descuento' in comp_atributos:
				descuento = float(comp_atributos['descuento'])
			elif 'Descuento' in comp_atributos:
				descuento = float(comp_atributos['Descuento'])

			if 'serie' in comp_atributos:
				serie = comp_atributos['serie']
			elif 'Serie' in comp_atributos:
				serie = comp_atributos['Serie']
			else:
				serie = ''

			fecha = comp_atributos['fecha'] if 'fecha' in comp_atributos else comp_atributos['Fecha']
			if 'T' in fecha:
				fecha = fecha.split('T')[0]

			try:
				folio = comp_atributos['folio'] if 'folio' in comp_atributos else comp_atributos['Folio']
			except:
				print 'El archivo %s no tiene folio' % xml_file
				folio = 'sin folio'

			# Voy a calcular los impuestos
			total_iva = 0.0
			if version == 3.2:
				impuestos = comprobante.getElementsByTagName('cfdi:Impuestos')[0]
				try:
					if impuestos.hasAttribute('totalImpuestosTrasladados'):
						total_iva = float(impuestos.getAttribute('totalImpuestosTrasladados'))
					else:
						print 'El comprobante %s no tiene IVA' % xml_file
				except Exception as e:
					print '%s' % str(e)
			elif version == 3.3:
				impuestos = comprobante.getElementsByTagName('cfdi:Impuestos')
				try:
					for item in impuestos:
						if item.hasAttribute('TotalImpuestosTrasladados'):
							total_iva = float(item.getAttribute('TotalImpuestosTrasladados'))
				except Exception as e:
					print '%s' % str(e)


			if descuento:
				subtotal_sin_desc = round(subtotal - descuento, 2)
			else: 
				subtotal_sin_desc = subtotal

			subtotal_iva = round(total_iva * 100 / 16, 2) # En esta linea calculo el subtotal en base al total del iva declarado
			diferencia_subs = round(subtotal_sin_desc - subtotal_iva, 2) # En esta linea calculo la diferencia entre el subtotal calculado y el declarado

			# Ahora compruebo si el subtotal_iva coincide con el subtotal declarado
			if diferencia_subs < .10:
				gastos16 = subtotal_sin_desc
				gastos0 = 0.0
			else:
				gastos16 = subtotal_iva
				gastos0 = subtotal_sin_desc - subtotal_iva

			# Obtengo los datos de la empresa con la que se trabaja, emisor o receptor dependiendo de la accion requerida
			if action == 'egresos' or action == 'acumulado-proveedores':
				tag_name_empresa = 'cfdi:Emisor'
				tag_domicilio = 'cfdi:DomicilioFiscal'
			elif action == 'ingresos' or action == 'acumulado-clientes':
				tag_name_empresa = 'cfdi:Receptor'
				tag_domicilio = 'cfdi:Domicilio'

			tag_empresa = comprobante.getElementsByTagName(tag_name_empresa)[0]
			empresa = tag_empresa.getAttribute('nombre') if tag_empresa.hasAttribute('nombre') else tag_empresa.getAttribute('Nombre')
			rfc = tag_empresa.getAttribute('rfc') if tag_empresa.hasAttribute('rfc') else tag_empresa.getAttribute('Rfc')

			calle = ''
			cp = ''
			municipio = ''
			estado = ''
			pais = ''
			try:
				tag_domicilio = comprobante.getElementsByTagName(tag_domicilio)[0]
				calle = tag_domicilio.getAttribute('calle')
				cp = tag_domicilio.getAttribute('codigoPostal')
				municipio = tag_domicilio.getAttribute('municipio')
				estado = tag_domicilio.getAttribute('estado')
				pais = tag_domicilio.getAttribute('pais')
			except:
				print 'El archivo %s no tiene domicilio' % self.xml_file
				pass

			no_exterior = ''
			colonia = ''
			localidad = ''
			try:
				no_exterior = tag_domicilio.getAttribute('noExterior')
				colonia = tag_domicilio.getAttribute('colonia')
				localidad = tag_domicilio.getAttribute('localidad')
			except:
				pass

			conceptos_list = []
			try:
				conceptos = comprobante.getElementsByTagName('cfdi:Concepto')
				for item in conceptos:
					conceptos_list.append(item.getAttribute('descripcion'))
			except:
				pass

			timbre_fiscal = comprobante.getElementsByTagName('tfd:TimbreFiscalDigital')[0]
			try:
				uuid = timbre_fiscal.getAttribute('UUID')
			except:
				print 'El xml %s no tiene UUID, verificar que este correcto en el validador del SAT' % xml_file
				uuid = None

			cfdi['serie'] = serie
			cfdi['fecha'] = fecha
			cfdi['descuento'] = descuento
			cfdi['subtotal'] = subtotal
			cfdi['total'] = total
			cfdi['file'] = xml_file
			cfdi['folio'] = folio
			cfdi['total_iva'] = total_iva
			cfdi['gastos16'] = gastos16
			cfdi['gastos0'] = gastos0
			cfdi['empresa'] = empresa
			cfdi['calle'] = calle
			cfdi['cp'] = cp
			cfdi['municipio'] = municipio
			cfdi['estado'] = estado
			cfdi['pais'] = pais
			cfdi['no_exterior'] = no_exterior
			cfdi['colonia'] = colonia
			cfdi['localidad'] = localidad
			cfdi['rfc'] = rfc
			cfdi['conceptos'] = conceptos_list
			cfdi['uuid'] = uuid

			if not self.is_duplicated(uuid, self.cfdi_list):
				self.cfdi_list.append(cfdi)
			else:
				duplicado = {'file': xml_file}
				print 'El archivo %s esta duplicado, puedes borrarlo!' % xml_file
		except Exception as e:
			exc_type, exc_obj, exc_tb = sys.exc_info()
			fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
			print '%s, %s, %s' % (exc_type, fname, exc_tb.tb_lineno)
			print 'Error: %s' % str(e)
			print 'El archivo %s tiene un error' % xml_file
			print '----------------------------------------'

	def clean_acumulado(self, ws):
		row_initial = 3
		row_final = len(ws.rows)
		for row in xrange(row_initial, row_final):
			for cell in xrange(0, 20):
				ws.write(row, cell, '')


	def generate_cfdi_on_xls(self, action):
		cfdi_list = self.cfdi_list
		bk = book()
		date_format = xlwt.easyxf('', num_format_str='YYYY-MM-DD')
		currency_format = xlwt.easyxf('', '#,###.00')
		wb = bk.get_workbook()
		cfdi_list.sort(key=lambda k: k['fecha']) # Ordeno todos los diccionarios por fecha ascendente
		start = 4
		
		if action == 'acumulado-proveedores' or action == 'acumulado-clientes':
			empresas_dict = {}
			if action == 'acumulado-proveedores':
				ws = wb.get_sheet(2)
				leyenda_total = 'Gran total de Egresos'
			elif action == 'acumulado-clientes':
				ws = wb.get_sheet(3)
				leyenda_total = 'Gran total de Ingresos'
			self.clean_acumulado(ws)
			for i in xrange(0, len(cfdi_list)):
				empresa = cfdi_list[i]['rfc']
				if not empresa in empresas_dict:
					empresas_dict[empresa] = { 
						1: 0.0,
						2: 0.0,
						3: 0.0,
						4: 0.0,
						5: 0.0,
						6: 0.0,
						7: 0.0,
						8: 0.0,
						9: 0.0,
						10: 0.0,
						11: 0.0,
						12: 0.0,
						'calle': cfdi_list[i]['calle'],
						'cp': cfdi_list[i]['cp'],
						'municipio': cfdi_list[i]['municipio'],
						'estado': cfdi_list[i]['estado'],
						'pais': cfdi_list[i]['pais'],
						'no_exterior': cfdi_list[i]['no_exterior'],
						'colonia': cfdi_list[i]['colonia'],
						'localidad': cfdi_list[i]['localidad'],
						'rfc': cfdi_list[i]['rfc'],
						'empresa': cfdi_list[i]['empresa']
					}

				fecha = cfdi_list[i]['fecha']
				mes = int(fecha.split('-')[1])
				empresas_dict[empresa][mes] += float(cfdi_list[i]['total'])

			row = start
			ind = 1
			for empresa, totales in sorted(empresas_dict.iteritems()):
				ws.write(row, 0, ind)
				ws.write(row, 1, totales['empresa'])
				for i in xrange(1, 13):
					if totales[i] > 0.0:
						ws.write(row, i+1, totales[i], currency_format)
				totalxempresa = "SUM(C%i:N%i)" % (row+1, row+1)
				ws.write(row, 14, xlwt.Formula(totalxempresa), currency_format)
				ws.write(row, 15, totales['rfc'])
				ws.write(row, 16, totales['calle'] + ' ' + empresas_dict[empresa]['no_exterior'] + ' ' + empresas_dict[empresa]['colonia'] + ' ' + empresas_dict[empresa]['localidad'])
				ws.write(row, 17, totales['cp'])
				ws.write(row, 18, totales['municipio'])
				ws.write(row, 19, totales['estado'])
				row += 1
				ind += 1

			row += 1
			rango = list(string.ascii_uppercase)[2:15]
			ws.write(row, 1, leyenda_total)
			for i in xrange(1, 14):
				col = rango[i-1]
				totalxmes = "SUM(%s%i:%s%i)" % (col, start+1, col, row-1)
				ws.write(row, i+1, xlwt.Formula(totalxmes), currency_format)


		if action == 'egresos':
			ws = wb.get_sheet(1)
			last_row = len(ws.rows) + 1
			for i in xrange(0, len(cfdi_list)):
				ind = last_row + i
				ws.write(ind, 0, cfdi_list[i]['serie'] + cfdi_list[i]['folio'])
				ws.write(ind, 1, cfdi_list[i]['fecha'], date_format)
				ws.write(ind, 2, cfdi_list[i]['empresa'])
				ws.write(ind, 3, cfdi_list[i]['gastos16'], currency_format)
				ws.write(ind, 4, cfdi_list[i]['gastos0'], currency_format)
				ws.write(ind, 5, cfdi_list[i]['total_iva'], currency_format)
				ws.write(ind, 6, cfdi_list[i]['total'], currency_format)
				ws.write(ind, 7, cfdi_list[i]['file'])

		if action == 'ingresos':
			ws = wb.get_sheet(0)
			last_row = len(ws.rows) + 1
			for i in xrange(0, len(cfdi_list)):
				ind = last_row + i
				ws.write(ind, 0, cfdi_list[i]['empresa'])
				ws.write(ind, 1, cfdi_list[i]['rfc'])
				ws.write(ind, 2, cfdi_list[i]['fecha'], date_format)
				ws.write(ind, 3, cfdi_list[i]['serie'] + cfdi_list[i]['folio'])
				ws.write(ind, 4, cfdi_list[i]['conceptos'])
				ws.write(ind, 5, cfdi_list[i]['gastos16'], currency_format)
				ws.write(ind, 6, cfdi_list[i]['total_iva'], currency_format)
				ws.write(ind, 7, cfdi_list[i]['total'], currency_format)

		workbook_name = bk.get_workbook_name()
		os.remove(workbook_name) # Borro el libro del cual cree la copia, para poder guardar los cambios con el mismo nombre
		wb.save(workbook_name)
		print 'Listo!'


	def is_duplicated(self, uuid, ls):
		for item in ls:
			if item['uuid'] == uuid:
				return True
		return False


class book():
	# Esta funcion devolvera una copia del libro contabilidad_ANIO_ACTUAL.xls o bien una copia del libro patron.xls con el nombre de contabilidad_ANIO_ACTUAL.xls
	def get_workbook(self):
		path = self.get_workbook_name()
		if os.path.isfile(path):
			wb = xlrd.open_workbook(path, formatting_info=True)
			workbook = copy(wb)
		else:
			patron = xlrd.open_workbook('patron.xls', formatting_info=True)
			workbook = copy(patron)
			workbook.save(path)
		return workbook

	def get_workbook_name(self):
		hoy = datetime.date.today()
		year = hoy.year
		path = 'contabilidad_%i.xls' % year
		return path


def main(argv):

	usage = "%prog [opciones] archivocfd.xml|*.xml"
	# usage = "%prog [opciones] archivocfd.xml|*.xml"
	add_help_option = False
	parser = OptionParser(usage=usage, add_help_option=add_help_option)

	parser.add_option("-e", "--egresos", dest="generaEgresos", 
		help=u"Genera la lista de egresos del mes que se pasa como argumento")

	parser.add_option("-h", "--help", action="help",
		help=u"Ejecutar generate_cfdi_xls.py con archivo_cfdi.xml accion o folder_mes/*.xml accion como parametro")

	
	(options, args) = parser.parse_args()

	actions = ['egresos', 'ingresos', 'acumulado-proveedores', 'acumulado-clientes']

	if len(args) <= 1:
	    parser.print_help()
	    sys.exit(0)

	action = args[0]
	
	if not action in actions:
		print 'Debes pasar una de las siguientes acciones %s' % actions
		print 'Pasaste %s' % action
		sys.exit(0)

	# Se obtiene la lista de archivos
	if len(args) == 1 and "*" not in args[1:2]:
		files = args
	elif len(args) == 1 and "*" in args[1:2]:
		files = glob.glob(args[1:2])
	else:
		files = args[1:]
		print 'Archivos XML a procesar: %i' % len(files)

	for item in files:
		xml_file = item
		if not os.path.isfile(xml_file):
			print "El archivo " + xml_file + " no existe."
		else:
			xmlcfdi = XmlCfdi(xml_file)
			xmlcfdi.get_cfdi_data(action)

	print 'Generando %s ...' % action
	xmlcfdi.generate_cfdi_on_xls(action)


if __name__ == "__main__":
	main(sys.argv[1:])