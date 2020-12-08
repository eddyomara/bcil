# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

import time
import tempfile
import binascii
import xlrd
import io
from odoo.tools import DEFAULT_SERVER_DATETIME_FORMAT, DEFAULT_SERVER_DATE_FORMAT
from datetime import date, datetime
from odoo.exceptions import Warning
from odoo import models, fields, exceptions, api, _

import logging
_logger = logging.getLogger(__name__)

try:
    import csv
except ImportError:
    _logger.debug('Cannot `import csv`.')
try:
    import xlwt
except ImportError:
    _logger.debug('Cannot `import xlwt`.')
try:
    import cStringIO
except ImportError:
    _logger.debug('Cannot `import cStringIO`.')
try:
    import base64
except ImportError:
    _logger.debug('Cannot `import base64`.')

class purchase_order(models.Model):
    _inherit = 'purchase.order'

    custom_seq = fields.Boolean('Custom Sequence')
    system_seq = fields.Boolean('System Sequence')
    purchase_name = fields.Char('Purchase Name')

class gen_purchase(models.TransientModel):
    _name = "gen.purchase"

    file = fields.Binary('File')
    sequence_opt = fields.Selection([('custom', 'Use Excel/CSV Sequence Number'), ('system', 'Use System Default Sequence Number')], string='Sequence Option',default='custom')
    import_option = fields.Selection([('csv', 'CSV File'), ('xls', 'XLS File')], string='Select', default='csv')
    stage = fields.Selection(
        [('draft', 'Import Draft Purchase'), ('confirm', 'Confirm Purchase Automatically With Import')],
        string="Purchase Stage Option", default='draft')
    import_prod_option = fields.Selection([('name', 'Name'),('code', 'Code'),('barcode', 'Barcode')],string='Import Product By ',default='name')        

    @api.multi
    def make_purchase(self, values):
        purchase_obj = self.env['purchase.order']
        if self.sequence_opt == "custom":
            pur_search = purchase_obj.search([
                ('name', '=', values.get('purchase_no')),
            ])
        else:
            pur_search = purchase_obj.search([
                ('purchase_name', '=', values.get('purchase_no')),
            ])
            
        if pur_search:
            if pur_search.partner_id.name == values.get('vendor'):
                if  pur_search.currency_id.name == values.get('currency'):
                    self.make_purchase_line(values, pur_search)
                    return pur_search
                else:
                    raise Warning(_('Currency is different for "%s" .\n Please define same.') % values.get('currency'))
            else:
                raise Warning(_('Customer name is different for "%s" .\n Please define same.') % values.get('vendor'))
        else:
            if values.get('seq_opt') == 'system':
                name = self.env['ir.sequence'].next_by_code('purchase.order')
            elif values.get('seq_opt') == 'custom':
                name = values.get('purchase_no')
            partner_id = self.find_partner(values.get('vendor'))
            currency_id = self.find_currency(values.get('currency'))
            if values.get('date'):
                pur_date = self.make_purchase_date(values.get('date'))
            else:
                pur_date = datetime.today()
            user_id  = self.find_user(values.get('user'))
            payment_id  = self.find_payment_term(values.get('payment_term'))
            status  = values.get('state')
            if partner_id.property_account_receivable_id:
                account_id = partner_id.property_account_payable_id
            else:
                account_search = self.env['ir.property'].search([('name', '=', 'property_account_expense_categ_id')])
                account_id = account_search.value_reference
                account_id = account_id.split(",")[1]
                account_id = self.env['account.account'].browse(account_id)
            pur_id = purchase_obj.create({
                'account_id' : account_id.id,
                'partner_id' : partner_id.id,
                'currency_id' : currency_id.id,
                'name':name,
                'date_order':pur_date,
                'custom_seq': True if values.get('seq_opt') == 'custom' else False,
                'system_seq': True if values.get('seq_opt') == 'system' else False,
                'purchase_name' : values.get('purchase_no'),
                'user_id' : user_id.id,
                'payment_term_id' : payment_id.id
            })
            if status == 'purchase':
                pur_id.button_confirm()
            elif status == '':
                pur_id.write({'state':'draft'})
            elif status == 'Purchase Order':
                pur_id.button_confirm()
            elif status == 'RFQ':
                pur_id.write({'state':'draft'})
            elif status == 'RFQ Sent':
                pur_id.write({'state':'sent'})
            elif status == 'To Approve':
                pur_id.write({'state':'to approve'})
            elif status == 'Locked':
                pur_id.write({'state':'done'})
            elif status == 'Cancelled':
                pur_id.write({'state':'cancel'})
            else:
                pur_id.write({'state':status})
        self.make_purchase_line(values, pur_id)
        return pur_id

    @api.multi
    def find_user(self, name):
        user_obj = self.env['res.users']
        user_search = user_obj.search([('name', '=', name)])
        if user_search:
            return user_search
        else:
            raise Warning(_(' "%s" User is not available.') % name)

    @api.multi
    def find_payment_term(self, name):
        payment_obj = self.env['account.payment.term']
        payment_search = payment_obj.search([('name', '=', name)])
        if payment_search:
            return payment_search
        else:
            raise Warning(_(' "%s" payment Term is not available.') % name)

    @api.multi
    def make_purchase_date(self, date):
        DATETIME_FORMAT = "%Y-%m-%d"
        i_date = datetime.strptime(date, DATETIME_FORMAT)
        return i_date

    @api.multi
    def make_purchase_sch_date(self, date):
        DATETIME_FORMAT = "%Y-%m-%d"
        i_date = datetime.strptime(date, DATETIME_FORMAT)
        return i_date

    @api.multi
    def make_purchase_line(self, values, pur_id):
        product_obj = self.env['product.product']
        account = False
        purchase_line_obj = self.env['purchase.order.line']
        current_time=datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        if self.import_prod_option == 'barcode':
          product_search = product_obj.search([('barcode',  '=',values['product'])])
        elif self.import_prod_option == 'code':
            product_search = product_obj.search([('default_code', '=',values['product'])])
        else:
            product_search = product_obj.search([('name', '=',values['product'])])

        product_uom = self.env['uom.uom'].search([('name', '=', values.get('uom'))])
        if product_uom.id == False:
            raise Warning(_(' "%s" Product UOM category is not available.') % values.get('uom'))

        if product_search:
            product_id = product_search
        else:
            if self.import_prod_option == 'name':
                product_id = product_obj.create({
                                                    'name':values.get('product'),
                                                    'lst_price':float(values.get('price')),
                                                    'uom_id':product_uom.id,
                                                    'uom_po_id':product_uom.id
                                                 })
            else:
                raise Warning(_('%s product is not found" .\n If you want to create product then first select Import Product By Name option .') % values.get('product'))

        if values.get('sch_date'):
            sch_date = self.make_purchase_sch_date(values.get('sch_date'))
        else:
            sch_date = datetime.today()
        
        if pur_id.state == 'draft':
                po_order_lines = purchase_line_obj.create({
                                                    'order_id':pur_id.id,
                                                    'product_id':product_id.id,
                                                    'name':product_id.name,
                                                    'date_planned':sch_date,
                                                    'product_qty':values.get('quantity'),
                                                    'product_uom':product_uom.id,
                                                    'price_unit':values.get('price')
                                                    })
        elif pur_id.state == 'sent':
            po_order_lines = purchase_line_obj.create({
                                                'order_id':pur_id.id,
                                                'product_id':product_id.id,
                                                'name':product_id.name,
                                                'date_planned':sch_date,
                                                'product_qty':values.get('quantity'),
                                                'product_uom':product_uom.id,
                                                'price_unit':values.get('price')
                                                })
        elif pur_id.state != 'sent' or pur_id.state != 'draft':
            raise Warning(_('We cannot import data in validated or confirmed order.')) 

        tax_ids = []
        if values.get('tax'):
            if ';' in  values.get('tax'):
                tax_names = values.get('tax').split(';')
                for name in tax_names:
                    tax= self.env['account.tax'].search([('name', '=', name),('type_tax_use','=','purchase')])
                    if not tax:
                        raise Warning(_('"%s" Tax not in your system') % name)
                    tax_ids.append(tax.id)

            elif ',' in  values.get('tax'):
                tax_names = values.get('tax').split(',')
                for name in tax_names:
                    tax= self.env['account.tax'].search([('name', '=', name),('type_tax_use','=','purchase')])
                    if not tax:
                        raise Warning(_('"%s" Tax not in your system') % name)
                    tax_ids.append(tax.id)
            else:
                tax_names = values.get('tax').split(',')
                for name in tax_names:
                    tax = self.env['account.tax'].search([('name', '=', name), ('type_tax_use', '=', 'purchase')])
                    if not tax:
                        raise Warning(_('"%s" Tax not in your system') % name)
                    tax_ids.append(tax.id)

        if tax_ids:
            po_order_lines.write({'taxes_id':([(6, 0, tax_ids)])})

        tag_ids = []
        if values.get('analytic_tag'):
            if ';' in  values.get('analytic_tag'):
                tag_names = values.get('analytic_tag').split(';')
                for name in tag_names:
                    tag= self.env['account.analytic.tag'].search([('name', '=', name)])
                    if not tag:
                        raise Warning(_('"%s" Analytic Tags not in your system') % name)
                    tag_ids.append(tag.id)

            elif ',' in  values.get('analytic_tag'):
                tag_names = values.get('analytic_tag').split(',')
                for name in tag_names:
                    tag= self.env['account.analytic.tag'].search([('name', '=', name)])
                    if not tag:
                        raise Warning(_('"%s" Analytic Tags not in your system') % name)
                    tag_ids.append(tag.id)
            else:
                tag_names = values.get('analytic_tag').split(',')
                tag= self.env['account.analytic.tag'].search([('name', '=', tag_names)])
                if not tag:
                    raise Warning(_('"%s" Analytic Tags not in your system') % tag_names)
                tag_ids.append(tag.id)
        if tag_ids:
            po_order_lines.write({'analytic_tag_ids':([(6, 0, tag_ids)])})

        if values.get('analytic_account_id'):
            analytic_account_id = self.env['account.analytic.account'].search([('name','=',values.get('analytic_account_id'))])
            if analytic_account_id:
                analytic_account_id = analytic_account_id
                po_order_lines.write({
                'account_analytic_id' : analytic_account_id.id
                })
            else:
                raise Warning(_(' "%s" Analytic Account is not available.') % values.get('analytic_account_id'))

        return True

    @api.multi
    def find_currency(self, name):
        currency_obj = self.env['res.currency']
        currency_search = currency_obj.search([('name', '=', name)])
        if currency_search:
            return currency_search
        else:
            raise Warning(_(' "%s" Currency are not available.') % name)

    @api.multi
    def find_partner(self, name):
        partner_obj = self.env['res.partner']
        partner_search = partner_obj.search([('name', '=', name)])
        if partner_search:
            return partner_search
        else:
            partner_id = partner_obj.create({
                'name' : name})
            return partner_id

    @api.multi
    def import_csv(self):
        """Load Inventory data from the CSV file."""
        if self.import_option == 'csv':
            keys = ['purchase_no', 'currency','date', 'product','analytic_account_id','analytic_tag', 'uom','sch_date','state','line_state','tax','quantity','user','vendor','price','payment_term']
            try:
                csv_data = base64.b64decode(self.file)
                data_file = io.StringIO(csv_data.decode("utf-8"))
                data_file.seek(0)
                file_reader = []
                purchase_ids = []
                csv_reader = csv.reader(data_file, delimiter=',')
                file_reader.extend(csv_reader)
            except Exception:
                raise exceptions.Warning(_("Invalid file!"))
            values = {}
            for i in range(len(file_reader)):
                #                val = {}
                field = list(map(str, file_reader[i]))
                values = dict(zip(keys, field))
                if values:
                    if i == 0:
                        continue
                    else:
                        values.update({'seq_opt':self.sequence_opt})
                        res = self.make_purchase(values)
                        purchase_ids.append(res)
                        
            if self.stage == 'confirm':
                for res in purchase_ids: 
                    if res.state in ['draft', 'sent']:
                        res.button_confirm()
        else:
            try:
                fp = tempfile.NamedTemporaryFile(delete= False,suffix=".xlsx")
                fp.write(binascii.a2b_base64(self.file))
                fp.seek(0)
                values = {}
                purchase_ids = []
                workbook = xlrd.open_workbook(fp.name)
                sheet = workbook.sheet_by_index(0)
            except Exception:
                raise exceptions.Warning(_("Invalid file!"))
            
            product_obj = self.env['product.product']
            date_string = False
            for row_no in range(sheet.nrows):
                val = {}
                tax_line = ''
                if row_no <= 0:
                    fields = map(lambda row:row.value.encode('utf-8'), sheet.row(row_no))
                else:
                    line = list(map(lambda row:isinstance(row.value, bytes) and row.value.encode('utf-8') or str(row.value), sheet.row(row_no)))
                    if line[2] != '':
                        a1 = int(float(line[2]))
                        a1_as_datetime = datetime(*xlrd.xldate_as_tuple(a1, workbook.datemode))
                        date_string = a1_as_datetime.date().strftime('%Y-%m-%d')
                    else:
                        date_string = False
                    if line[7] != '':
                        a2 = int(float(line[7]))
                        a2_as_datetime = datetime(*xlrd.xldate_as_tuple(a2, workbook.datemode))
                        date_string2 = a2_as_datetime.date().strftime('%Y-%m-%d')
                    else:
                        date_string2 = False
                    values.update({'purchase_no':line[0],
                                   'vendor': line[13],
                                   'currency': line[1],
                                   'product': line[3].split('.')[0],
                                   'analytic_account_id': line[4],
                                   'analytic_tag': line[5],
                                   'quantity': line[11],
                                   'uom': line[6],
                                   'sch_date': date_string2,
                                   'state': line[8],
                                   'line_state': line[9],
                                   'price': line[15],
                                   'tax': line[10],
                                   'user': line[12],
                                   'payment_term': line[16],
                                   'date': date_string,
                                   'seq_opt':self.sequence_opt

                                   })
                    res = self.make_purchase(values)
                    purchase_ids.append(res)
                    
            if self.stage == 'confirm':
                for res in purchase_ids: 
                    if res.state in ['draft', 'sent']:
                        res.button_confirm()
        return res

