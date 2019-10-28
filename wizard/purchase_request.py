import base64
import os
import datetime
from datetime import datetime
from docxtpl import DocxTemplate
from odoo import api, fields, models
from .PrintJob import print_job


def get_sample():
    return {
        'name': 12345,
        'date_start': '10-10-2019',
        'company': 'BLG',
        'net_amount_total': 96000000,
        'note': 'penggantian dan service rutin',
        'items': [
            {
                'product_name': 'ABCD-01',
                'description': 'Service Ganti Oli',
                'product_qty': 30,
                'product_uom': 'Unit',
                'discount': 10,
                'price_unit': 50000,
                'net_price_subtotal': 1589000
            },
            {
                'product_name': 'ABCD-02',
                'description': 'Spare part',
                'product_qty': 50,
                'product_uom': 'Unit',
                'discount': 10,
                'price_unit': 80000,
                'net_price_subtotal': 1544000
            },
            {
                'product_name': 'ABCD-03',
                'description': 'Service saja',
                'product_qty': 60,
                'product_uom': 'Unit',
                'discount': 15,
                'price_unit': 588000,
                'net_price_subtotal': 17889000
            }
        ]
    }


class PurchaseRequestReportOut(models.Model):
    _name = 'purchase.request.report.docx'
    _description = 'purchase request report'

    purchase_request_data = fields.Char('Name', size=256)
    file_name = fields.Binary('Print Purchase Request Report docx', readonly=True)


class WizardPurchaseRequest(models.Model):
    _name = 'wizard.purchase.request.print2'
    _description = "purchase request print wizard"

    @api.multi
    def get_data(self):
        self.ensure_one()
        order = self.env['purchase.request'].browse(self._context.get('active_ids', list()))
        rwcount = 1
        for rec in order:
            items = []
            for line in rec.line_ids:
                ++rwcount
                row = {'row_cnt': rwcount,
                       'product_name': str(line.product_id.name),
                       'description': str(line.name),
                       'product_qty':  line.product_qty,
                       'qty_buffer': line.qty_buffer,
                       'qty_usage_last_month1': line.qty_usage_last_month1,
                       'qty_usage_last_month2': line.qty_usage_last_month2,
                       'qty_usage_last_month3': line.qty_usage_last_month3,
                       'qty_available': line.qty_available,
                       'product_qty': line.product_qty,
                       'qty_avg_usage': line.qty_avg_usage,
                       'estimated_cost1': line.estimated_cost1,
                       'estimated_cost2': line.estimated_cost2,
                       'product_uom': str(line.product_uom_id.name),
                       'discount':  line.discount,
                       'price_unit':  line.price_unit,
                       'net_price_subtotal':  line.net_price_subtotal}
                items.append(row)
            data = {
                'document_number': str(rec.name),
                'doc_type': str(rec.doc_type),
                'date_start': str(rec.date_start),
                'company': (rec.company_id.name),
                'net_amount_total': str(rec.net_amount_total),
                'note': rec.description,
                'items': items
            }
        return data

    @api.multi
    def print_report(self):
        self.ensure_one()
        datadir = os.path.dirname(__file__)
        order = self.env['purchase.request'].browse(self._context.get('active_ids', list()))
        doctype = order.doc_type

        if doctype == 'K4':
            f = os.path.join(datadir, 'templates\purchase_request_k4.docx')
        elif doctype == 'K3':
            f = os.path.join(datadir, 'templates\purchase_request_k3.docx')
        elif doctype == 'K2':
            f = os.path.join(datadir, 'templates\purchase_request_k2.docx')
        elif doctype == 'K1':
            f = os.path.join(datadir, 'templates\purchase_request_k1.docx')
        else:
            f = os.path.join(datadir, 'templates\purchase_request_bppb.docx')

        template = DocxTemplate(f)
        context = self.get_data()
        # context = get_sample()
        template.render(context)
        filename = ('PurchaseRequestRep-' + str(datetime.today().date()) + '.docx')
        template.save(filename)
        fp = open(filename, "rb")
        file_data = fp.read()
        out = base64.encodestring(file_data)

        attach_vals = {
            'purchase_request_data': filename,
            'file_name': out,
        }

        act_id = self.env['purchase.request.report.docx'].create(attach_vals)
        fp.close()

        # print_job(filename) --> print to default printer

        return {
            'type': 'ir.actions.act_window',
            'res_model': 'purchase.request.report.docx',
            'res_id': act_id.id,
            'view_type': 'form',
            'view_mode': 'form',
            'context': self.env.context,
            'target': 'new',
        }