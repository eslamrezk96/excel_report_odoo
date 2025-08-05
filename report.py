# -*- coding: utf-8 -*-

from odoo import models, fields
import io
import xlsxwriter
import base64


class ReportWizard(models.TransientModel):
    _name = 'report.wizard'
    _description = 'Report Wizard'

    excel_file = fields.Binary('Excel File', readonly=True)

    def action_report_excel(self):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Report')
        
        """ Write code here """
        
        workbook.close()
        output.seek(0)
        self.excel_file = base64.b64encode(output.getvalue())
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model=report.wizard&id={self.id}&field=excel_file&download=true&filename=Report.xlsx',
            'target': 'self',
        }
