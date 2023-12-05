from odoo import fields, models, api, _
from odoo.exceptions import ValidationError, AccessError


class AccountMove(models.Model):
    _inherit = "account.move"

    fine_invoice = fields.Boolean(string='Fines Invoice',default=False)
    salik_invoice = fields.Boolean(string='Salik Invoice',default=False)

    def write(self, vals):
        res = super(AccountMove, self).write(vals)

        for record in self.invoice_line_ids:
           record.sale_line_ids
           print(record.sale_line_ids)
           fine_found_before = self.env['car.fines'].sudo().search([('sale_order_id', '=', record.sale_line_ids.order_id.id)])
           print(fine_found_before)

           if fine_found_before:
               fine_found_before.Paid_state = 'customer_paid'
           else:
               fine_found_before.Paid_state = fine_found_before.Paid_state


        return res


