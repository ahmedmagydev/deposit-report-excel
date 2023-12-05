from odoo import fields, models, api


class RentalProcessing(models.TransientModel):
    _inherit = 'rental.order.wizard'

class RentalProcessingLine(models.TransientModel):
        _inherit = 'rental.order.wizard.line'

        @api.model
        def _default_wizard_line_vals(self, line, status):
            delay_price = line.product_id._compute_delay_price(fields.Datetime.now() - line.return_date)
            return {
                'order_line_id': line.id,
                'product_id': line.product_id.id,
                'qty_reserved': line.product_uom_qty,
                'qty_delivered': 1,
                'qty_returned': line.qty_returned if status == 'pickup' else line.qty_delivered - line.qty_returned,
                'is_late': line.is_late and delay_price > 0
            }
#
#     # def apply(self):
#     #     super().apply()
#     #     # self.order_id.onchange_method_rental_status()


