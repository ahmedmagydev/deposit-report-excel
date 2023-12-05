from odoo import fields, models, api, _
from odoo.exceptions import ValidationError, AccessError


class SaleOrderLine(models.Model):
    _inherit = "sale.order.line"

    # @api.onchange('product_id')
    # def onchange_method(self):
    #     for rec in self:
    #         if rec.product_id:
    #             print(rec.order_id.id.origin)
    #             print(rec.order_id.id)
    #             sales_order_line = self.env['sale.order.line'].sudo().search([
    #                 # ('order_id', '!=', rec.order_id.id),
    #                 ('product_id', '=', rec.product_id.id),
    #                 ('state', '=', 'draft')
    #             ])
    #             print(sales_order_line)
    #             if sales_order_line:
    #                 raise ValidationError(_('You Can not add car(%s) to new rental order beacuse\n '
    #                                         'it already rent in order number : (%s).')
    #                                       % (rec.product_id.name, sales_order_line.order_id.name))

    def send_finish_rent_mail(self):
        template_id = self.env.ref('rental_car.email_finished_rent_date').id
        template = self.env['mail.template'].browse(template_id)
        template.send_mail(self.id, force_send=True)
