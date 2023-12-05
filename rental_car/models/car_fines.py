from odoo import fields, models, api,_


class CarFines(models.Model):
    _name = 'car.fines'

    name = fields.Char(string='Name')
    plate_number = fields.Char(string='Plate Number')
    plate_category = fields.Char(string='Plate Category')
    plate_code = fields.Char(string='Plate Code')
    license_number = fields.Char(string='License Number')
    license_from = fields.Char(string='License From')
    ticket_number = fields.Char(string='Ticket Number',required=False)
    ticket_date = fields.Char(string='Ticket Date')
    ticket_time = fields.Char(string='Ticket Time')
    fines_source = fields.Char(string='Fines source')
    ticket_fees = fields.Float(string='Ticket Fee', required=False)
    ticket_status = fields.Char(string='Ticket Status')
    fleet_id = fields.Many2one(
        comodel_name='fleet.vehicle',
        string='Fleet',
        required=False)
    state = fields.Selection(
        selection=[
            ('un_allocated', 'Un Allocated'),
            ('allocated', 'Allocated'),
        ],
        string='Status',)
    Paid_state = fields.Selection(
        selection=[
            ('un_paid', 'Un Paid'),
            ('company_paid', 'Company Paid'),
            ('customer_paid', 'Customer Paid'),
        ],
        string='Payment Status', )
    sale_order_id = fields.Many2one(
        comodel_name='sale.order',
        string='Order',
        required=False)


    @api.model_create_multi
    def create(self, vals_list):
        # We generate a standard reference
        for vals in vals_list:
            vals['name'] = self.env['ir.sequence'].next_by_code('car.fines') or '/'
        return super().create(vals_list)

