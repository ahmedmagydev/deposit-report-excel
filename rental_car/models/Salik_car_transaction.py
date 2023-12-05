from odoo import fields, models, api,_


class CarTransaction(models.Model):
    _name = 'salik.car.transaction'

    name = fields.Char(string='Transaction ID')
    transaction_id = fields.Char(string='Transaction ID')


    transaction_date = fields.Datetime(
        string='Transaction Date',
        required=False)
    gate = fields.Many2one(
        comodel_name='salik.gates',
        string='Gate',
        required=False)
    tag_number = fields.Char(string='Tag Number')
    plate = fields.Char(string='Plate')

    amount = fields.Float(
        string='Amount',
        required=False)
    fleet_id = fields.Many2one(
        comodel_name='fleet.vehicle',
        string='Fleet',
        required=False)
    sale_order_id = fields.Many2one(
        comodel_name='sale.order',
        string='Order',
        required=False)

    state = fields.Selection(
        selection=[
            ('un_allocated', 'Un Allocated'),
            ('allocated', 'Allocated'),
        ],
        string='Status', )


    @api.model_create_multi
    def create(self, vals_list):
        # We generate a standard reference
        for vals in vals_list:
            vals['name'] = self.env['ir.sequence'].next_by_code('salik.car.transaction') or '/'
        return super().create(vals_list)

