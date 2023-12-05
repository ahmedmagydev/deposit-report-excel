from odoo import fields, models, api

class FleetVehicle(models.Model):
    _inherit = 'fleet.vehicle'

    salik_tag_number = fields.Char(
        string='Salik Tag Number',
        required=False)

    _sql_constraints = [
        ('unique_salik_tag_number',
         'UNIQUE(salik_tag_number)',
         'Salik Tag Number must be unique'),
    ]

    transaction_ids = fields.Many2many(
        comodel_name='salik.car.transaction',
        string='all car Transaction', compute="calculate_all_related_transaction",
        required=False)
    transaction_count = fields.Integer(compute="_compute_count_transaction", string='transaction count')

    def _compute_count_transaction(self):
        transaction= self.env['salik.car.transaction'].sudo().search([('fleet_id', '=', self.id)])

        self.transaction_count=len(transaction)
    def calculate_all_related_transaction(self):

        all_transaction = self.env['salik.car.transaction'].search([
            ("fleet_id", "=", self.id)
        ])
        if all_transaction:
            self.transaction_ids = all_transaction.ids
        else:
            self.transaction_ids = False

    def action_view_transaction(self):
        self.ensure_one()
        action = self.env["ir.actions.actions"].sudo()._for_xml_id("rental_car.ModelName_act_window_salik_car_transaction")
        action['domain'] = [('id', 'in', self.transaction_ids.ids)]
        action['context'] = dict(self._context, create=False)
        return action
