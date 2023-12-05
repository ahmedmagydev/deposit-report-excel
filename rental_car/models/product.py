from odoo import fields, models, api,_
from odoo.exceptions import ValidationError, AccessError


class ProductTemplate(models.Model):

    _inherit = 'product.template'
    is_fleet = fields.Boolean(
        string='Is Fleet',
        required=False)

    fleet_models= fields.Many2one(
        comodel_name='fleet.vehicle.model',
        string='Fleet Model',
        required=False)
    related_fleet= fields.Many2one(
        comodel_name='fleet.vehicle',
        string='Fleet',
        required=False, store=True)

    transaction_ids = fields.Many2many(
        comodel_name='salik.car.transaction',
        string='all car Transaction', compute="calculate_all_related_transaction",
        required=False)
    transaction_count = fields.Integer(compute="_compute_count_transaction", string='transaction count')


    @api.onchange('is_fleet')
    def onchange_method(self):

       if self.is_fleet:
           if self.fleet_models:

                   related_fleet = self.env['fleet.vehicle'].sudo().create({
                       'model_id': self.fleet_models.id,
                   })

                   self.related_fleet=related_fleet.id

           # else:
           #     raise ValidationError(_('You must select model'))


    def _compute_count_transaction(self):
        transaction = self.env['salik.car.transaction'].sudo().search([('fleet_id', '=', self.related_fleet.id)])

        self.transaction_count = len(transaction)

    def calculate_all_related_transaction(self):

        all_transaction = self.env['salik.car.transaction'].search([
            ("fleet_id", "=", self.related_fleet.id)
        ])
        if all_transaction:
            self.transaction_ids = all_transaction.ids
        else:
            self.transaction_ids = False

    def action_view_transaction(self):
        self.ensure_one()
        action = self.env["ir.actions.actions"].sudo()._for_xml_id(
            "rental_car.ModelName_act_window_salik_car_transaction")
        action['domain'] = [('id', 'in', self.transaction_ids.ids)]
        action['context'] = dict(self._context, create=False)
        return action


    def action_view_fleet(self):
        self.ensure_one()
        # view = self.env.ref("fleet.fleet_vehicle_view_form")

        view_form_id = self.env.ref("fleet.fleet_vehicle_view_form").id
        # moves = self.mapped('move_ids.id')
        context = self._context
        # view_id = self.sudo().env['stock.view_picking_form']
        if self.related_fleet:
            return {
                'type': 'ir.actions.act_window',
                'name': 'View Picking',
                'res_model': 'fleet.vehicle',
                'view_type': 'form',
                'view_mode': 'form',
                'target': 'current',
                'res_id': self.sudo().related_fleet.id,
                # 'context': {'default_product_tmpl_id': self.id},
            }



