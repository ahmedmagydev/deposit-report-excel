from odoo import fields, models, api


class SalikGates(models.Model):
    _name = 'salik.gates'

    name = fields.Char()
