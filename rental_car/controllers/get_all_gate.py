# -*- coding: utf-8 -*-
from odoo import http
from odoo.http import request


class RentalCar(http.Controller):
    @http.route('/rental_car/get_salik_gates', auth='public', type='json')
    def check_car(self, **kw):
        salik_gates = request.env['salik.gates'].sudo().search([])
        if salik_gates:
            gates = []
            for rec in salik_gates:
                vals = {
                    'id': rec.id,
                    'name': rec.name
                }
                gates.append(vals)
            data = {
                'status': 200,
                'car salik tag number ': gates,

                    }
            return data
        else:
            data = {'status': 200,
                    'message':'No gates Created Yet'

                    }
            return data

