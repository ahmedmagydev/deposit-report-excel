# -*- coding: utf-8 -*-
from odoo import http
from odoo.http import request


class RentalCar(http.Controller):
    @http.route('/rental_car/get_cars', auth='public', type='json')
    def check_car(self, **kw):
        cars = request.env['fleet.vehicle'].sudo().search([])
        if cars:
            cars_ids = []
            for rec in cars:
                vals = {
                    'id': rec.salik_tag_number
                }
                cars_ids.append(vals)
            data = {
                'status': 200,
                'car salik tag number ': cars_ids,
                    }
            return data
        else:
            data = {'status': 200,
                    'message':'No any Car Found'

                    }
            return data

