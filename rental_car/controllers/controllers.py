# -*- coding: utf-8 -*-
from odoo import http
from odoo.http import request
from odoo import fields, models, api,_



class RentalCar(http.Controller):
    @http.route('/rental_car/rental_car', auth='public', type='json')
    def check_car(self, **kw):

        if kw['car'] and kw['gate']:
            fleet = request.env['fleet.vehicle'].sudo().search([('salik_tag_number', '=', kw['car']),], limit=1)
            get_gate = request.env['salik.gates'].sudo().search([('id', '=', kw['gate'])])
            print("fleet", fleet)
            print("get_gate", get_gate)
            if get_gate and fleet:
                create_transaction = request.env['salik.car.transaction'].sudo().create({
                    "name": "/",
                    "transaction_date": fields.Datetime.now(),
                    "gate":get_gate.id ,
                    "tag_number": fleet.salik_tag_number,
                    "plate": fleet.license_plate,
                    "amount": 4,
                    "fleet_id": fleet.id,
                })

            ############################ ok ###############################
            ##################################################################

                found_car = request.env['product.template'].sudo().search([
                    ('related_fleet', '=', fleet.id)
                ])
                print("found_car",found_car)

            ##################################################################
                get_invoice = request.env['account.move.line'].sudo().search([
                    ('product_id', '=', found_car.id),
                    ('move_id.state', '=', 'draft')
                ])

                print("invoice line", get_invoice)

                get_salik_product1 = request.env['account.move.line'].sudo().search([
                    ('product_id', '=', 4),
                    ('move_id', '=', get_invoice.move_id.id)
                ])
                get_salik_product2 = request.env['account.move.line'].sudo().search([
                ('product_id', '=', 5),
                ('move_id', '=', get_invoice.move_id.id)])

                if get_invoice and not get_salik_product1 and not get_salik_product2:
                    request.env['account.move.line'].sudo().create({
                        'product_id': 4,
                        'quantity': 1,
                        'move_id': get_invoice.move_id.id,
                    })
                    request.env['account.move.line'].sudo().create({
                        'product_id': 5,
                        'quantity': 1,
                        'move_id': get_invoice.move_id.id,
                    })

                    data = {
                        'status': 200,
                        'message ': 'Thanks,Car Transaction Recieved Successfully',
                    }
                    return data

                elif get_invoice and get_salik_product1 and get_salik_product2:

                    get_salik_product1.sudo().write({'quantity':get_salik_product1.quantity+1})
                    get_salik_product2.sudo().write({'quantity':get_salik_product2.quantity+1})
                    data = {
                        'status': 200,
                        'message ': 'Thanks,Car Transaction Recieved Successfully and update quantity',
                    }
                    return data
                
                else:
                    data = {
                        'status': 200,
                        'message ': 'No open Invoice for this car to record transition phase',
                    }
                    return data

            else:

                data = {
                    'status': 200,
                    'message ': 'No registered car or gate for data you enter',
                    
                }
                return data
        else:
            data = {
                'status': 200,
                'message ': 'you must Enter Car Salik Tag Number and Gate Number',
                    }
            return data

##################################################################

