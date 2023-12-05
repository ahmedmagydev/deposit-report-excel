from odoo import fields, models, api, _
from odoo.exceptions import UserError, ValidationError, AccessError, RedirectWarning
import io
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
# Import math Library
import math
import os
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from datetime import datetime, time
from odoo import Command


class SaleOrder(models.Model):
    _inherit = "sale.order"

    def onchange_method_rental_status(self):
        pick_sales_orders = self.env['sale.order'].sudo().search([('rental_status', '=', 'pickup'),
                                                                  ('invoice_ids', '!=', False),
                                                                  ('invoice_ids.state', '=', 'draft')])
        print("pick_sales_orders:", pick_sales_orders)
        for record in pick_sales_orders:
            if record.rental_status == 'pickup':
                if record.order_line:
                    for rec in record.order_line:
                        if rec.product_id.rent_ok and rec.product_uom_qty >= rec.qty_delivered:
                            rec.qty_delivered += 1








                        # if rec.product_uom_qty > rec.qty_delivered:
                        #     rec.send_finish_rent_mail()

                # if record.invoice_ids and record.invoice_ids.state == 'draft':
                #     for rec in record.invoice_ids.invoice_line_ids:
                #         if rec.product_id.rent_ok:
                #             rec.sudo().write({'quantity': rec.quantity+1})
                #             rec._compute_totals()
                #

    ######################################################################################################################
    #   Salik
    ######################################################################################################################

    def search_file(self):
        # SERVICE_ACCOUNT_FILE1 = os.path.realpath(os.path.join(os.path.dirname(__file__), '../../../superior-386008-1be184fb4024.json'))
        # print("kkkkkkkkkkkkkkkkkkkkkkkkkkkkkkkk",SERVICE_ACCOUNT_FILE1)

        SERVICE_ACCOUNT_FILE = '/home/odoo/src/user/rental_car/superior-386008-1be184fb4024.json'
        # SERVICE_ACCOUNT_FILE = '/home/osama/odoo/custom/superior_4/rental_car/superior-386008-1be184fb4024.json'

        SCOPES = ['https://www.googleapis.com/auth/drive']

        # Load the service account credentials from the key file
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

        # access_token = creds.token

        try:
            # create drive api client
            service = build('drive', 'v3', credentials=creds)
            files = []
            page_token = None
            while True:
                response = service.files().list(q="name='test_salik'",
                                                spaces='drive',
                                                fields='nextPageToken,''files(id,name,owners)',
                                                pageToken=page_token).execute()
                # for file in response.get('files', []):
                #     # Process change
                #     print(F'Found file: {file.get("name")}, {file.mget("id")}')
                files.extend(response.get('files', []))
                page_token = response.get('nextPageToken', None)

                for file in response.get('files', []):
                    file_id = file.get('id')
                    print(file_id)

                    # file = service.files().get(fileId=file_id).execute()
                    file_content = service.files().export(fileId=file_id,
                                                          mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').execute()
                    file_bytes = io.BytesIO(file_content)

                    # read the Excel sheet into a Pandas DataFrame
                    df = pd.read_excel(file_bytes)

                    # display the data
                    # print(df.head())
                    # print(df)
                    # print(df.dtypes)
                    # assuming 'df' is the DataFrame containing the Excel sheet data
                    for row in df.iterrows():
                        print("###################################")
                        # print(row[1][0])
                        # print(row[1][1])
                        # print(row[1][2])

                    # column_data = df['Direction']
                if page_token is None:
                    break

        except HttpError as error:
            print(F'An error occurred: {error}')
            files = None

        return df.iterrows()

    def get_salik_transaction(self):
        for row in self.search_file():


            if not math.isnan(row[1][0]):
                print("#########Tag Number############Plate Number##############")
                new_time = datetime.strptime(str(row[1][2]), '%H:%M:%S').time()
                new_datetime = datetime.combine(row[1][1], new_time)
                # salik_transction_date1 = fields.Datetime.to_string(new_datetime)
                salik_transction_date = fields.datetime.strptime(str(new_datetime), "%Y-%m-%d %H:%M:%S")

                end_rent_date = fields.Datetime.from_string(self.date_order) + relativedelta(
                    days=int(self.order_line.product_uom_qty))

                tag_number = str(row[1][6])
                fleet = self.env['fleet.vehicle'].sudo().search([('salik_tag_number', '=', tag_number)], limit=1)
                search_transaction_found = self.env['salik.car.transaction'].sudo().search([('transaction_id', '=', int(row[1][0]))])

                if not search_transaction_found:
                    for rec in self.order_line:

                        if rec.product_id:
                            if rec.product_id.product_tmpl_id == fleet.fleet_product_id:
                                # if rec.product_id.related_fleet == fleet:
                                # get_invoice = self.env['account.move.line'].sudo().search([
                                #     ('product_id', '=', rec.product_id.id),
                                #     ('move_id.state', '=', 'draft')
                                # ])

                                get_invoice = self.env['account.move'].sudo().search([
                                    ('partner_id', '=', self.partner_id.id),
                                    ('salik_invoice', '=', True),
                                    ('state', '=', 'draft')
                                ])

                                product_expense_id = self.env.ref('rental_car.product_product_salik_expense_id')
                                product_service_id = self.env.ref('rental_car.product_product_salik_service_id')
                                if not product_service_id:
                                    raise UserError(_("You must enter service product"))
                                if not product_expense_id:
                                    raise UserError(_("You must enter expense product"))

                                get_salik_expense_product = self.env['account.move.line'].sudo().search([
                                    ('product_id', '=', product_expense_id.id),
                                    ('move_id', '=', get_invoice.id),
                                    ('move_id.state', '=', 'draft')
                                ])

                                get_salik_service_product = self.env['account.move.line'].sudo().search([
                                    ('product_id', '=', product_service_id.id),
                                    ('move_id.state', '=', 'draft'),
                                    ('move_id', '=', get_invoice.id)])

                                if get_invoice:
                                    if not get_salik_expense_product:
                                        self.env['account.move.line'].sudo().create({
                                            'product_id': product_expense_id.id,
                                            'quantity': 1,
                                            'move_id': get_invoice.move_id.id,
                                            'sale_line_ids': [Command.link(rec.id)],
                                        })
                                    else:
                                        get_salik_expense_product.sudo().write(
                                            {'quantity': get_salik_expense_product.quantity + 1})

                                    if not get_salik_service_product:
                                        self.env['account.move.line'].sudo().create({
                                            'product_id': product_service_id.id,
                                            'quantity': 1,
                                            'move_id': get_invoice.move_id.id,
                                            'sale_line_ids': [Command.link(rec.id)],
                                        })
                                    else:
                                        get_salik_service_product.sudo().write(
                                            {'quantity': get_salik_service_product.quantity + 1})




                                        print("salik_transction_date", salik_transction_date)
                                        print("salik_transction_date", type(salik_transction_date))
                                        print("date_order", self.date_order)
                                        print("date_order", type(self.date_order))
                                        print("end_rent_date", end_rent_date)
                                        print("end_rent_date", type(end_rent_date))

                                    create_transaction = self.env['salik.car.transaction'].sudo().create({
                                        "name": "/",
                                        "transaction_date": fields.Datetime.now(),
                                        # "gate": get_gate.id,
                                        "transaction_id": int(row[1][0]),
                                        "tag_number": row[1][7],
                                        "plate": row[1][0],
                                        "amount": 4,
                                        "fleet_id": fleet.id,
                                        "sale_order_id": self.id if self.date_order <= salik_transction_date <= end_rent_date else False,
                                        "state": 'allocated' if self.date_order <= salik_transction_date <= end_rent_date else 'un_allocated',

                                    })
                                else:
                                    # get_invoice = self.env['account.move'].sudo().search([
                                    #     ('partner_id', '=', self.partner_id.id),
                                    #     ('state', '=', 'draft')
                                    # ])
                                    # if get_invoice:
                                    #     for rec in get_invoice.invoice_line_ids:
                                    #         if rec.product_id.id == self.env.ref(
                                    #                 'rental_car.product_product_salik_expense_id').id:
                                    #             rec.quantity += 1
                                    #         if rec.product_id.id == self.env.ref(
                                    #                 'rental_car.product_product_salik_service_id').id:
                                    #             rec.quantity += 1
                                    #
                                    #             create_transaction = self.env['salik.car.transaction'].sudo().create({
                                    #                 "name": "/",
                                    #                 "transaction_date": fields.Datetime.now(),
                                    #                 # "gate": get_gate.id,
                                    #                 "transaction_id": int(row[1][0]),
                                    #                 "tag_number": row[1][7],
                                    #                 "plate": row[1][0],
                                    #                 "amount": 4,
                                    #                 "fleet_id": fleet.id,
                                    #             })
                                    #         else:
                                    #             self.env['account.move.line'].sudo().create({
                                    #                 'product_id': product_expense_id.id,
                                    #                 'quantity': 1,
                                    #                 'move_id': get_invoice.id,
                                    #                 'sale_line_ids': [Command.link(rec.id)],
                                    #             })
                                    #             self.env['account.move.line'].sudo().create({
                                    #                 'product_id': product_service_id.id,
                                    #                 'quantity': 1,
                                    #                 'move_id': get_invoice.id,
                                    #                 'sale_line_ids': [Command.link(rec.id)],
                                    #             })
                                    #             create_transaction = self.env['salik.car.transaction'].sudo().create({
                                    #                 "name": "/",
                                    #                 "transaction_date": fields.Datetime.now(),
                                    #                 # "gate": get_gate.id,
                                    #                 "transaction_id": int(row[1][0]),
                                    #                 "tag_number": row[1][7],
                                    #                 "plate": row[1][0],
                                    #                 "amount": 4,
                                    #                 "fleet_id": fleet.id,
                                    #             })
                                    #
                                    # else:
                                        inv = self.env['account.move'].create(
                                            {
                                                'partner_id': self.partner_id.id,
                                                'currency_id': self.currency_id.id,
                                                'move_type': 'out_invoice',
                                                'salik_invoice': True,
                                                'invoice_date': fields.Date.today(),
                                            })

                                        self.env['account.move.line'].sudo().create({
                                            'product_id': product_expense_id.id,
                                            'quantity': 1,
                                            'move_id': inv.id,
                                            'sale_line_ids': [Command.link(rec.id)],


                                        })
                                        self.env['account.move.line'].sudo().create({
                                            'product_id': product_service_id.id,
                                            'quantity': 1,
                                            'move_id': inv.id,
                                            'sale_line_ids': [Command.link(rec.id)],
                                        })



                                        # print(self.date_order+relativedelta(days=int(self.order_line.product_uom_qty)))
                                # 'date_begin': datetime.datetime.now() + relativedelta(days=-1),
                                # 'date_end': datetime.datetime.now() + relativedelta(days=1),

                                        create_transaction = self.env['salik.car.transaction'].sudo().create({
                                            "name": "/",
                                            "transaction_date": fields.Datetime.now(),
                                            # "gate": get_gate.id,
                                            "transaction_id": int(row[1][0]),
                                            "tag_number": row[1][7],
                                            "plate": row[1][0],
                                            "amount": 4,
                                            "fleet_id": fleet.id,
                                            "sale_order_id": self.id if self.date_order <= salik_transction_date <= end_rent_date else False,
                                            "state": 'allocated' if self.date_order <= salik_transction_date <= end_rent_date else 'un_allocated',
                                        })
                        else:
                            raise ValidationError(_('No Car Found in Lines Please add it First'))

                else:
                    raise ValidationError(_('This Transaction Found before'))

######################################################################################################################
                                                         # Fines
######################################################################################################################
    def search_fines(self):

        SERVICE_ACCOUNT_FILE = '/home/odoo/src/user/rental_car/superior-386008-1be184fb4024.json'
        # SERVICE_ACCOUNT_FILE = '/home/osama/odoo/custom/superior_4/rental_car/superior-386008-1be184fb4024.json'

        SCOPES = ['https://www.googleapis.com/auth/drive']

        # Load the service account credentials from the key file
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

        # access_token = creds.token

        try:
            # create drive api client
            service = build('drive', 'v3', credentials=creds)
            files = []
            page_token = None
            while True:
                response = service.files().list(q="name='Fines'",
                                                spaces='drive',
                                                fields='nextPageToken,''files(id,name,owners)',
                                                pageToken=page_token).execute()
                # for file in response.get('files', []):
                #     # Process change
                #     print(F'Found file: {file.get("name")}, {file.mget("id")}')
                files.extend(response.get('files', []))
                page_token = response.get('nextPageToken', None)

                for file in response.get('files', []):
                    file_id = file.get('id')
                    # print(file_id)

                    # file = service.files().get(fileId=file_id).execute()
                    file_content = service.files().export(fileId=file_id,
                                                          mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').execute()
                    file_bytes = io.BytesIO(file_content)

                    # read the Excel sheet into a Pandas DataFrame
                    df = pd.read_excel(file_bytes)

                    # display the data
                    # print(df.head())
                    # print(df)
                    # print(df.dtypes)
                    # assuming 'df' is the DataFrame containing the Excel sheet data
                    for row in df.iterrows():
                        print("###################################")
                    # column_data = df['Direction']
                if page_token is None:
                    break

        except HttpError as error:
            print(F'An error occurred: {error}')
            files = None

        return df.iterrows()

    def get_Fine_transaction(self):
        for rec in self.order_line:
            if rec.product_id:
                found_in_excel=False
                fleet = self.env['fleet.vehicle'].sudo().search([('fleet_product_id', '=', rec.product_id.product_tmpl_id.id)], limit=1)
                print(fleet)
                if fleet:
                    for row in self.search_fines():
                        if not math.isnan(row[1][0]):
                            plate_number = int(row[1][0])
                            plate_code = str(row[1][2])
                            license_plate = str(plate_number)+plate_code
                            if fleet.license_plate == license_plate:
                                found_in_excel = True
                    if not found_in_excel:
                        raise ValidationError(_('No fines Found for this Car (%s) in this  Excel file') % license_plate)
                else:
                    raise ValidationError(_('No Fleet Related to this Product in Rental order Line'))

        for row in self.search_fines():

            if not math.isnan(row[1][0]):
                new_time = datetime.strptime(row[1][7], '%H:%M %p').time()
                date_obj = datetime.strptime(row[1][6], "%d/%m/%Y").date()
                new_datetime = datetime.combine(date_obj, new_time)

                plate_number = str(row[1][0])
                plate_code = str(row[1][2])
                license_plate = plate_number+plate_code
                ticket_number = str(row[1][5])
                fleet = self.env['fleet.vehicle'].sudo().search([('license_plate', '=', license_plate)], limit=1)
                fine_found_before = self.env['car.fines'].sudo().search([('ticket_number', '=', ticket_number)])

                pickup_date = 0
                return_date = 0
                for rec in self.order_line:
                    for record in self.pick_return_line_ids:
                        if record.action_type == 'pickup' and record.deliver_to == 'customer':
                            pickup_date=record.action_time
                        if record.action_type == 'return' and record.received_from == 'customer':
                            return_date = record.action_time

                    if pickup_date:
                        print("pickup time:", pickup_date)
                        if return_date:
                            print("return_date :", return_date)

                if not fine_found_before:
                    if fleet:
                           if pickup_date:
                               if return_date:
                                   if rec.product_id.product_tmpl_id == fleet.fleet_product_id:

                                       get_invoice = self.env['account.move'].sudo().search([
                                           ('partner_id', '=', self.partner_id.id),
                                           ('fine_invoice', '=', True),
                                           ('state', '=', 'draft')
                                       ])

                                       if get_invoice and get_invoice.invoice_line_ids.sale_line_ids.order_id.id == self.id:
                                           get_fine_expense_product = self.env['account.move.line'].sudo().search([
                                               ('product_id', '=',
                                                self.env.ref('rental_car.product_product_fines_expense_id').id),
                                               ('move_id', '=', get_invoice.id)
                                           ])

                                           if not get_fine_expense_product:
                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_expense_id').id,
                                                   'quantity': 1,
                                                   'move_id': get_invoice.id,
                                                   'price_unit': float(row[1][9]),
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })


                                           else:
                                               get_fine_expense_product.sudo().write(
                                                   {'quantity': get_fine_expense_product.quantity + 1,
                                                    'price_unit': get_fine_expense_product.price_unit + float(
                                                        row[1][9]),
                                                    })

                                           ############################### Fine Surcharage Product
                                           product_product_fines_surcharges = self.env[
                                               'account.move.line'].sudo().search([
                                               ('product_id', '=',
                                                self.env.ref('rental_car.product_product_fines_surcharges').id),
                                               ('move_id', '=', get_invoice.id)
                                           ])

                                           if not product_product_fines_surcharges:
                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').id,
                                                   'quantity': 1,
                                                   'move_id': get_invoice.id,
                                                   'price_unit': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').lst_price,
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })


                                           else:
                                               product_product_fines_surcharges.sudo().write(
                                                   {'quantity': product_product_fines_surcharges.quantity + 1,
                                                    'price_unit': product_product_fines_surcharges.price_unit + self.env.ref(
                                                        'rental_car.product_product_fines_surcharges').lst_price,
                                                    })

                                           create_fine = self.env['car.fines'].sudo().create({
                                               "name": "/",
                                               "plate_number": row[1][0],
                                               "plate_category": row[1][1],
                                               "plate_code": row[1][2],
                                               "license_number": row[1][3],
                                               "license_from": row[1][4],
                                               "ticket_number": row[1][5],
                                               "ticket_date": row[1][6],
                                               "ticket_time": row[1][7],
                                               "fines_source": row[1][8],
                                               "ticket_fees": int(row[1][9]),
                                               "ticket_status": row[1][10],
                                               "fleet_id": fleet.id,
                                               "Paid_state": 'un_paid',
                                               "state": 'allocated' if pickup_date <= new_datetime <= return_date else 'un_allocated',
                                               "sale_order_id": self.id if pickup_date <= new_datetime <= return_date else False,

                                           })
                                       else:
                                           get_invoice = self.env['account.move'].sudo().search([
                                               ('partner_id', '=', self.partner_id.id),
                                               ('fine_invoice', '=', True),
                                               ('state', '=', 'draft')
                                           ])

                                           if get_invoice:
                                               if get_invoice.invoice_line_ids:
                                                   for rec in get_invoice.invoice_line_ids:
                                                       if rec.product_id.id == self.env.ref(
                                                               'rental_car.product_product_fines_expense_id').id:
                                                           rec.quantity += 1
                                                           rec.price_unit += float(row[1][9]),

                                                       if rec.product_id.id == self.env.ref(
                                                               'rental_car.product_product_fines_surcharges').id:
                                                           rec.quantity += 1
                                                           rec.price_unit += rec.product_id.lst_price
                                                           create_fine = self.env['car.fines'].sudo().create({
                                                               "name": "/",
                                                               "plate_number": row[1][0],
                                                               "plate_category": row[1][1],
                                                               "plate_code": row[1][2],
                                                               "license_number": row[1][3],
                                                               "license_from": row[1][4],
                                                               "ticket_number": row[1][5],
                                                               "ticket_date": row[1][6],
                                                               "ticket_time": row[1][7],
                                                               "fines_source": row[1][8],

                                                               "ticket_fees": int(row[1][9]),
                                                               "ticket_status": row[1][10],
                                                               "fleet_id": fleet.id,
                                                               "Paid_state": 'un_paid',
                                                               "state": 'allocated' if pickup_date <= new_datetime <= return_date else 'un_allocated',
                                                               "sale_order_id": self.id if pickup_date <= new_datetime <= return_date else False,

                                                           })
                                                       else:
                                                           self.env['account.move.line'].sudo().create({
                                                               'product_id': self.env.ref(
                                                                   'rental_car.product_product_fines_expense_id').id,
                                                               'quantity': 1,
                                                               'price_unit': float(row[1][9]),
                                                               'move_id': get_invoice.id,
                                                               'sale_line_ids': [Command.link(rec.id)],
                                                           })

                                                           self.env['account.move.line'].sudo().create({
                                                               'product_id': self.env.ref(
                                                                   'rental_car.product_product_fines_surcharges').id,
                                                               'quantity': 1,
                                                               'price_unit': self.env.ref(
                                                                   'rental_car.product_product_fines_surcharges').lst_price,
                                                               'move_id': get_invoice.id,
                                                               'sale_line_ids': [Command.link(rec.id)],
                                                           })

                                                           create_fine = self.env['car.fines'].sudo().create({
                                                               "name": "/",
                                                               "plate_number": row[1][0],
                                                               "plate_category": row[1][1],
                                                               "plate_code": row[1][2],
                                                               "license_number": row[1][3],
                                                               "license_from": row[1][4],
                                                               "ticket_number": row[1][5],
                                                               "ticket_date": row[1][6],
                                                               "ticket_time": row[1][7],
                                                               "fines_source": row[1][8],
                                                               "ticket_fees": int(row[1][9]),
                                                               "ticket_status": row[1][10],
                                                               "fleet_id": fleet.id,
                                                               "Paid_state": 'un_paid',
                                                               "state": 'allocated' if pickup_date <= new_datetime <= return_date else 'un_allocated',
                                                               "sale_order_id": self.id if pickup_date <= new_datetime <= return_date else False,

                                                           })
                                               else:

                                                   self.env['account.move.line'].sudo().create({
                                                       'product_id': self.env.ref(
                                                           'rental_car.product_product_fines_expense_id').id,
                                                       'quantity': 1,
                                                       'price_unit': float(row[1][9]),
                                                       'sale_line_ids': [Command.link(rec.id)],
                                                       'move_id': get_invoice.id,
                                                   })
                                                   self.env['account.move.line'].sudo().create({
                                                       'product_id': self.env.ref(
                                                           'rental_car.product_product_fines_surcharges').id,
                                                       'quantity': 1,
                                                       'price_unit': self.env.ref(
                                                           'rental_car.product_product_fines_surcharges').lst_price,
                                                       'move_id': get_invoice.id,
                                                       'sale_line_ids': [Command.link(rec.id)],
                                                   })

                                                   create_fine = self.env['car.fines'].sudo().create({
                                                       "name": "/",
                                                       "plate_number": row[1][0],
                                                       "plate_category": row[1][1],
                                                       "plate_code": row[1][2],
                                                       "license_number": row[1][3],
                                                       "license_from": row[1][4],
                                                       "ticket_number": row[1][5],
                                                       "ticket_date": row[1][6],
                                                       "ticket_time": row[1][7],
                                                       "fines_source": row[1][8],
                                                       "ticket_fees": int(row[1][9]),
                                                       "ticket_status": row[1][10],
                                                       "fleet_id": fleet.id,
                                                       "Paid_state": 'un_paid',
                                                       "state": 'allocated' if pickup_date <= new_datetime <= return_date else 'un_allocated',
                                                       "sale_order_id": self.id if pickup_date <= new_datetime <= return_date else False,

                                                   })
                                           else:
                                               inv = self.env['account.move'].create(
                                                   {
                                                       'partner_id': self.partner_id.id,
                                                       'currency_id': self.currency_id.id,
                                                       'move_type': 'out_invoice',
                                                       'invoice_date': fields.Date.today(),
                                                       'fine_invoice': True,
                                                   })

                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_expense_id').id,
                                                   'quantity': 1,
                                                   'price_unit': float(row[1][9]),
                                                   'move_id': inv.id,
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })

                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').id,
                                                   'quantity': 1,
                                                   'price_unit': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').lst_price,
                                                   'move_id': inv.id,
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })

                                               create_fine = self.env['car.fines'].sudo().create({
                                                   "name": "/",
                                                   "plate_number": row[1][0],
                                                   "plate_category": row[1][1],
                                                   "plate_code": row[1][2],
                                                   "license_number": row[1][3],
                                                   "license_from": row[1][4],
                                                   "ticket_number": row[1][5],
                                                   "ticket_date": row[1][6],
                                                   "ticket_time": row[1][7],
                                                   "fines_source": row[1][8],
                                                   "ticket_fees": int(row[1][9]),
                                                   "ticket_status": row[1][10],
                                                   "fleet_id": fleet.id,
                                                   "Paid_state": 'un_paid',
                                                   "state": 'allocated' if pickup_date <= new_datetime <= return_date else 'un_allocated',
                                                   "sale_order_id": self.id if pickup_date <= new_datetime <= return_date else False,

                                               })
                               else:
                                   if rec.product_id.product_tmpl_id == fleet.fleet_product_id:

                                       get_invoice = self.env['account.move'].sudo().search([
                                           ('partner_id', '=', self.partner_id.id),
                                           ('fine_invoice', '=', True),
                                           ('state', '=', 'draft')
                                       ])

                                       if get_invoice and get_invoice.invoice_line_ids.sale_line_ids.order_id.id == self.id:
                                           get_fine_expense_product = self.env['account.move.line'].sudo().search([
                                               ('product_id', '=',
                                                self.env.ref('rental_car.product_product_fines_expense_id').id),
                                               ('move_id', '=', get_invoice.id)
                                           ])

                                           if not get_fine_expense_product:
                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_expense_id').id,
                                                   'quantity': 1,
                                                   'move_id': get_invoice.id,
                                                   'price_unit': float(row[1][9]),
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })


                                           else:
                                               get_fine_expense_product.sudo().write(
                                                   {'quantity': get_fine_expense_product.quantity + 1,
                                                    'price_unit': get_fine_expense_product.price_unit + float(
                                                        row[1][9]),
                                                    })

                                           ############################### Fine Surcharage Product
                                           product_product_fines_surcharges = self.env[
                                               'account.move.line'].sudo().search([
                                               ('product_id', '=',
                                                self.env.ref('rental_car.product_product_fines_surcharges').id),
                                               ('move_id', '=', get_invoice.id)
                                           ])

                                           if not product_product_fines_surcharges:
                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').id,
                                                   'quantity': 1,
                                                   'move_id': get_invoice.id,
                                                   'price_unit': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').lst_price,
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })


                                           else:
                                               product_product_fines_surcharges.sudo().write(
                                                   {'quantity': product_product_fines_surcharges.quantity + 1,
                                                    'price_unit': product_product_fines_surcharges.price_unit + self.env.ref(
                                                        'rental_car.product_product_fines_surcharges').lst_price,
                                                    })

                                           create_fine = self.env['car.fines'].sudo().create({
                                               "name": "/",
                                               "plate_number": row[1][0],
                                               "plate_category": row[1][1],
                                               "plate_code": row[1][2],
                                               "license_number": row[1][3],
                                               "license_from": row[1][4],
                                               "ticket_number": row[1][5],
                                               "ticket_date": row[1][6],
                                               "ticket_time": row[1][7],
                                               "fines_source": row[1][8],
                                               "ticket_fees": int(row[1][9]),
                                               "ticket_status": row[1][10],
                                               "fleet_id": fleet.id,
                                               "Paid_state": 'un_paid',
                                               "state": 'allocated' if pickup_date <= new_datetime else 'un_allocated',
                                               "sale_order_id": self.id if pickup_date <= new_datetime else False,

                                           })
                                       else:
                                           get_invoice = self.env['account.move'].sudo().search([
                                               ('partner_id', '=', self.partner_id.id),
                                               ('fine_invoice', '=', True),
                                               ('state', '=', 'draft')
                                           ])

                                           if get_invoice:
                                               if get_invoice.invoice_line_ids:
                                                   for rec in get_invoice.invoice_line_ids:
                                                       if rec.product_id.id == self.env.ref(
                                                               'rental_car.product_product_fines_expense_id').id:
                                                           rec.quantity += 1
                                                           rec.price_unit += float(row[1][9]),

                                                       if rec.product_id.id == self.env.ref(
                                                               'rental_car.product_product_fines_surcharges').id:
                                                           rec.quantity += 1
                                                           rec.price_unit += rec.product_id.lst_price
                                                           create_fine = self.env['car.fines'].sudo().create({
                                                               "name": "/",
                                                               "plate_number": row[1][0],
                                                               "plate_category": row[1][1],
                                                               "plate_code": row[1][2],
                                                               "license_number": row[1][3],
                                                               "license_from": row[1][4],
                                                               "ticket_number": row[1][5],
                                                               "ticket_date": row[1][6],
                                                               "ticket_time": row[1][7],
                                                               "fines_source": row[1][8],

                                                               "ticket_fees": int(row[1][9]),
                                                               "ticket_status": row[1][10],
                                                               "fleet_id": fleet.id,
                                                               "Paid_state": 'un_paid',
                                                               "state": 'allocated' if pickup_date <= new_datetime else 'un_allocated',
                                                               "sale_order_id": self.id if pickup_date <= new_datetime else False,

                                                           })
                                                       else:
                                                           self.env['account.move.line'].sudo().create({
                                                               'product_id': self.env.ref(
                                                                   'rental_car.product_product_fines_expense_id').id,
                                                               'quantity': 1,
                                                               'price_unit': float(row[1][9]),
                                                               'move_id': get_invoice.id,
                                                               'sale_line_ids': [Command.link(rec.id)],
                                                           })

                                                           self.env['account.move.line'].sudo().create({
                                                               'product_id': self.env.ref(
                                                                   'rental_car.product_product_fines_surcharges').id,
                                                               'quantity': 1,
                                                               'price_unit': self.env.ref(
                                                                   'rental_car.product_product_fines_surcharges').lst_price,
                                                               'move_id': get_invoice.id,
                                                               'sale_line_ids': [Command.link(rec.id)],
                                                           })

                                                           create_fine = self.env['car.fines'].sudo().create({
                                                               "name": "/",
                                                               "plate_number": row[1][0],
                                                               "plate_category": row[1][1],
                                                               "plate_code": row[1][2],
                                                               "license_number": row[1][3],
                                                               "license_from": row[1][4],
                                                               "ticket_number": row[1][5],
                                                               "ticket_date": row[1][6],
                                                               "ticket_time": row[1][7],
                                                               "fines_source": row[1][8],
                                                               "ticket_fees": int(row[1][9]),
                                                               "ticket_status": row[1][10],
                                                               "fleet_id": fleet.id,
                                                               "Paid_state": 'un_paid',
                                                               "state": 'allocated' if pickup_date <= new_datetime else 'un_allocated',
                                                               "sale_order_id": self.id if pickup_date <= new_datetime else False,

                                                           })
                                               else:

                                                   self.env['account.move.line'].sudo().create({
                                                       'product_id': self.env.ref(
                                                           'rental_car.product_product_fines_expense_id').id,
                                                       'quantity': 1,
                                                       'price_unit': float(row[1][9]),
                                                       'sale_line_ids': [Command.link(rec.id)],
                                                       'move_id': get_invoice.id,
                                                   })
                                                   self.env['account.move.line'].sudo().create({
                                                       'product_id': self.env.ref(
                                                           'rental_car.product_product_fines_surcharges').id,
                                                       'quantity': 1,
                                                       'price_unit': self.env.ref(
                                                           'rental_car.product_product_fines_surcharges').lst_price,
                                                       'move_id': get_invoice.id,
                                                       'sale_line_ids': [Command.link(rec.id)],
                                                   })

                                                   create_fine = self.env['car.fines'].sudo().create({
                                                       "name": "/",
                                                       "plate_number": row[1][0],
                                                       "plate_category": row[1][1],
                                                       "plate_code": row[1][2],
                                                       "license_number": row[1][3],
                                                       "license_from": row[1][4],
                                                       "ticket_number": row[1][5],
                                                       "ticket_date": row[1][6],
                                                       "ticket_time": row[1][7],
                                                       "fines_source": row[1][8],
                                                       "ticket_fees": int(row[1][9]),
                                                       "ticket_status": row[1][10],
                                                       "fleet_id": fleet.id,
                                                       "Paid_state": 'un_paid',
                                                       "state": 'allocated' if pickup_date <= new_datetime else 'un_allocated',
                                                       "sale_order_id": self.id if pickup_date <= new_datetime else False,

                                                   })
                                           else:
                                               inv = self.env['account.move'].create(
                                                   {
                                                       'partner_id': self.partner_id.id,
                                                       'currency_id': self.currency_id.id,
                                                       'move_type': 'out_invoice',
                                                       'invoice_date': fields.Date.today(),
                                                       'fine_invoice': True,
                                                   })

                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_expense_id').id,
                                                   'quantity': 1,
                                                   'price_unit': float(row[1][9]),
                                                   'move_id': inv.id,
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })

                                               self.env['account.move.line'].sudo().create({
                                                   'product_id': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').id,
                                                   'quantity': 1,
                                                   'price_unit': self.env.ref(
                                                       'rental_car.product_product_fines_surcharges').lst_price,
                                                   'move_id': inv.id,
                                                   'sale_line_ids': [Command.link(rec.id)],
                                               })

                                               create_fine = self.env['car.fines'].sudo().create({
                                                   "name": "/",
                                                   "plate_number": row[1][0],
                                                   "plate_category": row[1][1],
                                                   "plate_code": row[1][2],
                                                   "license_number": row[1][3],
                                                   "license_from": row[1][4],
                                                   "ticket_number": row[1][5],
                                                   "ticket_date": row[1][6],
                                                   "ticket_time": row[1][7],
                                                   "fines_source": row[1][8],
                                                   "ticket_fees": int(row[1][9]),
                                                   "ticket_status": row[1][10],
                                                   "fleet_id": fleet.id,
                                                   "Paid_state": 'un_paid',
                                                   "state": 'allocated' if pickup_date <= new_datetime else 'un_allocated',
                                                   "sale_order_id": self.id if pickup_date <= new_datetime else False,

                                               })
                           else:

                               raise ValidationError(_('Car Not delivered To Customer Yet'))

                else:
                    raise ValidationError(_('This Fine Found Before'))


            else:
                raise ValidationError(_('Data in Excel Sheet Not arrange Correctly'))
#
