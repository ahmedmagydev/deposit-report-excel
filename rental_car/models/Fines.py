from odoo import fields, models, api,_
from odoo.exceptions import UserError, ValidationError, AccessError, RedirectWarning
import io
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
# Import math Library
import math

#
# class SaleOrder(models.Model):
#     _inherit = "sale.order"
#
#     def search_file(self):
#
#         SERVICE_ACCOUNT_FILE = '//home/osama/PycharmProjects/pythonProject1/odoo16/custom/superior/rental_car/superior-386008-1be184fb4024.json'
#         SCOPES = ['https://www.googleapis.com/auth/drive']
#         creds = service_account.Credentials.from_service_account_file(
#             SERVICE_ACCOUNT_FILE, scopes=SCOPES)
#         try:
#             service = build('drive', 'v3', credentials=creds)
#             files = []
#             page_token = None
#             while True:
#                 response = service.files().list(q="name='Test_Fines'",
#                                                 spaces='drive',
#                                                 fields='nextPageToken,''files(id,name,owners)',
#                                                 pageToken=page_token).execute()
#                 # for file in response.get('files', []):
#                 #     # Process change
#                 #     print(F'Found file: {file.get("name")}, {file.mget("id")}')
#                 files.extend(response.get('files', []))
#                 page_token = response.get('nextPageToken', None)
#
#                 for file in response.get('files', []):
#                     file_id = file.get('id')
#                     print(file_id)
#
#                     # file = service.files().get(fileId=file_id).execute()
#                     file_content = service.files().export(fileId=file_id,
#                                                           mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').execute()
#                     file_bytes = io.BytesIO(file_content)
#
#                     # read the Excel sheet into a Pandas DataFrame
#                     df = pd.read_excel(file_bytes)
#
#                     # display the data
#                     # print(df.head())
#                     # print(df)
#                     # print(df.dtypes)
#                     # assuming 'df' is the DataFrame containing the Excel sheet data
#                     # for row in df.iterrows():
#                     #     print("###################################")
#                         # print(row[1][0])
#                         # print(row[1][1])
#                         # print(row[1][2])
#                         # print(row[1][3])
#                         # print(row[1][4])
#                         # print(row[1][5])
#                         # print(row[1][6])
#                         # print(row[1][7])
#                         # print(row[1][8])
#                         # print("###################################")
#
#                     # column_data = df['Direction']
#                 # if page_token is None:
#                 #     break
#             return df.iterrows()
#         except HttpError as error:
#             print(F'An error occurred: {error}')
#             files = None
#
#
#
#
#
#     def get_Fine_transaction(self):
#         print(self.search_file())
#         for row in self.search_file():
#             print("#########Tag Number############Plate Number##############")
#
#             # if row[1][0] != None:
#             if not math.isnan(row[1][0]):
#                 x = str(row[1][6])
#                 y = str(row[1][7])
#                 fleet1 = self.env['fleet.vehicle'].sudo().search([('salik_tag_number', '=', x)], limit=1)
#                 fleet = self.env['fleet.vehicle'].sudo().search([('license_plate', '=', y)], limit=1)
#                 if fleet and fleet1:
#                     print("fleet", fleet)
#                     print("fleet1", fleet1)
#
#                 print(int(row[1][0]))
#                 search_transaction = self.env['salik.car.transaction'].sudo().search([('transaction_id','=',int(row[1][0]))])
#                 if not search_transaction:
#                     create_transaction = self.env['salik.car.transaction'].sudo().create({
#                         "name": "/",
#                         "transaction_date": fields.Datetime.now(),
#                         # "gate": get_gate.id,
#                         "transaction_id": int(row[1][0]),
#                         "tag_number": row[1][7],
#                         "plate": row[1][6],
#                         "amount": 4,
#                         "fleet_id": fleet.id,
#                     })
#                 # print("Tag Number",row[1][6])
#                 # print("Plate Number",row[1][7])
#                 # print(int(row[1][0]))
#                 # print(type(row[1][0]))
#
#
#                     for rec in self.order_line:
#                         print(rec.product_id.related_fleet,fleet)
#                         if rec.product_id.related_fleet == fleet:
#                             get_invoice = self.env['account.move.line'].sudo().search([
#                                 ('product_id', '=', rec.product_id.id),
#                                 ('move_id.state', '=', 'draft')
#                             ])
#
#                             print("invoice line", get_invoice)
#
#                             get_salik_expense_product = self.env['account.move.line'].sudo().search([
#                                 ('product_id', '=', 51),
#                                 ('move_id', '=', get_invoice.move_id.id)
#                             ])
#                             get_salik_service_product = self.env['account.move.line'].sudo().search([
#                                 ('product_id', '=', 52),
#                                 ('move_id', '=', get_invoice.move_id.id)])
#
#                             if get_invoice:
#                                 if not get_salik_expense_product:
#                                     self.env['account.move.line'].sudo().create({
#                                         'product_id': 51,
#                                         'quantity': 1,
#                                         'move_id': get_invoice.move_id.id,
#                                     })
#                                 else:
#                                      get_salik_expense_product.sudo().write({'quantity': get_salik_expense_product.quantity + 1})
#
#                                 if not get_salik_service_product:
#                                     self.env['account.move.line'].sudo().create({
#                                         'product_id': 52,
#                                         'quantity': 1,
#                                         'move_id': get_invoice.move_id.id,
#                                     })
#                                 else:
#                                     get_salik_service_product.sudo().write({'quantity': get_salik_service_product.quantity + 1})
#
#
#                             else:
#                                 raise ValidationError(_('No Draft invoice Found For This Product'))
#                 else:
#                     raise ValidationError(_('This Transaction Found before'))
#





