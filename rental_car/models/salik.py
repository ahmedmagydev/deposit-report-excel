from odoo import fields, models, api,_
import io
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class ReadDriveSheet(models.Model):
    _name = 'readdrive.sheet'
    _description = 'Read Drive Sheet'


    def search_file(self):
        """Search file in drive location

        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
        """
        SERVICE_ACCOUNT_FILE = '/home/osama/odoo/custom/superior_4/rental_car/superior-386008-1be184fb4024.json'

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
                        print(row[1][0])
                        print(row[1][1])
                        print(row[1][2])
                        print(row[1][3])
                        print(row[1][4])
                        print(row[1][5])
                        print(row[1][6])
                        print(row[1][7])
                        print(row[1][8])
                        print("###################################")

                    # column_data = df['Direction']
                if page_token is None:
                    break

        except HttpError as error:
            print(F'An error occurred: {error}')
            files = None

        return file_id

    def check_car(self, **kw):

        if kw['car'] and kw['gate']:
            fleet = self.env['fleet.vehicle'].sudo().search([('salik_tag_number', '=', kw['car']), ], limit=1)
            get_gate = self.env['salik.gates'].sudo().search([('id', '=', kw['gate'])])
            print("fleet", fleet)
            print("get_gate", get_gate)
            if get_gate and fleet:
                create_transaction = self.env['salik.car.transaction'].sudo().create({
                    "name": "/",
                    "transaction_date": fields.Datetime.now(),
                    "gate": get_gate.id,
                    "tag_number": fleet.salik_tag_number,
                    "plate": fleet.license_plate,
                    "amount": 4,
                    "fleet_id": fleet.id,
                })

                ############################ ok ###############################
                ##################################################################

                found_car = self.env['product.template'].sudo().search([
                    ('related_fleet', '=', fleet.id)
                ])
                print("found_car", found_car)

                ##################################################################
                get_invoice = self.env['account.move.line'].sudo().search([
                    ('product_id', '=', found_car.id),
                    ('move_id.state', '=', 'draft')
                ])

                print("invoice line", get_invoice)

                get_salik_product1 = self.env['account.move.line'].sudo().search([
                    ('product_id', '=', 4),
                    ('move_id', '=', get_invoice.move_id.id)
                ])
                get_salik_product2 = self.env['account.move.line'].sudo().search([
                    ('product_id', '=', 5),
                    ('move_id', '=', get_invoice.move_id.id)])

                if get_invoice and not get_salik_product1 and not get_salik_product2:
                    self.env['account.move.line'].sudo().create({
                        'product_id': 4,
                        'quantity': 1,
                        'move_id': get_invoice.move_id.id,
                    })
                    self.env['account.move.line'].sudo().create({
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

                    get_salik_product1.sudo().write({'quantity': get_salik_product1.quantity + 1})
                    get_salik_product2.sudo().write({'quantity': get_salik_product2.quantity + 1})
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


