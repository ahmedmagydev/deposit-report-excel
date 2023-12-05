from odoo import fields,api ,models
import base64 ,io
from io import BytesIO
from datetime import datetime 
import logging
from xlsxwriter.utility import xl_range

_logger = logging.getLogger(__name__)
class Desposit(models.TransientModel):
    _name = 'desposit.wizard'
    
    customer=fields.Many2one('res.partner',string='Customer',required=True)
    datef=fields.Date(string='Date from',required=True)
    dateto=fields.Date(string='Date to',required=True)
    account=fields.Many2one('account.account',string='Account',required=True)
    
    
    
    def button_xlxs_wizard(self):
        print('excel data')
        data={
            
            'customer': self.customer.name,
            'datef': self.datef,
            'dateto': self.dateto,
            'account': self.account.name,
        }
        return self.env.ref('deposit_report.action_report_deposit').report_action(self,data=data)
    
class Desposit_wizard_xlsx(models.AbstractModel):
    _name='report.deposit_report.deposit_xlsx'    
    
    _inherit='report.report_xlsx.abstract'
    
    def generate_xlsx_report(self, workbook, data, wizard):
        print('mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmm')
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

        sheet = workbook.add_worksheet('Deposit Report')

        # Your SQL query with parameters
        sql_query = """
            
            
            SELECT TO_CHAR(am.date, 'YYYY-MM-DD') AS date ,
            am.name ,am.debit , am.credit 
             FROM account_move_line am
             JOIN account_account aa ON am.account_id = aa.id
             JOIN res_partner pp ON am.partner_id = pp.id
             WHERE am.date BETWEEN %s AND %s
             and pp.id = %s
             and aa.name= %s;

        """
        
        # Execute the SQL query  am.date ,am.name ,am.debit , am.credit 
        self.env.cr.execute(sql_query, (wizard.datef, wizard.dateto,wizard.customer.id,wizard.account.name))

        # Fetch all results
        results = self.env.cr.fetchall()
        
        # print(results[0])
        # _logger.error("============= order_line ============.%s",result[2]))
        
        
        
        cell_format_ry = workbook.add_format( {'bg_color': '#BEBEBE', 'font_color': '#000000' , 'bold': 1, 'border': 1 , 'align': 'center'} )
        cell_form=workbook.add_format({'border': 1 ,'align': 'center'}) 
                # Write headers to the worksheet
        headers = [field[0] for field in self.env.cr.description]
        for col_num, header in enumerate(headers ,start=1):
            
            sheet.write(6, col_num, header,cell_format_ry)

        # Write data to the worksheet
        total_debit = 0.0
        total_credit = 0.0
        balanc=0.0
        first_debit = 0.0
        total_refund=0.0
        to_refunded=0.0
        
        for row_num, result in enumerate(results, start=7):
            for col_num, value in enumerate(result,start=1):
                           
                           sheet.write(row_num, col_num, value,cell_form)
        
                                  
            debit = result[2] if result[2] else 0.0
            credit = result[3] if result[3] else 0.0
            difference = debit - credit
            balanc += difference
            # Write the difference to the new column
            sheet.write(row_num, len(result)+1, balanc,cell_form)
            total_debit += debit
            total_credit += credit
            
            if row_num == 7:
                first_debit = debit
            
        

        sheet.write(1,1 ,'customer' ,cell_format_ry)
        sheet.write(1,2 ,wizard.customer.name ,cell_format_ry)
        
        sheet.write(1,4,"Date Range",cell_format_ry)
        sheet.write(1,5,str(wizard.dateto) + ' : ' + str(wizard.datef),cell_format_ry)
        
        sheet.write(2,1,'mobile NO',cell_format_ry)
        sheet.write(2,2,wizard.customer.mobile or '',cell_format_ry) 
        
        sheet.write(3,1,"account",cell_format_ry)
        sheet.write(3,2,wizard.account.name,cell_format_ry)
        
        sheet.write(2,4,'Date',cell_format_ry)
        sheet.write(2,5,datetime.now().strftime("%Y-%m-%d"),cell_format_ry)
        
        sheet.write(4,4,'Bending Balance',cell_format_ry)
        sheet.write(4,5,first_debit,cell_format_ry)
        # _logger.error("============= order_line ============.%s",result[2])
        

            # Write the difference to the new column
        sheet.write(6,5, 'Runing Balance',cell_format_ry)
        
        
        
        sheet.write(row_num + 2, 1, 'Total Debit',cell_format_ry)
        sheet.write(row_num + 2, 2, total_debit,cell_format_ry)

        sheet.write(row_num + 3, 1, 'Total Credit',cell_format_ry)
        sheet.write(row_num + 3, 2, total_credit,cell_format_ry)
        
        
        sheet.write(row_num + 2, 4, 'Ending Balance Tel '+str(wizard.dateto),cell_format_ry)
        sheet.write(row_num + 2, 5,balanc,cell_format_ry) 
        
        sheet.write(row_num + 3, 4,'Current Balance',cell_format_ry)
        sheet.write(row_num + 3, 5,first_debit-balanc,cell_format_ry) 
        
        sheet.write(row_num + 5 , 1, 'Deposit',cell_format_ry)
        
        sql_deposit=""" SELECT
    dd.name,
    dd.direct_type AS type,
    dd.to_refund ,
    TO_CHAR(dd.date, 'YYYY-MM-DD') AS date ,
    dd.rent_contract_id AS Agreemant,
    dd.deposit_value AS amount,
    dd.used_amount AS used,
    refunded_amount AS Refunded,
    dd.state
FROM
    deposit_deposit dd
JOIN
    res_partner pp ON dd.partner_id = %s
WHERE
    (
        (dd.date BETWEEN %s AND %s AND pp.name = 'Deco Addict') OR
        (dd.date < %s AND pp.name = 'Deco Addict' AND dd.state = 'in_hold' AND dd.deposit_type = 'direct') OR
        (dd.date < %s AND pp.name = 'Deco Addict' AND dd.state != 'expired' AND dd.deposit_type = 'preauth')
    );"""
        
        
        self.env.cr.execute(sql_deposit, (wizard.customer.id,wizard.datef, wizard.dateto,wizard.datef,wizard.datef))

    
        results_deposit = self.env.cr.fetchall()
        deposits = [field[0] for field in self.env.cr.description]
        for col_num, header in enumerate(deposits ,start=1):
            # sheet.write(row_num+6, col_num, "", cell_form)  
            sheet.write(row_num + 6, col_num, header,cell_format_ry)
        
                   
        for row_num, result in enumerate(results_deposit, start=row_num + 7):
            for col_num, value in enumerate(result,start=1):
                        #    sheet.write(row_num, col_num, "", cell_form)  
                           sheet.write(row_num, col_num, value,cell_form)    
            to_refunded = result[2] if result[2] else 0.0
            
            total_refund +=to_refunded
           
         
        sheet.write(row_num + 3, 6, 'Amount to be Refunded ',cell_format_ry)
        sheet.write(row_num + 3, 7, total_refund,cell_format_ry)    
    # def get_border_format(self, workbook):
    #     border_format = workbook.add_format()
    #     border_format.set_border(1)  # 1 represents the border style (thin)
        
    #     return border_format    