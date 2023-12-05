# -*- coding: utf-8 -*-

##############################################################################

## By Samia in Aug-2022
import datetime
from report_xlsx.report.report_xlsx import ReportXlsx
from openerp import exceptions
from openerp import models, fields, api

class xx_sal_902_xls_wizard(models.Model):
    _name = "xx_sal_902_xls_wizard"

    xx_from_date = fields.Date('From Create Invoice Date')
    xx_to_date = fields.Date('To Create Invoice Date')
    
    xx_BU = fields.Many2one('xx.product.inv.cat',string='Business Unit')
    xx_salesteam = fields.Many2many('crm.case.section', 'xx_rep_sal_902_teams','xx_wizard_id','xx_team',string='Sales Team(s)')
    xx_report_type = fields.Selection([("Summarized1",'by SKU'), ("Summarized2",'by Sub-Category'), ("Summarized3",'by Main-Category')], default='Summarized3')

    xx_class=fields.Many2many('xx_sales_channel', 'xx_rep_sal_902_class','xx_wizard_id','xx_channel',string='Sales Channel(s)')
    xx_team_or_class = fields.Selection([("Team",'Sales Team'), ("Class",'Sales Channel')])

    ## Print All Customer Totals Columns 
    xx_all_cols = fields.Selection([('P','Customers Performance'), ('PP','Customers Performance and Profile'), ('All','All Columns')], default='P', string='Display Columns')

    xx_segmented_by = fields.Selection([("val",'Net Value'), ("push_val",'Push Value')], string='Segmented By', default='val')
    xx_sort_by = fields.Selection([("val",'Net Value'), ("push_val",'Push Value')], string='Sorted By', default='val')
    xx_thrsh1 = fields.Float( string= 'Segment A-B Threshold (EGP/Month)', default=1000 )
    xx_thrsh2 = fields.Float( string= 'Segment B-C Threshold (EGP/Month)', default=700 )

    @api.multi
    def export_xls(self):
        context = self._context
        datas = {'ids': context.get('active_ids', [])}
        datas['model'] = 'account.invoice.line'
        datas['form'] = self.read()[0]
        for field in datas['form'].keys():
            if isinstance(datas['form'][field], tuple):
                datas['form'][field] = datas['form'][field][0]
        if context.get('xls_export'):
            return {'type': 'ir.actions.report.xml',
                    'report_name': 'reporting_extended.xx_rep_sal_902_xls.xlsx',
                    'datas': datas,
                    'name': 'Customers_Performance_SAL_902'
                    }

class xx_rep_sal_902_xls_ReportXls(ReportXlsx):

    def get_lines(self, data, worksheet, teams_ids, lvl, sortby):

        sort_by=data['form']['xx_sort_by'] ## default
		
        sort_by2=""" val """
        if sortby==2:
            sort_by2=""" qty """   
		
        lines = []

        if data.get('form', False) and data['form'].get('xx_from_date', 'xx_to_date'): 
        ## , 'xx_branch', 'xx_report_type'):
            fdate="""'"""+str(data['form']['xx_from_date'])+"""'"""
            tdate="""'"""+str(data['form']['xx_to_date'])+"""'"""

        buCondition = """ """           
        if data['form']['xx_BU']:
           buCondition = " AND pt.xx_inv_cat='" + str(data['form']['xx_BU']) + "'"

        teamCondition = """ """         
        if data['form']['xx_salesteam']:
            teamCondition = " AND i.section_id in (" + str(teams_ids) + ")"
 
        select_Condition=""" """
        group_by_condition=""" """
        
        sales_channel_condition =""" """
        if data['form']['xx_class']:
            class_v="("+str(data['form']['xx_class'][0])
            for vals in data['form']['xx_class']:
                class_v+=","+str(vals)
            class_v+=")"
            sales_channel_condition= "And t.xx_sales_channel in  "+str(class_v)+" "
                   
        ## Details by SKU
        if data['form']['xx_report_type']=='Summarized1' and worksheet==2 and lvl =='prod':  
                select_Condition="""SELECT team, '' as salesPerson, parent_group_name, group_name, product_id, 
                 sum(quantity)  as qty, sum(subtotal)  as val, 
                count(distinct customer_id) as no_cust, count(distinct Inv_only) as no_inv, 
                salesarea, 0 as sales_id, area_id, team_id, xx_push_or_pull, cust_name, cust_code, cust_id 
                , ROUND( sum(quantity_rtrn) ::numeric,2) as qty_rtrn
				,0,0,0,0,0,0
				from ("""
                group_by_condition=""") as qry 
                group by team, team_id, area_id, salesarea, cust_id, cust_name, cust_code, parent_group_name, group_name, product_id, xx_push_or_pull 
                order by team, team_id, area_id, salesarea, cust_id, """ + str(sort_by2) + """ DESC, product_id;  
                """     
        ## Details by Sub-Category
        if data['form']['xx_report_type']=='Summarized2' and worksheet==2 and lvl =='prod': 
                select_Condition="""SELECT team, '' as salesPerson, parent_group_name, group_name, group_id, 
                ROUND( sum(quantity) ::numeric,2) as qty, ROUND( sum(subtotal) ::numeric,2) as val, 
                count(distinct customer_id) as no_cust, count(distinct Inv_only) as no_inv, 
                salesarea, 0 as sales_id, area_id, team_id, '', cust_name, cust_code, cust_id 
                , ROUND( sum(quantity_rtrn) ::numeric,2) as qty_rtrn
				,0,0,0,0,0,0
				from ("""
                group_by_condition=""") as qry 
                group by team, team_id, area_id, salesarea, cust_id, cust_name, cust_code, parent_group_name, group_name, group_id 
                order by team, team_id, area_id, salesarea, cust_id, """ + str(sort_by2) + """ DESC, group_id; 
                """     
        ## Details by Main-Category
        if data['form']['xx_report_type']=='Summarized3' and worksheet==2 and lvl =='prod':  
                select_Condition="""SELECT team, '' as salesPerson, parent_group_name, '' as group_name, parent_group_id, 
                ROUND( sum(quantity) ::numeric,2) as qty, ROUND( sum(subtotal) ::numeric,2) as val, 
                count(distinct customer_id) as no_cust, count(distinct Inv_only) as no_inv, 
                salesarea, 0 as sales_id, area_id, team_id, '', cust_name, cust_code, cust_id 
                , ROUND( sum(quantity_rtrn) ::numeric,2) as qty_rtrn
				,0,0,0,0,0,0
				from ("""
                group_by_condition=""") as qry 
                group by team, team_id, area_id, salesarea, cust_id, cust_name, cust_code, parent_group_name, parent_group_id 
                order by team, team_id, area_id, salesarea, cust_id, """ + str(sort_by2) + """ DESC, parent_group_id; 
                """    
		## Totals by customer 
        if worksheet==2 and lvl =='cust':
                select_Condition="""SELECT team, '' as salesPerson, '' as parent_group_name, '' as group_name, 0 as parent_group_id, 
                ROUND( sum(quantity) ::numeric,2) as qty, ROUND( sum(subtotal) ::numeric,2) as val, 
                count(distinct customer_id) as no_cust, count(distinct Inv_only) as no_inv, 
                salesarea, 0 as sales_id, area_id, team_id, '', cust_name, cust_code, cust_id 
                , ROUND( sum(quantity_rtrn) ::numeric,2) as qty_rtrn
                , min(invoice_date)	as first_invoice, max(invoice_date)	as last_invoice
                , COALESCE( (max(invoice_date) - min(invoice_date)) ,1) as days_invoice	
                , ROUND( sum(push_val) ::numeric,2) as push_val
                , ROUND( sum(Golden_num) ::numeric,2) as Golden_num
				, count(distinct num_active_months) as num_active_months 
				from ("""
				
                group_by_condition=""") as qry 
                group by team, team_id, area_id, salesarea, cust_id, cust_name, cust_code 
                order by team, team_id, area_id, salesarea, """ + str(sort_by) + """ DESC, cust_id, cust_name, cust_code; 
                """    
				
        customer_secuirty=" "
        product_secuirty=" "
        user_cat=self.env['res.users'].browse(self.env.uid).xx_inv_cat
        if user_cat:
            product_secuirty=self.env['xx_get_bu_product_partner'].get_product_bu_indirct('temp',user_cat.id,'l','prod')
            customer_secuirty=self.env['xx_get_bu_product_partner'].get_partner_bu('p',user_cat.id)

        statment=str(select_Condition)+"""
                    SELECT  i.id as inv_id, 
					CASE WHEN i.type='out_invoice' THEN i.date_invoice ELSE null END as invoice_date,
                       
                    p.name as cust_name, p.ref as cust_code, p.id as cust_id,
                    u2.name as salesPerson, t.name as team, i.partner_id as customer_id,
                    pt.name as product, t.id as team_id,
                    
                    CASE WHEN (i.type='out_invoice' and pt.type!='service') THEN l.quantity ELSE 0 END as quantity,

                    CASE WHEN (i.type='out_refund'  and pt.type!='service') THEN 0-l.quantity ELSE 0 END as quantity_rtrn,

                    CASE WHEN pt.type!='service'  
							 THEN (CASE WHEN i.type='out_invoice' THEN l.price_subtotal ELSE 0-l.price_subtotal END) 
							 ELSE 0 END as subtotal

                    ,l.xx_prmotion_line as Is_Promotion, l.product_id as product_id  
                    ,l.id as invoice_line_id, 
					                    
                    c0.id as parent_group_id, c0.xx_latin_name as parent_group_name, c.id as group_id, c.xx_latin_name as group_name,
                    
                    i.create_date::date as create_invoice_date,
                    r.name as salesarea, i.user_id as sales_id, r.id as area_id
                    
                    , CASE WHEN i.type='out_invoice' THEN i.create_date::date ELSE null END as Inv_only  --repeated invoice per day will be counted as 1
                    , pt.xx_push_or_pull
                    , CASE WHEN (i.type='out_invoice' and pt.xx_push_or_pull='push' and pt.type!='service') THEN l.price_subtotal ELSE 0 END as push_val
                    , CASE WHEN (i.type='out_invoice' and l.product_id= 1109) THEN 1 ELSE 0 END as Golden_num
                    , CASE WHEN i.type='out_invoice' THEN date_part('month', i.create_date::date) ELSE null END as num_active_months  

                    from account_invoice i, account_invoice_line l,  
						 product_product pr, product_template pt, 
						 product_category c, product_category c0,
						 res_users u, res_partner u2, crm_case_section t, 
						 res_partner p, xx_res_city_area r
                    where l.invoice_id=i.id
                    and pt.id=pr.product_tmpl_id 
					and (pt.type!='service' or l.product_id= 1109)
                    and pt.categ_id = c.id  and c.parent_id= c0.id
                    and pr.id=l.product_id  
                    and l.quantity!=0
                    and t.id=i.section_id
                    and u.id=i.user_id
                    and u.partner_id=u2.id
                    and i.create_date::date >= """+str(fdate)+"""
                    and i.create_date::date <= """+str(tdate)+str(customer_secuirty)+str(product_secuirty)+"""
                    and i.partner_id=p.id and r.id = p.xx_area_id
					and p.active = 't' 
                    and i.type in ('out_invoice','out_refund') 
                    and i.state in ('open', 'paid')
                    """+str(buCondition)+str(teamCondition)+str(sales_channel_condition)+"""
                    and l.product_id is not null and i.user_id is not null 
					                    
                    """+str(group_by_condition)
        #print "Qry Statment ==================================================>"
        #print statment
        self.env.cr.execute(statment) 
        rows=self.env.cr.fetchall()

        ## ===================================================================
        if (worksheet==2):  
            for obj in rows:
                vals = {                
                    'team'              : obj[0] or '',
                    'salesPerson'       : obj[1] or '',
                    'parent_group_name' : obj[2] or '',
                    'sub_group_name'    : obj[3] or '',
                    'id'                : obj[4] or 0,
                    'qty'               : obj[5] or 0,
                    'val'               : obj[6] or 0,
                    'no_cust'           : obj[7] or 0,
                    'no_inv'            : obj[8] or 0,
                    'area'              : obj[9] or 0,
                    'sales_id'          : obj[10] or 0,
                    'area_id'           : obj[11] or 0,
                    'team_id'           : obj[12] or 0,
                    'xx_push_or_pull'   : obj[13] or '',
                    'cust_name'   		: obj[14] or '',
                    'cust_code'   		: obj[15] or '',
                    'cust_id'   		: obj[16] or 0,
                    'qty_rtrn'   		: obj[17] or 0,
                    'first_invoice' 	: obj[18] or 0,
                    'last_invoice'   	: obj[19] or 0,
                    'days_invoice' 		: obj[20] or 0,
                    'push_val'   		: obj[21] or 0,
                    'Golden_num' 		: obj[22] or 0,
                    'num_active_months' : obj[23] or 0,   
                    ##'product'         : short_product_name or '',
                    ##'latin_product_name': latin_product_name or '',
                    }
                lines.append(vals)

        ## In all cases return lines  
        return lines
		
## ========================================================================================

    def generate_xlsx_report(self, workbook, data, lines):

        ## Add Format 
        ## Options = {'font': {'bold': True, 'name': 'Arial', 'color': 'red','size': 12} , 'fill': {'color': 'green'}, }
        cell_format_ry = workbook.add_format( {'bg_color': '#fff700', 'font_color': '#9C0006' , 'bold': 1, 'border': 1, 'align': 'left'} )
        cell_format_r = workbook.add_format( {'bg_color': '#FFC7CE', 'font_color': '#9C0006' , 'bold': 1, 'border': 1, 'align': 'left'} )
        cell_format_g = workbook.add_format( {'bg_color': '#C6EFCE', 'font_color': '#006100' , 'bold': 1, 'border': 1, 'align': 'left'} )
        cell_format_y = workbook.add_format( {'bg_color': '#fff700', 'font_color': '#000000' , 'bold': 1, 'border': 1, 'align': 'left'} )
        cell_format_x = workbook.add_format( {'bg_color': '#C1E9F2', 'font_color': '#000000' , 'bold': 1, 'border': 1, 'align': 'left'} )
        cell_format_b = workbook.add_format( {'bg_color': '#92BDFE', 'font_color': '#000000' , 'bold': 1, 'border': 1, 'align': 'left'} )
        cell_format_zrb = workbook.add_format( {'font_color': '#000000' , 'border': 1, 'bold': 1, 'align': 'right'} )
        cell_format_zgb = workbook.add_format( {'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'bold': 1, 'align': 'right'} )
        cell_format_zr = workbook.add_format( {'font_color': '#000000' , 'border': 1, 'align': 'right'} )
        cell_format_zl = workbook.add_format( {'font_color': '#000000' , 'border': 1, 'align': 'left'} )
        cell_format_yc = workbook.add_format( {'bg_color': '#fff700', 'font_color': '#000000' , 'bold': 1, 'border': 1, 'align': 'center'} )
        cell_format_hl = workbook.add_format( {'font_color': '#000000' , 'bold': 1, 'align': 'left'} )
     
        ## Worksheet2: by SKU matrix by customer
        worksheet2 = workbook.add_worksheet('Customer Perfomance')
        
        buCondition = """ """   
        teamCondition = """ """     
        teams_ids = """ """     

        catg_num = 0        
        catgs=[]
        
        ## Print the Parameters in all sheets 
        sort_by=data['form']['xx_sort_by']
        segmented_by=data['form']['xx_segmented_by'] 
        thrsh1= data['form']['xx_thrsh1']  
        thrsh2= data['form']['xx_thrsh2']
        xx_all_cls = data['form']['xx_all_cols']
        xx_all_cols_name = 'Customers Performance'
        if xx_all_cls == 'PP' :
		   xx_all_cols_name = 'Customers Performance and Profile'
        elif xx_all_cls == 'All' :        
		   xx_all_cols_name = 'All Columns'	
		
        ## =================================================================== 
        ## get channels and sale teams ids and names
        sales_channel_condition =""" """
        st_name = """ """
        worksheet2.write('D1', 'Sales Channel ', cell_format_x)
        
        if data['form']['xx_class']:
            class_name=''
            class_v="("+str(data['form']['xx_class'][0])
            for vals in data['form']['xx_class']:
                class_v+=","+str(vals)
            class_v+=")"
            get_c_name="SELECT name from xx_sales_channel where id in "+str(class_v)+" ; "
            self.env.cr.execute(get_c_name)
            cname = self.env.cr.fetchall()
            for v in cname:
                class_name+=" //"+str(v[0])
            worksheet2.write('E1',class_name)
            
            ## get all sales teams in the channel 
            get_c_teams="SELECT id from crm_case_section where xx_sales_channel in "+str(class_v)+" Order by xx_sales_channel, id; "
            self.env.cr.execute(get_c_teams)
            rows = self.env.cr.fetchall()
            idss = []
            for obj in rows:
                idss.append(obj[0])
                
            data['form']['xx_salesteam'] = idss
        ## --------------------------------------
        ## get teams ids and names
        ##teams_names=[] 
        if data['form']['xx_salesteam']:
                teams_ids= str(data['form']['xx_salesteam'][0])
                ##print data['form']['xx_salesteam']
                for val in data['form']['xx_salesteam']:
                   teams_ids=teams_ids+","+str(val)
                   team_name = self.env['crm.case.section'].browse(val).name
                   ##teams_names.append(team_name) ## team_name
                   st_name=st_name + " - " + team_name
                teamCondition = " AND i.section_id in (" + str(teams_ids) + ")"
                worksheet2.write('D2', 'Sales Team(s)', cell_format_x)
                worksheet2.write('E2', st_name)

        if data.get('form', False) and data['form'].get('xx_from_date', 'xx_to_date'):
            fdate=str(data['form']['xx_from_date'])
            tdate=str(data['form']['xx_to_date'])

            Period_lnth	= 0 
            input = data['form']['xx_to_date']
            y2 = int(fields.Datetime.from_string(input).year)
            m2 = int(fields.Datetime.from_string(input).month)
            input = data['form']['xx_from_date']
            y1 = int(fields.Datetime.from_string(input).year)
            m1 = int(fields.Datetime.from_string(input).month)
			
            if y1 == y2:
			   Period_lnth = (m2 - m1) + 1
            else: 
			   Period_lnth = (m2 - m1) + (y2 - y1) * 12 + 1
             
            worksheet2.write('A1', 'From Date', cell_format_x)
            worksheet2.write('A2', 'To Date', cell_format_x)
            worksheet2.write('B1', fdate)
            worksheet2.write('B2', tdate)                       
            worksheet2.write('A3', 'Business Unit', cell_format_x)
            worksheet2.write('D3', 'Report Display', cell_format_x)
			
            worksheet2.write('A4', 'Sorted By', cell_format_x)
            worksheet2.write('B4', 'Net Value')
            if sort_by != 'val':
               worksheet2.write('B4', 'Push Value')
						
            worksheet2.write('D4', 'All Columns', cell_format_x)
            worksheet2.write('E4', xx_all_cols_name)
			
            worksheet2.write('G1', 'Segmented By', cell_format_x)
            worksheet2.write('H1', 'Net Value')                       
            if segmented_by != 'val':
               worksheet2.write('H1', 'Push Value')
            worksheet2.write('G2', 'Segment A-B Threshold (EGP/Month)', cell_format_x)
            worksheet2.write('H2', thrsh1)                       
            worksheet2.write('G3', 'Segment B-C Threshold (EGP/Month)', cell_format_x)
            worksheet2.write('H3', thrsh2)                       	

            if data['form']['xx_BU']:
                bu_name=self.env['xx.product.inv.cat'].browse(data['form']['xx_BU']).name
                worksheet2.write('B3', bu_name)
                buCondition = " AND pt.xx_inv_cat='" + str(data['form']['xx_BU']) + "'"                                       

        ## Print the Data -----------------------------------------------
        check_restriction = self.env['xx_check_report_retricted'].check_report_retricted('reporting_extended.xx_rep_sal_902_xls.xlsx')

        if check_restriction == True:
            msg = 'You Can NOT Run This Report In Peak Hour'
            worksheet1.write('G9', msg, cell_format_ry)
            worksheet2.write('G9', msg, cell_format_ry)
        else:
            ##--------------------------------------------------------------------------------
            ## Print Actual products_vals in sheet2 by customer - 2D Array/Table MATRIX        
 
            worksheet = 2
 
            pd_col = 8  	## padding columns 
            pd_row = 8  	## padding rows 
            col_tot  = 24  	## for customer perfomance columns 			
            col_tot2 = 14  	## for customer profile columns 
			
            if xx_all_cls == 'All' :
				cl = pd_col + col_tot + col_tot2 - 1
				worksheet2.write(3, cl, 'Group Name', cell_format_g) 
				worksheet2.write(5, cl, 'ID', cell_format_g)               
				## get the list of Products for Horizontal columns
				if data['form']['xx_report_type']=='Summarized1':  ##'by SKU'
					statment_prods = """ SELECT distinct c0.xx_latin_name as parent_group_name, 
										c.xx_latin_name as group_name,
										p.id as product_ID, 
										pt.xx_latin_name as Latin_Full_Name, 
										-- pt.name as Product_Short_Name,           
										-- pt.xx_short_name as Arabic_Full_Name, 
										pt.xx_sequence, c0.id, c.id
										from product_category c, product_category c0, product_product p, product_template pt
										Where  
										p.product_tmpl_id = pt.id and pt.type!='service'
										and pt.categ_id = c.id  
										and c.parent_id = c0.id 
										and pt.xx_sequence != 0
										and pt.active = true 
										
										--and p.id in (1176, 1603, 808, 811, 812)   
										
										"""+str(buCondition)+"""
										order by pt.xx_sequence, c0.id, c.id ; """
					worksheet2.write(4, cl, 'Sub Group Name', cell_format_g) 
					worksheet2.write(6, cl, 'Product Name', cell_format_g)               
					worksheet2.write('E3', 'by SKU')
				
				elif data['form']['xx_report_type']=='Summarized2':  ##'by Sub-Category'
					statment_prods = """ SELECT distinct c0.xx_latin_name as parent_group_name, 
									c.xx_latin_name as group_name, c.id, ''
									from product_category c, product_category c0, product_product p, product_template pt
									Where  
									p.product_tmpl_id = pt.id and pt.type!='service'
									and pt.categ_id = c.id  
									and c.parent_id = c0.id 
									and pt.xx_sequence != 0
									and pt.active = true
									"""+str(buCondition)+"""
									order by c0.xx_latin_name, c.xx_latin_name; """
					worksheet2.write(4, cl, 'Sub Group Name', cell_format_g) 
					worksheet2.write('E3', 'by Sub-Category')
				
				else: ##'by Main-Category'
					statment_prods = """ SELECT distinct c0.xx_latin_name as parent_group_name, 
									'' , c0.id, ''
									from product_category c, product_category c0, product_product p, product_template pt
									Where  
									p.product_tmpl_id = pt.id and pt.type!='service'
									and pt.categ_id = c.id  
									and c.parent_id = c0.id 
									and pt.xx_sequence != 0
									and pt.active = true
									"""+str(buCondition)+"""
									order by c0.xx_latin_name; """
					worksheet2.write('E3', 'by Main-Category')
							
				self.env.cr.execute(statment_prods)   
				get_products = self.env.cr.fetchall()
			
			## --------------------------------------------------------------
			## Main Columns 
            worksheet2.write('A8', 'Channel', cell_format_y)
            worksheet2.write('B8', 'Team', cell_format_y)
            worksheet2.write('C8', 'Sales Area ID', cell_format_y)
            worksheet2.write('D8', 'Sales Area Name', cell_format_y)
            worksheet2.write('E8', 'Customer Code', cell_format_yc)
            worksheet2.write('F8', 'Customer Name', cell_format_yc)
			## Customer Perfomance Columns 
            worksheet2.write('H8', 'Sold Volume', cell_format_r)
            worksheet2.write('I8', 'Returned Volume', cell_format_r)
            worksheet2.write('J8', 'Net Value', cell_format_r)
            worksheet2.write('K8', 'No. of Invoices', cell_format_r)
            worksheet2.write('L8', 'Push Value', cell_format_r)
            worksheet2.write('M8', 'Push %', cell_format_r)
            worksheet2.write('N8', 'No. of Golden Invoice', cell_format_r)    	
            worksheet2.write('O8', 'First Invoice Date', cell_format_r)    	
            worksheet2.write('P8', 'Last Invoice Date', cell_format_r)    	
            worksheet2.write('Q8', 'Frequency/Selling Cycle (days)', cell_format_r)    	
            worksheet2.write('R8', 'Segment', cell_format_r)    	
            worksheet2.write('S8', 'Active Months', cell_format_r)    	    
            worksheet2.write('T8', 'Average Sales per Month', cell_format_r)    	    
            worksheet2.write('U7', 'Top 5 SKU(s)/Category(s) by Value', cell_format_r)    	    
            worksheet2.write('U8', '1st', cell_format_r)    	    
            worksheet2.write('V8', '2nd', cell_format_r)    	    
            worksheet2.write('W8', '3rd', cell_format_r)    	    
            worksheet2.write('X8', '4th', cell_format_r)    	    
            worksheet2.write('Y8', '5th', cell_format_r)    	    

            worksheet2.write('AA7', 'Top 5 SKU(s)/Category(s) by Volume', cell_format_r)    	    
            worksheet2.write('AA8', '1st', cell_format_r)    	    
            worksheet2.write('AB8', '2nd', cell_format_r)    	    
            worksheet2.write('AC8', '3rd', cell_format_r)    	    
            worksheet2.write('AD8', '4th', cell_format_r)    	    
            worksheet2.write('AE8', '5th', cell_format_r)    	    
            ##--------------------------------------------------------------------------------
            if xx_all_cls == 'All' :				
				prods=[]              
				col = pd_col + col_tot + col_tot2
				row = pd_row - 5
				col_set = 4
				prod_num=0
				for p in get_products:
					prods.append(p[2])  ## p_ID
					worksheet2.write(row  , col,  p[0], cell_format_hl) # Parent Category
					worksheet2.write(row+1, col,  p[1], cell_format_hl) # Sub Category
					worksheet2.write(row+2, col,  p[2], cell_format_hl) # ID
					worksheet2.write(row+3, col,  p[3], cell_format_hl) # Product Name
					worksheet2.write(row+4, col  , 'Sold Volume', cell_format_g)
					worksheet2.write(row+4, col+1, 'Return Volume', cell_format_g)
					worksheet2.write(row+4, col+2, 'Value', cell_format_g)
					
					col += col_set
					prod_num += 1
				                        
            ##--------------------------------------------------------------------------------
            ## get the list of active customer(s) and Grand totals for Vertical rows 
                        
            custs=[]  
            col = 0			
            row = pd_row
            cust_num = 0

            ##--------------------------------------------------------------------------------
            ## Display Customer Perfomance				
            ## Totals 
            tqty = 0 
            tqty_rtrn = 0 
            tval = 0                 
            tno_inv = 0 
            tpush_val = 0 
            tGolden_num = 0 

            segA = 0 
            segB = 0 
            segC = 0 
			
            lvl ='cust'
            get_line1 = self.get_lines(data, worksheet, teams_ids, lvl, 1)			
            for each in get_line1:
                custs.append(each['cust_id'])  ## Cust_id 
                
                ##Channel, Team, Area_ID, Area_Name, Customer_Code, Customer_Name                   
                worksheet2.write(row, col  , self.env['crm.case.section'].browse(each['team_id']).xx_class, cell_format_zl) ## channel
                worksheet2.write(row, col+1, each['team'], cell_format_zl)  
                worksheet2.write(row, col+2, each['area_id'], cell_format_zr)
                worksheet2.write(row, col+3, each['area'], cell_format_zl)
                worksheet2.write(row, col+4, int(each['cust_code']) , cell_format_zr)
                worksheet2.write(row, col+5, each['cust_name'], cell_format_zl)
				
				## Sold Volume,	Returned Volume,	Net Value,		
                worksheet2.write(row, col+7, each['qty'], cell_format_zr)
                worksheet2.write(row, col+8, each['qty_rtrn'], cell_format_zr)
                worksheet2.write(row, col+9, each['val'], cell_format_zr)
				
				## No. of Invoices, Push_Value,	No. of Golden Invoice					
                worksheet2.write(row, col+10, each['no_inv'], cell_format_zr)
                worksheet2.write(row, col+11, each['push_val'], cell_format_zr)
				
                push=0				
                if (each['val'])!=0: 
				   push = round( float(each['push_val']) / (float(each['val'])) ,4)
                worksheet2.write(row, col+12, push, cell_format_zr) ## push %
				
				
                worksheet2.write(row, col+13, each['Golden_num'], cell_format_zr)
				
				## First Invoice Date,	Last Invoice Date,	Frequency (days)
                worksheet2.write(row, col+14, each['first_invoice'], cell_format_zr)
                worksheet2.write(row, col+15, each['last_invoice'], cell_format_zr)
                Freq=1				
                if (each['no_inv']-1)!=0: 
				   Freq = round( float(each['days_invoice']) / (float(each['no_inv']-1)) ,1)
                worksheet2.write(row, col+16, Freq, cell_format_zr) # Frequency

                ## Calculate the Sales Segment
                ## ------------------------------------------------------------------		
                ## Example: 
                ## 1. More than thrsh1 EGP/Month and we call them customers A
                ## 2. Between thrsh1 and thrsh2 EGP/Month and we call them customers B
                ## 3. Lower than thrsh2 EGP/Month and we call them customers C
                ## ------------------------------------------------------------------		
                num_active_months = each['num_active_months'] ## get number of active months for each customer  

                if num_active_months!=0:
					val_per_month = 0
					if segmented_by == 'val':   ## segment by net value
						 val_per_month = float(each['val']) / float(num_active_months)  
					else: 
						 val_per_month = float(each['push_val']) / float(num_active_months)  ## segment by push value

					if val_per_month >= thrsh1:  ## 'A'
						 seg = 'A'   
						 segA += 1   
					elif val_per_month < thrsh1 and val_per_month >= thrsh2:  ## 'B'
						 seg = 'B'
						 segB += 1   
					elif val_per_month < thrsh2: ## 'C'
						 seg = 'C'	
						 segC += 1   
					
					worksheet2.write(row, col+17, seg, cell_format_zr) # Segment
					worksheet2.write(row, col+18, num_active_months, cell_format_zr) # Number of active months
					worksheet2.write(row, col+19, val_per_month, cell_format_zr) # value months
                else:				
					worksheet2.write(row, col+17, '', cell_format_zr) # Segment
					worksheet2.write(row, col+18, '', cell_format_zr) # Number of active months
					worksheet2.write(row, col+19, '', cell_format_zr) # Number of active months				
				
                top5sku = ''   
                worksheet2.write(row, col+20, top5sku, cell_format_zr) # Top 5 SKUs by value
                worksheet2.write(row, col+21, top5sku, cell_format_zr) # Top 5 SKUs by value
                worksheet2.write(row, col+22, top5sku, cell_format_zr) # Top 5 SKUs by value
                worksheet2.write(row, col+23, top5sku, cell_format_zr) # Top 5 SKUs by value
                worksheet2.write(row, col+24, top5sku, cell_format_zr) # Top 5 SKUs by value

                worksheet2.write(row, col+26, top5sku, cell_format_zr) # Top 5 SKUs by volum
                worksheet2.write(row, col+27, top5sku, cell_format_zr) # Top 5 SKUs by volum
                worksheet2.write(row, col+28, top5sku, cell_format_zr) # Top 5 SKUs by volum
                worksheet2.write(row, col+29, top5sku, cell_format_zr) # Top 5 SKUs by volum
                worksheet2.write(row, col+30, top5sku, cell_format_zr) # Top 5 SKUs by volum

                # fill in all products cells with Zero
                if xx_all_cls == 'All' :				
					ci = pd_col + col_tot + col_tot2
					for p in range(prod_num):
						worksheet2.write(row, ci  , 0, cell_format_zr)
						worksheet2.write(row, ci+1, 0, cell_format_zr)                      
						worksheet2.write(row, ci+2, 0, cell_format_zr)                      
						ci += col_set
                                                        
				## Totals 
                tqty +=each['qty'] 
                tqty_rtrn+=each['qty_rtrn'] 
                tval+=each['val'] 
                
                ## No. of Invoices, Push_Value                
                tno_inv+=each['no_inv']
                tpush_val+=each['push_val']
                tGolden_num+=each['Golden_num']

                row+= 1
                cust_num += 1

            ##--------------------------------------------------------------------------------
            lvl ='prod'
            get_line2 = self.get_lines(data, worksheet, teams_ids, lvl, 1)

            ## Print Top 5 SKU(s) by Value 
            tot_col = pd_col + col_tot - 12
            current_cust=0
            for each in get_line2:
				if each['cust_id']!=current_cust:
					Top5 = 0
					current_cust = each['cust_id']
				if Top5 < 5 and each['parent_group_name']!='Services' and each['id']!=1109:
					try:						
						## get Customer Row index      
						ir = custs.index(each['cust_id']) 
						r = ir + pd_row								
						if data['form']['xx_report_type']=='Summarized3':  ##'by Main-Category' 
							worksheet2.write(r, tot_col+Top5, each['parent_group_name'], cell_format_zl)												
						elif data['form']['xx_report_type']=='Summarized2':  ##'by Sub-Category'
							worksheet2.write(r, tot_col+Top5, each['group_name'], cell_format_zl)												
						else: ##'by SKU'
							worksheet2.write(r, tot_col+Top5, self.env['product.product'].browse(each['id']).name_template, cell_format_zl)						
					except ValueError as ve:
					   worksheet2.write('AX3', ' ')  
					Top5 +=1		

            ## Print Top 5 SKU(s) by Volume 
            lvl ='prod'
            get_line2 = self.get_lines(data, worksheet, teams_ids, lvl, 2)
            tot_col = pd_col + col_tot - 6
            current_cust=0
            for each in get_line2:
				if each['cust_id']!=current_cust:
					Top5 = 0
					current_cust = each['cust_id']
				if Top5 < 5 and each['parent_group_name']!='Services' and each['id']!=1109:
					try:						
						## get Customer Row index      
						ir = custs.index(each['cust_id']) 
						r = ir + pd_row								
						if data['form']['xx_report_type']=='Summarized3':  ##'by Main-Category' 
							worksheet2.write(r, tot_col+Top5, each['parent_group_name'], cell_format_zl)												
						elif data['form']['xx_report_type']=='Summarized2':  ##'by Sub-Category'
							worksheet2.write(r, tot_col+Top5, each['group_name'], cell_format_zl)												
						else: ##'by SKU'
							worksheet2.write(r, tot_col+Top5, self.env['product.product'].browse(each['id']).name_template, cell_format_zl)						
					except ValueError as ve:
					   worksheet2.write('AX3', ' ')  
					Top5 +=1		

            ##--------------------------------------------------------------------------------
            ## Print Totals
            worksheet2.write('A7', 'Total Number of Customers', cell_format_y) 
            worksheet2.write('B7', cust_num, cell_format_y) 
            worksheet2.write('H6', 'Customers Performance Totals', cell_format_r) 
            row = 6
            col = 0
            ## Sold Volume, Returned Volume,    Net Value,      
            worksheet2.write(row, col+7, tqty, cell_format_zrb)
            worksheet2.write(row, col+8, tqty_rtrn, cell_format_zrb)
            worksheet2.write(row, col+9, tval, cell_format_zrb)
                
            ## No. of Invoices, Push_Value                    
            worksheet2.write(row, col+10, tno_inv, cell_format_zrb)
            worksheet2.write(row, col+11, tpush_val, cell_format_zrb)
                
            push=0              
            if (tval)!=0: 
               push = round( float(tpush_val) / (float(tval)) ,4)
            worksheet2.write(row, col+12, push, cell_format_zrb) ## push %
			
            worksheet2.write(row, col+13, tGolden_num, cell_format_zrb) ## Golden

            if Period_lnth!=0:				 
					if segmented_by == 'val':  
						 val_per_month = float(tval) / float(Period_lnth)  ## segment by net value
					else: 
						 val_per_month = float(tpush_val) / float(Period_lnth)  ## segment by push value

            worksheet2.write(row, col+18, Period_lnth, cell_format_zrb)  
            worksheet2.write(row, col+19, val_per_month, cell_format_zrb)  


            worksheet2.write('Q1', 'Count of Segment A', cell_format_y) 
            worksheet2.write('Q2', 'Count of Segment B', cell_format_y) 
            worksheet2.write('Q3', 'Count of Segment C', cell_format_y) 
            worksheet2.write('R1', segA, cell_format_y)
            worksheet2.write('R2', segB, cell_format_y)
            worksheet2.write('R3', segC, cell_format_y)
            prcnt = round( float(segA) / (float(cust_num)) ,4)
            worksheet2.write('S1', prcnt, cell_format_y)
            prcnt = round( float(segB) / (float(cust_num)) ,4)
            worksheet2.write('S2', prcnt, cell_format_y)
            prcnt = round( float(segC) / (float(cust_num)) ,4)
            worksheet2.write('S3', prcnt, cell_format_y)

            ##--------------------------------------------------------------------------------
            ## Display Customer Profile			
            if xx_all_cls != 'P' :
               col = pd_col + col_tot -1
               row = pd_row - 1
               worksheet2.write(row-1, col+1, 'Customers Profile', cell_format_b) 
               ##worksheet2.write(row, col,   'Customer ID', cell_format_b)
			   
               worksheet2.write(row, col+1, 'Geographical Area', cell_format_b)
               worksheet2.write(row, col+2, 'Longitude', cell_format_b)
               worksheet2.write(row, col+3, 'Latitude', cell_format_b)
			   
               worksheet2.write(row, col+4, 'Store Type', cell_format_b)
			   
               worksheet2.write(row, col+5, 'Activation Date with UD', cell_format_b)
               worksheet2.write(row, col+6, 'in Market Since', cell_format_b)
               worksheet2.write(row, col+7, 'No. of Employees', cell_format_b)
               worksheet2.write(row, col+8, 'Size of Store', cell_format_b)
               worksheet2.write(row, col+9, 'IT Readiness', cell_format_b)

               worksheet2.write(row, col+10, 'Decision Maker', cell_format_b)
               worksheet2.write(row, col+11, 'Phone', cell_format_b)
               worksheet2.write(row, col+12, 'Mobile', cell_format_b)
			   
               row = pd_row 
               for each in get_line1:
				   res_partner = self.env['res.partner'].browse(each['cust_id'])
				   ##worksheet2.write(row, col, each['cust_id'], cell_format_zr)  ## Cust_id 
				   if res_partner.xx_nilson_area: ## Geo-area name  
					  worksheet2.write(row, col+1, res_partner.xx_nilson_area.name, cell_format_zl)  
				   else:   
					  worksheet2.write(row, col+1, '', cell_format_zr)  
					  
				   worksheet2.write(row, col+2, res_partner.partner_longitude, cell_format_zr)   
				   worksheet2.write(row, col+3, res_partner.partner_latitude, cell_format_zr)
				   
				   worksheet2.write(row, col+4, res_partner.xx_class, cell_format_zr) 
				   
				   worksheet2.write(row, col+5, res_partner.create_date, cell_format_zr)   

				   if res_partner.xx_market_since:   
					  worksheet2.write(row, col+6, res_partner.xx_market_since, cell_format_zr) 
				   else:   
					  worksheet2.write(row, col+6, '', cell_format_zr)  

				   worksheet2.write(row, col+7, res_partner.xx_no_of_store_employee, cell_format_zr)   
				   worksheet2.write(row, col+8, res_partner.xx_size_of_store, cell_format_zr) 

				   if res_partner.xx_it_readiness:   
					  worksheet2.write(row, col+9, res_partner.xx_it_readiness, cell_format_zr) 
				   else:   
					  worksheet2.write(row, col+9, '', cell_format_zr)  		      
			 		   
				   worksheet2.write(row, col+10, 'Contact Person !!!', cell_format_zr)   
				   worksheet2.write(row, col+11, res_partner.phone, cell_format_zr)   
				   worksheet2.write(row, col+12, res_partner.mobile, cell_format_zr)   
				   row +=1
			
            ##--------------------------------------------------------------------------------
            ## Display SKU(s)/Category(s)
            if xx_all_cls == 'All' : 
				tot_col = pd_col + col_tot + col_tot2
				for each in get_line2:
					if each['parent_group_name']!='Services' and each['id']!=1109:
						try:
							## get product Col index  
							ic = prods.index(each['id']) 
							c = ic * col_set + tot_col
							
							## get Customer Row index      
							ir = custs.index(each['cust_id']) 
							r = ir + pd_row
							
							worksheet2.write(r, c+0, each['qty'], cell_format_zr)
							worksheet2.write(r, c+1, each['qty_rtrn'], cell_format_zr)
							worksheet2.write(r, c+2, each['val'], cell_format_zr)
							 
						except ValueError as ve:
							worksheet2.write('AX1', 'Some Products Have NO Sequence and Have been DROPPED! ID=', cell_format_r)  
							worksheet2.write('AY1', each['id'], cell_format_r)  

				worksheet2.write(1, tot_col-1, 'Total Number of Items', cell_format_g) 
				worksheet2.write(1, tot_col, prod_num, cell_format_g) 			
			
            ##--------------------------------------------------------------------------------

       ##workbook.close()                   
xx_rep_sal_902_xls_ReportXls('report.reporting_extended.xx_rep_sal_902_xls.xlsx', 'account.invoice.line')


Modify this report SAL902 by adding new paramter "Sales-Rep" which is an optional parameter (i.e. can be null).

if the user select a certain sales-rep in  sales-team limit the results to this sales-rep only. 


