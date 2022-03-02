# -*- coding: utf-8 -*-
"""
Created on Sun Jun  7 08:29:15 2020

@author: alext
"""


import sys
import jinja2
import pandas as pd
import pdfkit
import datetime


sys.path.insert(0, 'C:/Users/alext/Documents/GitHub/Python Scripts')
import read_google_sheet as rgs

def load_gsheet_data(gsheet_id, gsheet_sheet, length_of_gsheet):
    gsheet = rgs.get_google_sheet(gsheet_id, gsheet_sheet)
    gdf = rgs.gsheet2df(gsheet, length_of_gsheet)  
    return gdf

def clean_order(e):
    if e == 'Current':
        n = datetime.datetime.now()
        return n.year
    else:
        return int(e)

if __name__ == '__main__':
    
    ## configure pdfkit so it runs properly
    path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    
    ## pull data from google sheets CREDENTIALS NEED TO BE RESET, 
    # entries_df = load_gsheet_data('1Td8EmLCUS3avoe3sVya7RMu7Rdaqcf01_cNYfgf7wh4', 'entries', 100)
    entries_df = pd.read_excel('cates_cv_data.xlsx', sheet_name='entries')
    entries_df.columns = entries_df.iloc[0]
    entries_df.drop(0, axis = 0, inplace = True)
    entries_df = entries_df[entries_df.in_resume == 'T']
    
    # textblocks_df = load_gsheet_data('1Td8EmLCUS3avoe3sVya7RMu7Rdaqcf01_cNYfgf7wh4', 'text_blocks', 100)
    textblocks_df = pd.read_excel('cates_cv_data.xlsx', sheet_name='text_blocks')
    textblocks_df.columns = textblocks_df.iloc[0]
    textblocks_df.drop(0, axis = 0, inplace = True)
    textblocks_df = textblocks_df[textblocks_df.in_resume == 'T']
    textblocks_df.set_index('loc', inplace = True)
    
    # contact_df = load_gsheet_data('1Td8EmLCUS3avoe3sVya7RMu7Rdaqcf01_cNYfgf7wh4', 'contact_info', 100)
    contact_df = pd.read_excel('cates_cv_data.xlsx', sheet_name='contact_info')
    contact_df.columns = contact_df.iloc[0]
    contact_df.drop(0, axis = 0, inplace = True)
    
    contact_df.set_index('loc', inplace = True)
    
    
    template_vars = {'name': 'Alexander Cates',
                     'phone': contact_df.loc['phone', 'contact'],
                     'email': contact_df.loc['email','contact'],
                     'linkedin': contact_df.loc['linkedin','contact'],
                     'email_link': contact_df.loc['email','link'],
                     'linkedin_link': contact_df.loc['linkedin','link'],
                     'desired_job_title': 'Ph.D. Candidate',
                     'mission_statement': textblocks_df.loc['intro', 'text'],
                     'website': contact_df.loc['website', 'contact'],
                     'website_link': 'https://' + contact_df.loc['website', 'link']
                     }
    
    
    entries_df['order_end_year'] = entries_df['end'].apply(lambda x: clean_order(x))
    for cat in ['education', 'research', 'teaching', 'publications', 
                'awards', 'presentations', 'service', 'membership', 'data']:
        catlist = []
        for rrow in entries_df[entries_df.section == cat].sort_values('order_end_year', ascending = False).iterrows():
            row = rrow[1]
            rowdf = {}
            rowdf['title'] = row.title
            rowdf['company'] = row.institution
            rowdf['start'] = int(row.start)
            if row.end == 'Current':
                rowdf['end'] = 'Current'
            else:
                rowdf['end'] = int(row.end)
            if cat in ['research', 'teaching', 'service', 'data']:
                desc = row.description_1.split(': ')
                rowdf['desc'] = desc[0] + ':'
                rowdf['bullets'] = desc[1].split('; ')
            else:
                rowdf['desc'] = row.description_1
            catlist.append(rowdf)
        template_vars[cat] = catlist
            
    
    ##load the template html dox
    templateLoader = jinja2.FileSystemLoader(searchpath="./")
    templateEnv = jinja2.Environment(loader=templateLoader)
    TEMPLATE_FILE = "templates/th_resume_loop_bullets.html"
    template = templateEnv.get_template(TEMPLATE_FILE)
    
    #fill in the variables with the data
    html_out = template.render(template_vars)
    
    #create and save to an html file
    html_file = open('test_resume.html', 'w', encoding = 'utf-8')
    html_file.write(html_out)
    html_file.close()
    
    #save html to a pdf, make sure to load the appropriate style sheet
    pdfkit.from_file('test_resume.html', 'test_resume.pdf', 
                     options = {'page-size': 'Letter',
                                'margin-top': '0.5in',
                                'margin-right': '0in',
                                'margin-bottom': '0.5in',
                                'margin-left': '0in'},
                     configuration=config,
                     css = 'static/th_style_loop.css')
