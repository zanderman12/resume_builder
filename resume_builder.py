# -*- coding: utf-8 -*-
"""
Created on Sun Jun  7 08:29:15 2020

@author: alext
"""


import sys
import jinja2
import pandas as pd
import pdfkit


sys.path.insert(0, 'C:/Users/alext/Documents/GitHub/Python Scripts')
import read_google_sheet as rgs

def load_gsheet_data(gsheet_id, gsheet_sheet, length_of_gsheet):
    gsheet = rgs.get_google_sheet(gsheet_id, gsheet_sheet)
    gdf = rgs.gsheet2df(gsheet, length_of_gsheet)  
    return gdf

if __name__ == '__main__':
    
    ## configure pdfkit so it runs properly
    path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
    
    ## pull data from google sheets, 
    entries_df = load_gsheet_data('1Td8EmLCUS3avoe3sVya7RMu7Rdaqcf01_cNYfgf7wh4', 'entries', 100)
    entries_df.columns = entries_df.iloc[0]
    entries_df.drop(1, axis = 0, inplace = True)
    
    textblocks_df = load_gsheet_data('1Td8EmLCUS3avoe3sVya7RMu7Rdaqcf01_cNYfgf7wh4', 'text_blocks', 100)
    textblocks_df.columns = textblocks_df.iloc[0]
    textblocks_df.drop(1, axis = 0, inplace = True)
    textblocks_df.set_index('loc', inplace = True)
    
    contact_df = load_gsheet_data('1Td8EmLCUS3avoe3sVya7RMu7Rdaqcf01_cNYfgf7wh4', 'contact_info', 100)
    contact_df.columns = contact_df.iloc[0]
    contact_df.drop(1, axis = 0, inplace = True)
    contact_df.set_index('loc', inplace = True)
    
    template_vars = {'name': 'Alexander Cates',
                     'phone': contact_df.loc['phone', 'contact'],
                     'email': contact_df.loc['email','contact'],
                     'linkedin': contact_df.loc['linkedin','contact'],
                     'email_link': contact_df.loc['email','link'],
                     'linkedin_link': contact_df.loc['linkedin','link'],
                     'desired_job_title': 'PhD Student',
                     'mission_statement': textblocks_df.loc['intro', 'text'],
                     'website': contact_df.loc['website', 'contact'],
                     'website_link': contact_df.loc['website', 'link']
                     }
    
    
    
    for cat in ['education', 'research', 'teaching', 'publications', 'awards']:
        catlist = []
        for rrow in entries_df[entries_df.section == cat].iterrows():
            row = rrow[1]
            rowdf = {}
            rowdf['title'] = row.title
            rowdf['company'] = row.institution
            rowdf['start'] = int(row.start)
            if row.end == 'Current':
                rowdf['end'] = 'Current'
            else:
                rowdf['end'] = int(row.end)
            rowdf['desc'] = row.description_1
            catlist.append(rowdf)
        template_vars[cat] = catlist
        
    
    
    
    ## assign data to variables in the template
    
    
    ##load the template html dox
    templateLoader = jinja2.FileSystemLoader(searchpath="./")
    templateEnv = jinja2.Environment(loader=templateLoader)
    TEMPLATE_FILE = "templates/th_resume_loop.html"
    template = templateEnv.get_template(TEMPLATE_FILE)
    
    #fill in the variables with the data
    html_out = template.render(template_vars)
    
    #create and save to an html file
    html_file = open('test_resume.html', 'w', encoding = 'utf-8')
    html_file.write(html_out)
    html_file.close()
    
    #save html to a pdf, make sure to load the appropriate style sheet
    pdfkit.from_file('th_resume_loop.html', 'test_resume.pdf', 
                     options = {'page-size': 'Letter',
                                'margin-top': '0in',
                                'margin-right': '0in',
                                'margin-bottom': '0in',
                                'margin-left': '0in'},
                     configuration=config,
                     css = 'static/th_style_loop.css')
