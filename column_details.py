#!/usr/bin/env python
# -*- coding: utf-8 -*-

import psycopg2
import psycopg2.extras
import xlrd
import xlsxwriter
import unicodedata


workbook = xlsxwriter.Workbook('dictionnaire.xlsx')
worksheet = workbook.add_worksheet()

bold = workbook.add_format({'bold': True})

worksheet.write('A1', u"Modèle", bold)
worksheet.write('B1', 'Codification', bold)
worksheet.write('C1', u"Désignation", bold)
worksheet.write('D1', 'Type', bold)
worksheet.write('E1', 'Taille', bold)
worksheet.write('F1', 'Observation', bold)



conn = psycopg2.connect("host='localhost' dbname='socolait_v1.4.11' user='odoo10' password='odoo10' ")
print conn.encoding
conn.set_client_encoding('latin1')
cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
print conn.encoding

in_workbook = xlrd.open_workbook("model.xlsx")
in_worksheet = in_workbook.sheet_by_index(0)
models = []
for rownum in xrange(in_worksheet.nrows):
    if rownum > 0:
        rline = in_worksheet.row_values(rownum)
        models.append(rline[0])
iterator = 1
for model in models:
    print model
    sql = u"""select *
                   from information_schema.columns
                   --where table_schema NOT IN ('information_schema', 'pg_catalog')
                   where table_schema='public' and table_name='%s'
                   order by table_schema, table_name""" % model
    cur.execute(sql)
    for row in cur:
        table = row['table_name']
        cur_row = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
##        cur_row.execute(u"""SELECT description FROM pg_catalog.pg_description WHERE objsubid = 
##            (
##                SELECT ordinal_position FROM  information_schema.columns WHERE table_name='%s' AND column_name='%s'
##            ) 
##            and objoid = 
##            (
##                SELECT oid FROM pg_class WHERE relname = 'account_account' AND relnamespace = 
##                            (
##                                SELECT oid FROM pg_catalog.pg_namespace WHERE nspname = 'public'
##                            )
##                    )""" % (model,row['column_name']) )
        dot_model = model.replace('_','.')
        
        query = """
                    select field_description, size, required from ir_model_fields where model_id = (select id from ir_model where model = '%s' limit 1) and name = '%s'
                """ % (dot_model,row['column_name'])
        cur_row.execute(query)
        description  = ""
        for r in cur_row:
            description = r[0] if r else ""
            if description:
                cur_trans = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
                helping = description.replace("'","''")
                sql = """
                    SELECT value FROM ir_translation
                    WHERE lang='fr_FR' AND type in ('model') AND 
                    src='%s'
                """ % helping
                cur_trans.execute(sql)
                for rr in cur_trans:
                    description = rr[0] if rr else description
      
        worksheet.write(iterator, 0, row['table_name'])
        worksheet.write(iterator, 1, row['column_name'])
        worksheet.write(iterator, 2, description.decode('latin1'))
        worksheet.write(iterator, 3, row['data_type'])
        worksheet.write(iterator, 4, row['character_maximum_length'] if row['character_maximum_length'] else row['numeric_precision'])
        worksheet.write(iterator, 5, row['is_nullable'])
        iterator += 1
workbook.close()
    
    #print "schema: {schema}, table: {table}, column: {col}, type: {type}, description: {description}".format(
        #schema = row['table_schema'], table = row['table_name'],
        #col = row['column_name'], type = row['data_type'], description = description)
    #print "column: {col}, type: {type}, description: {description}".format(
        #col = row['column_name'], type = row['data_type'], description = description)
    
    
