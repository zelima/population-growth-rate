#!/usr/bin/python
# -*- coding: ascii -*-

import os
import urllib
import unicodecsv as csv # deals with unussual symbols
import xlrd


source = 'https://esa.un.org/unpd/wpp/DVD/Files/1_Indicators%20(Standard)/EXCEL_FILES/1_Population/WPP2015_POP_F02_POPULATION_GROWTH_RATE.XLS'

def setup():
    '''Crates the directorie for archive if they don't exist
    
    '''
    if not os.path.exists('archive'):
        os.mkdir('archive')

def retrieve(source):
    '''Downloades xls data to archive directory
    
    '''
    urllib.urlretrieve(source,'archive/external-data.xls')

def get_data():
    '''Gets the data from xls file and yields lists of it's data row by row
    
    '''
    countries = {}
    fo = xlrd.open_workbook('archive/external-data.xls')
    sheet = fo.sheet_by_index(0) 
    num_rows = sheet.nrows
    num_cols = sheet.ncols
    for i in range(17, num_rows):
        row = []
        for n in range(2, num_cols):
            if not n == 3:  # skips notes column
                value = sheet.cell_value(i,n)
                if n < 5:
                    row.append(value)
                else:
                    period = sheet.cell_value(16,n)
                    new_row = row + [period] + [value]
                    yield new_row

def process(data):
    '''takes generator funtion as input and writes data into csv file
    
    '''
    fo = open('data/population-growth-rate.csv', 'w')
    header = [u'Region', u'Country code', u'Period', u'Growth rate']
    writer = csv.writer(fo, encoding='utf-8')
    writer.writerow(header)
    for row in data:
        writer.writerow(row)
    fo.close()
        
if __name__ == '__main__':
    setup()
    retrieve(source)
    process(get_data())