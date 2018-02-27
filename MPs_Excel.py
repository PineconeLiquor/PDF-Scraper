''' This file searches a directory and all sub directories for Duralab
and LIMS MP pdfs and returns the measurements to an excel file.'''

import pdfquery
import os
import fnmatch
import pandas as pd
# import matplotlib.pyplot as plt

# rootdir = input('Copy parent filepath')
rootdir = 'C:\TEMP'

# Name of CSV file to Output
csv_name = 'csv_test' + '.csv'


def dura_mp(file_path):
    '''Takes pdf from file_path and returns dictionary of MPs.
     Only for 1 page, only for Duralab format. '''

    pdf = pdfquery.PDFQuery(file_path)
    pdf.load()
    test1 = (pdf.extract([
        ('with_formatter', 'text'),
        ('Mechanical Test',
            'LTTextLineHorizontal:overlaps_bbox("270, 720, 280, 740")'),
    ]))
    for k, v in test1.items():
        if 'Mechanical' in v:
            # Coordinates based on Duralabs Cert
            return (pdf.extract([
                ('with_formatter', 'text'),
                ('Ultimate Load',
                 'LTTextLineHorizontal:overlaps_bbox("180, 388, 200, 399")'),
                ('Ultimate(mpa)',
                 'LTTextLineHorizontal:overlaps_bbox("360, 380, 382, 389")'),
                ('Yield(mpa)',
                 'LTTextLineHorizontal:overlaps_bbox("360, 360, 382, 363")'),
                ('BHN',
                 'LTTextLineHorizontal:overlaps_bbox("180, 350, 190, 363")'),
                ('Elongation',
                 'LTTextLineHorizontal:overlaps_bbox("180, 330, 190, 337")'),
            ]))
        else:
            # Coordinates based on LIMS Cert
            return (pdf.extract([
                ('with_formatter', 'text'),
                ('Ultimate Load',
                 'LTTextLineHorizontal:overlaps_bbox("134, 288, 135, 296")'),
                ('Ultimate(mpa)',
                 'LTTextLineHorizontal:overlaps_bbox("430, 280, 440, 286")'),
                ('Yield(mpa)',
                 'LTTextLineHorizontal:overlaps_bbox("430, 246, 440, 260")'),
                ('BHN',
                 'LTTextLineHorizontal:overlaps_bbox("134, 240, 144, 250")'),
                ('Elongation',
                 'LTTextLineHorizontal:overlaps_bbox("134, 230, 136, 234")'),
            ]))


# Lists that have data appended to
BHN = []
Elong = []
Ult_mpa = []
Yield_mpa = []
Ult_load = []

# Blank dictionary for pdf_query return
dict1 = {}
# Searches the files for string 'Mechanical' and 'MPs'


def pdf_query(rootdir):
    ''' Searches the files for string 'Mechanical' and 'MP' '''
    count = 0
    for subdir, dirs, files in os.walk(rootdir):
        for filename in files:
            if fnmatch.fnmatch(filename, '*Mechanical*.pdf') or \
                    fnmatch.fnmatchcase(filename, '*MPs*.pdf'):
                dict1 = dura_mp(os.path.join(subdir, filename))
                Elong.append(dict1['Elongation'])
                BHN.append(dict1['BHN'])
                Ult_mpa.append(dict1['Ultimate(mpa)'])
                Yield_mpa.append(dict1['Yield(mpa)'])
                Ult_load.append(dict1['Ultimate Load'])
                count += 1
    print(str(count) + ' files scraped')


pdf_query(rootdir)


def list_config(L):
    '''Combines multiple values into nested csv list if they are a string
        and removes % and lbs if they are present'''

    for string in L:
        if len(string) > 8:
            L.pop(L.index(string))
            s = string.split(' ')
            s = [x for x in s if x != '%']
            s = [x for x in s if x != 'lbs']
            s = [x for x in s if x != 'lb']
            L.extend(s)


list_config(BHN)
list_config(Yield_mpa)
list_config(Ult_mpa)
list_config(Ult_load)
list_config(Elong)


# Prints each list for testing purposes

# print('BHN: ' + str(BHN))
# print('Elongation: ' + str(Elong))
# print('Ultimate(mpa): ' + str(Ult_mpa))
# print('Ultimate Load: ' + str(Ult_load))
# print('Yield(mpa): ' + str(Yield_mpa))


all_dict = {'BHN': BHN, 'Elongation': Elong, 'Ultimate(mpa)': Ult_mpa,
            'Ultimate Load': Ult_load, 'Yield(mpa)': Yield_mpa}
df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in all_dict.items()]))
df.to_csv(csv_name)
