#! python3

""" Events Registrations

Process compilation file for running events registration review reports.
"""

import pyperclip
import pandas as pd
import subprocess


def copy_sql(report_type):
    """ Copy SQL code

    Create complete SQL query from code snippets
    :param report_type: type of query to run
    """
    if report_type == 'registration':
        sql_filename = 'Event Registrations.sql'
    else:
        sql_filename = 'Communication Recipients.sql'
    sql_file_path = open('Z:\\03 Information Specialist\\Scripts\\Events\\' + sql_filename)
    sql_file_data = sql_file_path.read()

    pyperclip.copy(sql_file_data)
    print('SQL query copied')


def create_csv(data_filename):
    """ Create csv file

    Read clipboard and convert and save to csv file

    :param data_filename: file name of csv file
    """
    results = pyperclip.paste()
    results = [x.split(sep='\t') for x in results.splitlines()]
    df = pd.DataFrame(results, columns=results.pop(0))
    df.to_csv('U:\\COE Advancement\\Work Requests\\Temp Files\\'+data_filename+'.csv',
                            encoding='latin-1', index=False)
    print('csv file created')

answer = 'y'
count = 0
while answer.lower() == 'y' or answer.lower() == 'yes':
    report_type = ''
    while report_type not in ('registration', 'recipient'):
        report_type = input('1. Enter type of event report (registration, recipient):')
    copy_sql(report_type)
    if count == 0:
        # open remote desktop to run query
        subprocess.Popen('mstsc')
    if report_type == 'registration':
        print('2. Paste SQL query into SQL Server and obtain event LookupID')
    elif report_type == 'recipient':
        print('2. Paste SQL query into SQL Server and obtain communication project number')
    print('3. Copy results of query with headers')
    output_filename = input('4. Enter desired name of csv file (without extension):')
    create_csv(output_filename)
    answer = input('Run another event report? (y/n)')
    count += 1
