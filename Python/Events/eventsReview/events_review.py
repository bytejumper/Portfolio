#! python3

""" eventsReview

Manipulate data to create review lists and tables in Excel
"""

import pandas as pd
import numpy as np
import os


def basic(read_file):
    """ Read data file

    :param read_file: file name
    :return: df: dataframe of constituents
    """
    df = pd.read_csv(read_file,
                     dtype={'HOUSEHOLDLOOKUPID': str,
                            'PROSPECTLOOKUPID': str,
                            'SPOUSELOOKUPID': str},
                     encoding='latin1')
    return df


def management(df):
    """ Create composite column of PGL, PM, and LAG managers

    :param df: dataframe of constituents
    :return df: modified dataframe
    """
    def row_apply(row):
        """ Create composite of PGL, PM, and LAG managers

        :param row: dataframe row
        :return : composite value
        """
        values = {row['PGLname'], row['PLANMANAGERPLANTYPE']}
        values = {x for x in values if x == x}
        # values.discard(np.nan)
        return '|'.join(values)
    df['Management'] = df.apply(row_apply, axis=1)
    return df


def delete(df, format_type):
    """ Delete excess columns

    'UIF' and 'ENG' formats delete degree detail columns; leaves giving
    information.
    'Event Flag' format leaves PC membership, and EVS donor information.

    :param df: dataframe of constituents
    :param format_type: type of formatting
    :return df: modified dataframe
    """
    if format_type == 'Event Flag':
        stops = [
            [df.columns.get_loc('PRESIDENTSCOUNCILFY'), df.columns.get_loc('ENGDegree3Concat')],
            [df.columns.get_loc('ISDECEASED'), df.columns.get_loc('EngHHGiving')],
            [df.columns.get_loc('HOUSEHOLDLOOKUPID'), df.columns.get_loc('HOUSEHOLDDISPLAYNAME')]
        ]
    else:
        stops = [
            [df.columns.get_loc('PRESIDENTSCOUNCIL'), df.columns.get_loc('PRESIDENTSCOUNCILFY')],
            [df.columns.get_loc('totaluidegrees'), df.columns.get_loc('majordescription3')],
            [df.columns.get_loc('ZIPCODE'), df.columns.get_loc('Email')],
            [df.columns.get_loc('ISDECEASED'), df.columns.get_loc('SPOUSEDECEASED')],
            [df.columns.get_loc('HOUSEHOLDLOOKUPID'), df.columns.get_loc('HOUSEHOLDDISPLAYNAME')]
        ]
    for i in stops:
        df.drop(df.columns[i[0]:i[1] + 1], inplace=True, axis=1)
    return df


def event_pivots(df):
    """ Create pivot tables

    Pivot constituents by management and also by degree department.

    :param df: dataframe of constituents
    :return byPM: pivot table of managed constituents
    :return byDept: pivot table of department alumni
    """
    by_pm = pd.pivot_table(df[df['Management'] != ''],
                           index=['Management', 'PROSPECTNAME'],
                           values=['EngHHGiving', 'UrbanaHHGiving', 'LifeHHGiving'],
                           aggfunc=np.sum)

    by_dept = pd.pivot_table(df[df['ENGDegreeDeptsConcat'] != '0'],
                             index=['ENGDegreeDeptsConcat', 'PROSPECTNAME'],
                             values=['EngHHGiving', 'UrbanaHHGiving', 'LifeHHGiving'],
                             aggfunc=np.sum)

    # Reorder columns (pivot table defaults is to display in alpha order)
    for p in [by_pm, by_dept]:
        p.reindex_axis(['EngHHGiving', 'UrbanaHHGiving', 'LifeHHGiving'],
                       axis=1, copy=False)
    return by_pm, by_dept


def mgos(df, contentdir):
    """ Determine constituents under management by Engineering MGOs

    :param df: dataframe of constituents
    :param contentdir: path to gift officer file
    :return mgmt_df: dataframe of managed constituents
    :return mgo: dictionary of Engineering MGOs
    :return df: dataframe of un-managed ENG alumni constituents
    """
    mgo = eval(open(contentdir, 'r').read())

    mgmt_df = pd.DataFrame()
    for name in mgo:
        mgmt_df = pd.concat([mgmt_df, df[df['Management'].str.contains(name)]])

    df = df.merge(mgmt_df[['PROSPECTLOOKUPID', 'PROSPECTNAME']],
                  how='left', on='PROSPECTLOOKUPID')
    df = df[pd.isnull(df['PROSPECTNAME_y'])]
    df.drop('PROSPECTNAME_y', inplace=True, axis=1)
    df.rename(columns={'PROSPECTNAME_x': 'PROSPECTNAME'}, inplace=True)
    df.dropna(subset=['ENGDegreeDeptsConcat'], inplace=True)
    return mgmt_df, mgo, df


def format_file(df, dest_filename, format_type, contentdir, lookup):
    """ Creates output excel file

    'UIF' format creates four sheets: Summary Stats, Full List, Pivot Tables - PM, and
    Pivot Tables - Dept
    'ENG' format creates multiple sheets: Summary Stats, Full List, ENG departments, and
    ENG MGOs
    'Event Flag' format creates one sheet, filtered to PC members and EVS donors

    :param df: dataframe of constituents
    :param dest_filename: filename to save to
    :param format_type: type of format desired
    :param contentdir: full filepath of MGO data
    :param lookup: lookupid of event
    """
    dest_file = os.path.splitext(os.path.basename(dest_filename))[0] + '.xlsx'
    os.chdir(os.path.dirname(dest_filename))
    if format_type == 'UIF':
        mgmt_df, mgo = mgos(df, contentdir)[:2]
        for i in mgmt_df['Management'].index:
            temp_mgo = set()
            for name in mgo:
                if mgmt_df.ix[i, 'Management'].__contains__(name):
                    temp_mgo.add(name)
            mgmt_df.ix[i, 'Management'] = '|'.join(temp_mgo)
        by_pm = event_pivots(mgmt_df)[0]
        by_dept = event_pivots(df)[1]
    elif format_type == 'ENG':
        mgmt_df, mgo, limit_df = mgos(df, contentdir)
        limit_df.drop(['Management'], inplace=True, axis=1)
    elif format_type == 'Event Flag':
        df['EVSFundDonor'] = df['EVSFundDonor'].replace(0, np.nan)
        df.dropna(subset=['PRESIDENTSCOUNCIL', 'EVSFundDonor'],
                  thresh=1, inplace=True)
    df.drop(['Management'], inplace=True, axis=1)
    with pd.ExcelWriter(dest_file, engine='xlsxwriter') as writer:
        wb = writer.book
        # write summary sheet
        if format_type != 'Event Flag':
            ws = wb.add_worksheet('Summary Stats')
            ws.set_landscape()
            ws.set_column('A:A', 1)
            ws.set_row(0, 9)
            for c in ['B:B', 'D:D', 'F:F', 'H:H', 'J:J']:
                ws.set_column(c, 22)
            for c in ['C:C', 'E:E', 'G:G', 'I:I', 'K:K']:
                ws.set_column(c, 2)

            format_dict = {
                'font_name': ['Cambria', 'Calibri'],
                'color': ['#4F81BD', '#A6A6A6', '#FFFFFF'],
                'border_style': [2]
            }
            section_heading = wb.add_format({
                'font_name': format_dict['font_name'][0],
                'font_size': 14,
                'font_color': format_dict['color'][1],
                'top': format_dict['border_style'][0],
                'bottom': format_dict['border_style'][0],
                'border_color': format_dict['color'][1]})
            summary_heading = wb.add_format({
                'font_name': format_dict['font_name'][0],
                'font_size': 9,
                'font_color': format_dict['color'][2],
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'bg_color': format_dict['color'][0],
                'top': format_dict['border_style'][0],
                'left': format_dict['border_style'][0],
                'right': format_dict['border_style'][0],
                'border_color': format_dict['color'][0]})
            summary_content = wb.add_format({
                'font_name': format_dict['font_name'][1],
                'font_size': 20,
                'font_color': format_dict['color'][1],
                'num_format': '0',
                'align': 'center',
                'valign': 'vcenter',
                'left': format_dict['border_style'][0],
                'right': format_dict['border_style'][0],
                'bottom': format_dict['border_style'][0],
                'border_color': format_dict['color'][1]
            })
            table_heading = wb.add_format({
                'font_name': format_dict['font_name'][0],
                'font_size': 11,
                'font_color': format_dict['color'][2],
                'bold': True,
                'bg_color': format_dict['color'][0],
                'top': 1,
                'bottom': 1,
                'border_color': '#D9D9D9'})
            table_contents = wb.add_format({
                'font_name': format_dict['font_name'][1],
                'font_size': 10,
                'text_wrap': True,
                'top': 1,
                'bottom': 1,
                'border_color': '#D9D9D9'})

            ws.write('B2', dest_file[:-11] + ' (' + lookup + ')', wb.add_format({
                'font_name': format_dict['font_name'][0],
                'font_size': 24,
                'font_color': format_dict['color'][0]}))
            ws.write('I2', 'As of', wb.add_format({
                'italic': True,
                'font_name': format_dict['font_name'][1],
                'font_size': 11,
                'font_color': format_dict['color'][1],
                'align': 'right'}))
            ws.write('J2', dest_file[dest_file.index('_')+1:-5], wb.add_format({
                'num_format': 'mm/dd/yyyy',
                'font_name': format_dict['font_name'][0],
                'font_size': 20,
                'font_color': format_dict['color'][1],
                'top': format_dict['border_style'][0],
                'bottom': format_dict['border_style'][0],
                'border_color': format_dict['color'][1]}))
            ws.write('B3', 'College of Engineering Affiliates', wb.add_format({
                'font_name': format_dict['font_name'][1],
                'font_size': 18,
                'font_color': format_dict['color'][1]}))

            for i in [['B5:J5', 'SUMMARY'],
                      ['C5', ''],
                      ['D5', ''],
                      ['E5', ''],
                      ['F5', ''],
                      ['G5', ''],
                      ['H5', ''],
                      ['I5', ''],
                      ['J5', ''],
                      ['B10:D10', 'TOP 10 METRO REGIONS'],
                      ['C10', ''],
                      ['D10', ''],
                      ['H10:J10', 'RATINGS'],
                      ['I10', ''],
                      ['J10', '']]:
                ##TODO fix issue with merged cells
                # ws.merge_range(i[0], i[1], section_heading)
                ws.write(i[0], i[1], section_heading)
            for i in (['B', 'Total Registrants', "df"],
                      ['D', 'ENG HH Donor', "df[df['EngHHGiving'] > 0]"],
                      ['F', 'ENG Alumni', "df[df['ENGDegreeDeptsConcat'].notnull()]"],
                      ['H', 'Rated $25K+', "df[df['Rating'] < 'M']"],
                      ['J', 'Under PM', "df[df['PLANMANAGERPLANTYPE'].notnull()]"]
            ):
                    ws.write(i[0]+'7', i[1], summary_heading)
                    ws.write(i[0]+'8', len(eval(i[2])), summary_content)
            locale = df.pivot_table(index=['METROREGION'], values=['PROSPECTNAME'], aggfunc=np.count_nonzero)
            locale = locale.sort('PROSPECTNAME', ascending=False)[:10]
            ratings = df.pivot_table(index=['Rating'], values=['PROSPECTNAME'], aggfunc=np.count_nonzero)
            for i in range(len(locale)):
                ws.write('B'+str(i+11), locale.index[i], table_contents)
                ws.write('D'+str(i+11), int(locale.values[i][0]), table_contents)
            for i in range(len(ratings)):
                ws.write('H'+str(i+11), ratings.index[i], table_contents)
                ws.write('J'+str(i+11), int(ratings.values[i]), table_contents)
            ##TODO conditional formatting for table rows
            # for i in ('B13:D22', 'H13:J' + str(len(df['Rating'].unique()) + 12)):
            #     ws.conditional_format(i, {'type': 'formula',
            #                               'criteria': '=MOD(ROW(),2)=0)',
            #                               'format': wb.add_format({'bg_color': '#F2F2F2'})})
            df.drop(['METROREGION'], inplace=True, axis=1)

        # write full list
        df.to_excel(writer, sheet_name='Full List', index=False, startrow=7)

        # write pivot table sheets
        if format_type == 'UIF':
            by_pm.to_excel(writer, sheet_name='Pivot Tables - PM')
            by_dept.to_excel(writer, sheet_name='Pivot Tables - Dept')

        # write department and MGO sheets
        elif format_type == 'ENG':
            for name in mgo:
                temp_mgo = mgmt_df[mgmt_df['Management'].str.contains(name)]
                if not temp_mgo.empty:
                    temp_mgo.to_excel(writer, sheet_name=name, index=False, startrow=7)
            dept_list = set(limit_df['ENGDegreeDeptsConcat'].unique())
            dept_list.discard(np.nan)
            for d in dept_list.copy():
                dept_list.remove(d)
                dept_list.update(d.split('|'))
            for dept in dept_list:
                temp_dept = limit_df[limit_df['ENGDegreeDeptsConcat'].str.contains(dept)]
                temp_dept.to_excel(writer, sheet_name=dept, index=False, startrow=7)

        # set column widths
        sheetlist = list(wb.sheetnames.keys())
        if format_type != 'Event Flag':
            sheetlist.remove('Summary Stats')
        if format_type == 'UIF':
            sheetlist.remove('Pivot Tables - PM')
            sheetlist.remove('Pivot Tables - Dept')
        for sheet in sheetlist:
            ws = writer.sheets[sheet]
            conf = pd.Series(["** C O N F I D E N T I A L   I N F O R M A T I O N **",
                              "This information is proprietary, privileged, and confidential. The disclosure of this information would cause competitive",
                              "harm to the Foundation and/or Alumni Association and any unauthorized disclosure or distribution is prohibited."])
            for i in range(0, len(conf)):
                ws.write(i, 4, conf[i], wb.add_format({'align': 'center'}))
            ws.write('A6', dest_file[:-5], wb.add_format({'bold': True, 'font_size': 12}))
            for c in ['A:A', 'C:D', 'G:I']:
                ws.set_column(c, 10)
            for c in ['B:B', 'E:F', 'J:M', 'O:Q']:
                ws.set_column(c, 20)
            ws.set_column('N:N', 15)
            ws.set_column('R:S', 5)

        if format_type == 'UIF':
            # Compare to previous file, if applicable
            if 'Old' in df.columns:
                if len(df['Old'].unique()) > 1:
                    ws = writer.sheets['Full List']
                    ws.conditional_format('A9:Q' + str(len(df.index) + 8),
                                          {'type': 'formula', 'criteria': '=ISBLANK($R9)',
                                           'format': wb.add_format({'bg_color': '#FFFF00'})})
                ws.set_column('R:R', None, None, {'hidden': 1})
            # format Pivot Tables
            for sheet in ['Pivot Tables - PM', 'Pivot Tables - Dept']:
                ws = writer.sheets[sheet]
                ws.set_column('A:B', 25)
                ws.set_column('C:E', 10)
        elif format_type == 'ENG':
            # have Summary Stats and Full List at beginning, sort all other sheets alphabetically
            sheetlist.append('Summary Stats')
            sheetlist.sort()
            for i in ('Full List', 'Summary Stats'):
                s = sheetlist.index(i)
                sheetlist.insert(0, sheetlist.pop(s))
            wb.worksheets_objs.sort(key=lambda x: sheetlist.index(x.name))

        # explicitly save and implicitly close file
        writer.save()
