# pip3 install openpyxl

from operator import indexOf
from openpyxl import Workbook, load_workbook

# the study stats (both MesCoBraD and CBTI)
wb_ss = load_workbook('study_stats.xlsx')
# accessing values from 'Data type' sheet
ws_ss = wb_ss.active
# accessing values from 'Sheet2' sheet to analyze
wb_ss_s2 = wb_ss['Sheet2']
# accessing the values from sheet 'CBTI- SCHEDULE (2)'
ws3 = wb_ss['CBTI- SCHEDULE (2)']


# print(wb_ss.sheetnames)

# NIA demographics stats
wb_nia = load_workbook('NIA Demographics.xlsx')
ws_nia = wb_nia.active

# print('MRN list from study_stats data type sheet', ws_ss['A1'].value)
# new workbook for displaying statistics for NIA and study patients
wb = Workbook()
ws = wb.active
ws.title = 'Statistics'
ws.append(['NIA Patients MRN and Count', 'MesCoBraD Enrolled',
          'CBTI Enrolled', 'MesCoBraD and CBTI Enrolled', 'NIA- Not in study'])


# dictionary for keeping track of whhat has been completed for each patient care/research pathway

enrolled = {'MesCoBraD Enrolled': {'MRN': []}, 'CBTI Enrolled': {
    'MRN': []}, 'MesCoBraD and CBTI Enrolled': {'MRN': []}, 'NIA_total': {'MRN': []}, 'NIA- Not in study': {'MRN': []}}


# completed_appointments = {'MesCoBraD': {'MRN': 'NPT': '', 'PSG': [], 'Actigraphy': '', '1YNPT': ''}, 'CBTI': {
#     'MRN': '', 'completed_questionnaires': {'consent': '', 'cbs': '', 'cCbs': '', 'q-ess-a': '', 'ess-a(3m)': '', 'ess-a(1y)': '', 'q-isi-a': '', 'isi-a(3m)': '', 'isi-a(1y)': '', 'q-psqi': '', 'psqi(3m)': '', 'psqi(1y)': '', 'qnpi-q': '', 'q-cdr': ''}, 'arm': ''}}, 'MesCoBraD and CBTI Enrolled': {'MRN': []}, 'NIA_total': {'MRN': {}}, 'NIA- Not in study': {'MRN': {}}}


mescobrad_mrn_list = enrolled['MesCoBraD Enrolled'].keys()
cbti_mrn_list = enrolled['CBTI Enrolled'].keys()
nia_mrn_list = enrolled['NIA_total'].keys()
nia_no_study = enrolled['NIA- Not in study'].keys()
nia_both_study = enrolled['MesCoBraD and CBTI Enrolled'].keys()


def NIA_patient_stats():
    ''' # of NIA patients based on demographics sheet'''

    ws['A2'] = 0
    mrn_last_value = 0
    count = 0
    for mrn in ws_nia.rows:
        mrn_value = mrn[0].value
        if mrn_value == 'MRN':
            pass
        else:
            if mrn_value != mrn_last_value:
                count += 1
                ws.append([mrn_value])
                enrolled['NIA_total']['MRN'].append(mrn_value)
                mrn_last_value = mrn_value
    ws['A2'].value = count
    # nia_mrn_list = enrolled['NIA_total']['completed']['MRN']
    # print('NIA MRN List = ', nia_mrn_list)
    return


def MesCoBraD_enrolled():
    '''MesCoBraD # of enrolled'''
    ws['B2'] = 0
    mrn_last_value = 0
    count = 0
    for mrn in ws_ss.rows:
        if mrn[0].value == 'MRN':
            pass
        else:
            mrn_value = mrn[0].value
            if mrn_value != mrn_last_value:
                count += 1
                mrn_last_value = mrn_value
                enrolled['MesCoBraD Enrolled']['MRN'].append(mrn_value)
                next_open_cell = 'B' + str(len(
                    enrolled['MesCoBraD Enrolled']['MRN'])+2)
                ws[next_open_cell] = mrn_value
    ws['B2'].value = count
    # print('MRN count MesCoBraD', len(mescobrad_mrn_list))
    return wb.save('statistics.xlsx')


def CBTI_enrolled():
    '''function for CBTI enrolled count'''

    mrn_last_value = 0
    count = 0
    for mrn in ws3.rows:
        if mrn[0].value == 'MRN':
            pass
        else:
            mrn_value = mrn[0].value
            if mrn_value != mrn_last_value:
                count += 1
                mrn_last_value = mrn_value
                enrolled['CBTI Enrolled']['MRN'].append(mrn_value)
                next_open_cell = 'C' + str(len(
                    enrolled['CBTI Enrolled']['MRN'])+2)
                ws[next_open_cell] = mrn_value
    ws['C2'].value = count
    cbti_enrolled_mrn = enrolled['CBTI Enrolled']['MRN']
    print('# of patients enrolled in CBTI', len(cbti_enrolled_mrn))
    return


def pts_CBTI_and_MesCoBrad():
    '''MRN's of patients in both CBTI and MesCoBraD'''

    ws['D2'] = 0
    count = 0
    for mrn in enrolled['NIA_total']['MRN']:
        if mrn in enrolled['CBTI Enrolled']['MRN']:
            if mrn in enrolled['MesCoBraD Enrolled']['MRN']:
                enrolled['MesCoBraD and CBTI Enrolled']['MRN'].append(
                    mrn)
                next_open_cell = 'D' + \
                    str(len(
                        enrolled['MesCoBraD and CBTI Enrolled']['MRN'])+2)
                ws[next_open_cell] = mrn
                count += 1
    ws['D2'] = count
    print('# of patients in both studies = ', count)
    return


def nia_not_in_study():
    '''list of MRN for NIA patients not in any study'''
    nia_no_study = enrolled['NIA- Not in study']['MRN']

    ws['E2'] = 0
    count = 0
    for mrn in enrolled['NIA_total']['MRN']:
        if mrn in enrolled['CBTI Enrolled']['MRN']:
            pass
        elif mrn in enrolled['MesCoBraD Enrolled']['MRN']:
            pass
        else:
            enrolled['NIA- Not in study']['MRN'].append(mrn)
            next_open_cell = 'E' + str(len(nia_no_study)+2)
            # print('next open E cell', next_open_cell)
            ws[next_open_cell] = mrn
            count += 1
    ws['E2'] = count
    # print('list of MRN for NIA not in any study =', nia_no_study)
    # print('count of patients in no study for NIA =', len(nia_no_study))
    return


def completed_appointments():
    return


NIA_patient_stats()
MesCoBraD_enrolled()
CBTI_enrolled()
pts_CBTI_and_MesCoBrad()
nia_not_in_study()
wb.save('statistics.xlsx')
