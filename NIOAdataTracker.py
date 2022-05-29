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
ws.append(['NIA Patients MRN and Count', 'MesCoBraD enrolled MRN and Count',
          'CBTI enrolled MRN and Count', 'MesCoBrad and CBTI MRN and count', 'NIA- Not in study- MRN and Count'])


# dictionary for keeping track of whhat has been completed for each patient care/research pathway

completed_appts_qs = {'MesCoBraD enrolled MRN and Count': {'MRN': {}}, 'CBTI enrolled MRN and Count': {
    'MRN': {}}, 'MesCoBraD and CBTI MRN and Count': {'MRN': {}}, 'NIA_total': {'MRN': {}}, 'NIA- Not in study- MRN and Count': {'MRN': {}}}


mescobrad_mrn_list = completed_appts_qs['MesCoBraD enrolled MRN and Count']['MRN'].keys(
)
cbti_mrn_list = completed_appts_qs['CBTI enrolled MRN and Count']['MRN'].keys()
nia_mrn_list = completed_appts_qs['NIA_total']['MRN'].keys()
nia_no_study = completed_appts_qs['NIA- Not in study- MRN and Count']['MRN'].keys()
nia_both_study = completed_appts_qs['MesCoBraD and CBTI MRN and Count']['MRN'].keys(
)


def NIA_patient_stats():
    ''' # of NIA patients based on demographics sheet'''

    ws['A2'] = 0
    mrn_last_value = 0
    count = 0
    for mrn in ws_nia.rows:
        if mrn[0].value == 'MRN':
            pass
        else:
            mrn_value = mrn[0].value
            if mrn_value != mrn_last_value:
                count += 1
                ws.append([mrn_value])
                completed_appts_qs['NIA_total']['MRN'][mrn_value] = []

                mrn_last_value = mrn_value
        ws['A2'].value = count - 1
    # nia_mrn_list = completed_appts_qs['NIA_total']['completed']['MRN']
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
                # ws['B'] + (mrn)
                mrn_last_value = mrn_value
                completed_appts_qs['MesCoBraD enrolled MRN and Count']['MRN'][mrn_value] = [
                ]
                next_open_cell = 'B' + str(len(
                    completed_appts_qs['MesCoBraD enrolled MRN and Count']['MRN'].keys()) + 2)
                ws[next_open_cell] = mrn_value
        ws['B2'].value = count - 1
    # mescobrad_mrn_list = completed_appts_qs['MesCoBraD enrolled MRN and Count']['MRN'][0:-1]
    # print('MRN list MesCoBraD',
    #       mescobrad_mrn_list)
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
                completed_appts_qs['CBTI enrolled MRN and Count']['MRN'][mrn_value] = mrn_value
                next_open_cell = 'C' + str(len(
                    completed_appts_qs['CBTI enrolled MRN and Count']['MRN'].keys())+2)
                # print('next open cell', next_open_cell)
                ws[next_open_cell] = mrn_value
        ws['C2'].value = count - 1
        cbti_enrolled_mrn = completed_appts_qs['CBTI enrolled MRN and Count']['MRN'].keys(
        )
    # print('CBTI List of MRN patients enrolled', cbti_enrolled_mrn)
    print('# of patients enrolled in CBTI', len(cbti_enrolled_mrn))
    return


def pts_CBTI_and_MesCoBrad():
    '''MRN's of patients in both CBTI and MesCoBraD'''

    ws['D2'] = 0
    count = 0
    for mrn in completed_appts_qs['NIA_total']['MRN'].keys():
        if mrn in completed_appts_qs['CBTI enrolled MRN and Count']['MRN'].keys():
            if mrn in completed_appts_qs['MesCoBraD enrolled MRN and Count']['MRN'].keys():
                completed_appts_qs['MesCoBraD and CBTI MRN and Count']['MRN'][mrn] = [
                ]
                next_open_cell = 'D' + \
                    str(len(
                        completed_appts_qs['MesCoBraD and CBTI MRN and Count']['MRN'].keys())+2)
                ws[next_open_cell] = mrn
                count += 1
    ws['D2'] = count - 1
    print('# of patients in both studies = ', count)
    return


def nia_not_in_study():
    '''list of MRN for NIA patients not in any study'''
    nia_no_study = completed_appts_qs['NIA- Not in study- MRN and Count']['MRN'].keys()

    ws['E2'] = 0
    count = 0
    for mrn in completed_appts_qs['NIA_total']['MRN'].keys():
        if mrn in completed_appts_qs['CBTI enrolled MRN and Count']['MRN'].keys():
            pass
        elif mrn in completed_appts_qs['MesCoBraD enrolled MRN and Count']['MRN'].keys():
            pass
        else:
            completed_appts_qs['NIA- Not in study- MRN and Count']['MRN'][mrn] = []
            next_open_cell = 'E' + str(len(nia_no_study) + 2)
            # print('next open E cell', next_open_cell)
            ws[next_open_cell] = mrn
            count += 1
        ws['E2'] = count
    # print('list of MRN for NIA not in any study =', nia_no_study)
    # print('count of patients in no study for NIA =', len(nia_no_study))
    return


NIA_patient_stats()
MesCoBraD_enrolled()
CBTI_enrolled()
pts_CBTI_and_MesCoBrad()
nia_not_in_study()
wb.save('statistics.xlsx')
