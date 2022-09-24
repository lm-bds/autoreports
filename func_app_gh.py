
import pandas as pd
import datetime
import PySimpleGUI as sg
import os

##create the values to select from - the active directory
path = ".\\"
filelist = list(os.listdir(path))
###

###window
layout = [
          [sg.Text(size=(40,1), key='-OUTPUT-')],
          [sg.T("Target_file",pad=((3,0),0))],[sg.OptionMenu(values = (filelist), key='-target-',size=(20, 5))],
          [sg.T("Monthly_file",pad=((3,0),0))],[sg.OptionMenu(values = (filelist), key='-output-',size=(20, 5))],
          [sg.Text("Weekly_report")],[sg.Input(key='-weekly-',size=(50, 20))],
          [sg.T("Headers? (should be 0 unless first week of month)",pad=((3,0),0))],[sg.OptionMenu(values = (0, 1), key='-head-',size=(20, 5))],
	  [sg.T("Number of commas to split (should be 2)",pad=((3,0),0))],[sg.OptionMenu(values = (1,2,3,4,5,6,7,8,9,10), key='-delimiter-',size=(20, 5))],
          [sg.Text(size=(40,1), key='-OUTPUT2-')],
          [sg.Button('Run')]]

window = sg.Window('Report Creation: Monthly Palliative Care', layout)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Run':
        break

window.close()
##

###Capture values from the window - head variable issue with intelligenability for end user
target = values['-target-']

df = pd.read_csv(target, skiprows = 2)
drdf = pd.read_csv("Pal_care_names_to_IDs.csv")

monthly = values['-output-']
weekly = values['-weekly-']+'.csv'
head = bool(int(values['-head-']))
delimiter_num = int(values['-delimiter-'])

def xlookup(lookup_value, lookup_array, return_array, if_not_found:str = ''):
    match_value = return_array.loc[lookup_array.str.contains(lookup_value[:7])]
    if match_value.empty:
        return f'"{lookup_value}" not found!' if if_not_found == '' else if_not_found

    else:
        return match_value.tolist()[0]

def csv_clean():
    ### get data file and top and tail it
    dfchop = df
    dfchop = dfchop.drop(dfchop.index[-1])
    dfchop = dfchop.drop(dfchop.index[-1])
    dfchop = dfchop.loc[~dfchop['Message'].str.contains('REPLY TO')]
    return dfchop

###create the correct format for reporting for the palliative care report
def timedelta():
    palldf = csv_clean()
    ###average answer time
    palldf["TimeRecvd"]=pd.to_timedelta(palldf["TimeRecvd"])
    palldf["ClearedTime"]=pd.to_timedelta(palldf["ClearedTime"])
    ##

    palldf["Time_to_clear"]=palldf["ClearedTime"]-palldf["TimeRecvd"]
    ###split up to message column
    return palldf

def locations():
    drsdf = drdf
    palldf = csv_clean()
    palldf.Step = palldf.Step.fillna(' ')
    palldf['Service'] = palldf.Step.apply(lambda x: 'CAPS' if 'cen_' in x else ('SAPS' if 'sou_' in x else ('NAPS' if 'nor_' in x else 'N/A')))
    palldf['People'] = palldf.Step.apply(xlookup, args = (drsdf['ID'],drsdf['Name']))
    return palldf

def palliativeformat():
    palldf = locations()
    mess_form_df=palldf["Message"].str.split(expand=True,pat=',',n=delimiter_num)
    mess_form_df = mess_form_df.drop(columns=0)
    #mess_form_df[2] = mess_form_df[2].str.removeprefix("CONNECTED to ")
    #mess_form_df[2] = mess_form_df[2]

    mess_form_df = mess_form_df.rename(columns={1: "EscalationNo",2:"Description"})
    mess_form_df.EscalationNo = mess_form_df.EscalationNo.apply(lambda x: '' if 'Ph:' in x else x)
    return mess_form_df

def average_time():
    timedf = timedelta()
    average_time =str(timedf.Time_to_clear.mean())
    return average_time

def callcounter():
    palldf = timedelta()
    mess_form_df = palliativeformat()
    num_of_calls = palldf.Time_to_clear.count()
    return num_of_calls

def escalation_num_counter():
    escdf = palliativeformat()
    step_of_escalations = len(escdf['EscalationNo'].unique())
    return step_of_escalations

def escalation_step_counter():
    escdf = palliativeformat()
    step_of_escalations = escdf['EscalationNo'].value_counts()
    step_of_escalations = step_of_escalations.rename("Steps taken to Escalate")
    return step_of_escalations


def doc_count():
    docdf = concat()
    docdf = docdf.loc[~docdf['Doctor'].str.contains('not found!')]
    docsum = docdf.groupby("Doctor")["EscalationNo"].nunique()
    return docsum

def service_count():
    serdf = concat()
    serdf = serdf.loc[~serdf['Service'].str.contains('NaN')]
    sumserv = serdf.groupby("Service")["EscalationNo"].nunique()
    sumserv = sumserv.rename("Service Count")
    sumserv.drop('N/A',axis=0, inplace = True, errors ="ignore")
    return sumserv


def concat():
    palldf = locations()
    mess_form_df = palliativeformat()
    time_df = timedelta()
    cat_df = pd.concat([mess_form_df,palldf,time_df['Time_to_clear']], axis=1)
    # print(palldf)
    # print(output_df)
    return cat_df


def final_format():
    output_df = concat()
    output_df.Doctor = output_df.Doctor.replace("""" " not found!""","N/A")
    output_df = output_df.drop(columns=['Pgd', 'Confirmed','ClearedBy', 'TakenBy', 'ClearedBy', 'Step', 'Number','Message'])
    output_df = output_df[output_df.columns[[0,2,3,4,5,6,7,1]]]
    return output_df
##send table to Monthly csv
final_format().to_csv(monthly, index = False, mode = 'a', header = head)

# ###create weekly report
with open(weekly, 'w') as file:
    file.write(str(final_format().to_csv(index= False)))
    file.write('\n'+'\n')
    file.write(str(doc_count().to_csv()))
    file.write('\n')
    file.write(str(service_count().to_csv()))
    file.write('\n')
    file.write(str(escalation_step_counter().to_csv()))
    file.write('The average time for service delivery upon call was:\n')
    file.write(str(average_time())+"\n")
    file.write('The Number of Calls was:\n')
    file.write(str(callcounter())+'\n')
    file.write('The Number of Escalations was:\n')
    file.write(str(escalation_num_counter()))
