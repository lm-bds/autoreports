from string import capwords
import pandas as pd
import datetime
import PySimpleGUI as sg
import os
path = ".\\"
filelist = list(os.listdir(path))

layout = [
          [sg.Text(size=(40,1), key='-OUTPUT-')],
          [sg.T("Target_file",pad=((3,0),0))],[sg.OptionMenu(values = (filelist), key='-target-',size=(20, 5))],
          [sg.Text("Month")],[sg.Input(key='-month-',size=(20, 5))],
          [sg.Text("Year")],[sg.Input(key='-year-',size=(20, 5))],
          [sg.Text(size=(40,1), key='-OUTPUT2-')],
          [sg.Button('Run')]]


window = sg.Window('Report Creation: Monthly', layout)


while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Run':
        break


window.close()

target = values['-target-']
month = values['-month-']
year = str(2022) #values['-year-']
reportmonth = month+" "+year
df = pd.read_csv(target)
df["Time_to_clear"]= df.Time_to_clear.apply(lambda x: x[7:])
output_file = "Monthly Report " + reportmonth+".tex"

def average_time(df):
    timedf = df
    print(timedf['Time_to_clear'])
    timedf["Time_to_clear"]=pd.to_timedelta(timedf["Time_to_clear"])
    average_time =str(timedf.Time_to_clear.mean())[7:-10]
    print(average_time)
    return average_time

def escalation_num_counter(df):
    escdf = df
    step_of_escalations = len(escdf['EscalationNo'].unique())
    return step_of_escalations

def callcounter(df):
    palldf = df
    num_of_calls = palldf.Time_to_clear.count()
    return num_of_calls

def escalation_step_counter(df):
    escdf = df
    step_of_escalations = escdf['EscalationNo'].value_counts()
    step_of_escalations = step_of_escalations.rename("Steps taken to Escalate")
    return step_of_escalations

def doc_count(df):
    docdf = df
    docsum = docdf.groupby("Person")["EscalationNo"].nunique()
    docsum = docsum.rename("Number of Escalations")
    return docsum

def service_count(df):
    serdf = df
    sumserv = serdf.groupby("Service")["EscalationNo"].nunique()
    sumserv = sumserv.rename("Calls")
    sumserv.drop('N/A',axis=0, inplace = True, errors ="ignore")
    return sumserv


print(df)


df_nodes = df.drop('Description',axis =1)
descdf = df[["EscalationNo", 'Description']]

month_data_frame = df_nodes.to_latex(index=False)
doctor_frame = doc_count(df).to_latex()
service_frame = service_count(df).to_latex()
escalation_frame = escalation_step_counter(df).to_latex()
description_frame = descdf.to_latex(index=False)
avg_time_str = average_time(df)

escalation_num = escalation_num_counter(df)
num_of_calls = callcounter(df)

caps = df.Service.value_counts().CAPS
naps = df.Service.value_counts().NAPS
saps = df.Service.value_counts().SAPS



tex_output = """
\\documentclass[10pt]{article}

\\title{Monthly Report}

\\author{Scientific Team}

\\date{""" + reportmonth + """}

\\usepackage{graphicx}
\\usepackage{blindtext}
\\usepackage{multicol}
\\usepackage{pgfplots}
\\pgfplotsset{width=16.6cm,compat=1.7}
\\usepackage{lipsum}
\\usepackage{placeins}
\\usepackage{booktabs}
\\usepackage[figuresleft]{rotating}
\\AddToHook{shipout/background}{\put (0in,-\paperheight){\includegraphics[width=\paperwidth,height=\paperheight]{logo.jpg}}}

\\begin{document}



\\begin{titlepage}
\\maketitle



\\vspace{5cm}
\\begin{multicols}{2}

This report contains the monthly statistics for the Service line established May 2020.
The support line is hosted by the reporting institution in partnership with the organisation the Commission on Excellence and Innovation in Health,
the support line provides timely and direct access to advice and support from a specialist for professionals across country it is intended for use by:
This is the monthly report for the Hotline underpinned by the work of XXX.
\\begin{itemize}
    \item{example}
    \item{example}
    \item{example}
    \item{example}
    \item{example}
\\end{itemize}
Access to specialist care advice and support can enable professional's to care for people in their homes or facilities rather than having to do something.
A single phone number for professionals to use will improve access to timely support if they are unsure of the correct service to contact.


\\end{multicols}
\\end{titlepage}

\\tableofcontents
\\FloatBarrier
\\begin{figure}

\\section{Call volumes attributed to LHNs}

LHNs:

The support line is available for all places.
Figure {1} displays the places from which the calls originate.

\\begin{tikzpicture}
\\hspace{-2cm}
\\centering
\\begin{axis}
[
    ybar,
    ylabel={\#Number of Calls}, % the ylabel must precede a # symbol.
    xlabel={\ placess},
    symbolic x coords={}, % these are the specification of coordinates on the x-axis.
    xtick=data,
    xticklabel style={rotate=90},
     nodes near coords, % this command is used to mention the y-axis points on the top of the particular bar.
    nodes near coords align={vertical},
    ]
\\addplot coordinates {};

\\end{axis}

\\end{tikzpicture}

\\caption{} \label{fig:1}



\\end{figure}





\\begin{figure}

\\section{Call volumes by Metropolitan service}

Figure {2} displays the calls recieved by metropolitan service.


\\begin{tikzpicture}
\\hspace{-2cm}
\\begin{axis}
[
    ybar,
    ylabel={\#Number of Calls}, % the ylabel must precede a # symbol.
    xlabel={},
    symbolic x coords={}, % these are the specification of coordinates on the x-axis.
    xtick=data,
     nodes near coords, % this command is used to mention the y-axis points on the top of the particular bar.
    nodes near coords align={vertical},
    ]
\\addplot coordinates {};

\\end{axis}
\\end{tikzpicture}
 \\caption{Calls by Metropolitan Service} \label{fig:2}
\\end{figure}



\\FloatBarrier

\\begin{figure}
\\section{Summary Statistics}

There was a total of """+str(num_of_calls)+""" calls this month.

Of these calls """ +str(escalation_num) +""" were Escalations.


The average response time for an escalation was """+avg_time_str+""" hours


The below tables summarize the data collected for the month


\centering


\\vspace{0.5cm}
""" +  str(description_frame) +"""
\\caption{A log of how each call was connected by the paging service} \label{fig:3}
\\end{figure}


\\begin{figure}
\centering
"""+   str(doctor_frame) + """
\\caption{} \label{fig:4}
\\end{figure}

\\FloatBarrier

\\begin{figure}
\centering
\\vspace{0.5cm}
""" +  str(service_frame) +"""
\\caption{} \label{fig:5}
\\end{figure}



\\begin{figure}
\centering
\\vspace{0.5cm}
""" +  str(escalation_frame) +"""
\\caption{Steps taken to connect each call} \label{fig:6}
\\end{figure}







\\FloatBarrier
\\pagebreak
\\section{Monthly Data Table}
\\vspace{2cm}
\\hspace{-2cm}
"""+ str(month_data_frame) +"""



\\end{document}
"""
with open(output_file,'w') as f:
    f.write(tex_output)
