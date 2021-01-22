import numpy as numericLibrary
import pandas as spreadsheetLibrary

from datetime import datetime

reportingTime = 0
recordingDate = 0
recordingAverage = 0
recordingPercentile90 = 0
recordingPercentile75 = 0
recordingPercentile50 = 0

controlledGlucose = False

controlledRecords = []
uncontrolledRecords = []

controlledPercentage = 0
uncontrolledPercentage = 0

spreadsheetEntries = spreadsheetLibrary.read_excel('glucoseReadings.xlsx')
spreadsheetRecords = spreadsheetEntries.to_numpy()

reportingTime = datetime.strftime(datetime.now(), "%d/%m/%Y %H:%M:%S")
reportingFile = open("glucoseReporting.html", "w")

reportingFile.write(f'<html>\n')
reportingFile.write(f'\t<head>\n')

reportingFile.write(f'\t\t<title>Glucose Reporting @ {reportingTime}</title>\n')

reportingFile.write(f'\t\t<style>\n')
reportingFile.write(f'\t\t\t* {{color: #523547;}}\n')
reportingFile.write(f'\t\t\t* {{font-family: Abadi, Futura, Arial, Helvetica;}}\n')
reportingFile.write(f'\t\t\tbody {{background-color: #FF9966;}}\n')
reportingFile.write(f'\t\t\tth, td {{padding: 3px;}}\n')
reportingFile.write(f'\t\t\tdiv, td, th {{border: 1px solid #5E7BBD;}}\n')
reportingFile.write(f'\t\t\tdiv {{border-radius: 6px;}}\n')
reportingFile.write(f'\t\t\tdiv {{padding: 6px;}}\n')
reportingFile.write(f'\t\t\tdiv, table {{margin: 9px;}}\n')

reportingFile.write(f'\t\t\t.roundTable thead th:first-child {{border-radius: 6px 0 0 0;}}\n')
reportingFile.write(f'\t\t\t.roundTable thead th:last-child {{border-radius: 0 6px 0 0;}}\n')
reportingFile.write(f'\t\t\t.roundTable tbody tr:last-child td:first-child {{border-radius: 0 0 0 6px;}}\n')
reportingFile.write(f'\t\t\t.roundTable tbody tr:last-child td:last-child {{border-radius: 0 0 6px 0;}}\n')

reportingFile.write(f'\t\t</style>\n')

reportingFile.write(f'\t</head>\n')
reportingFile.write(f'\t<body>\n')

reportingFile.write(f'\t\t<div>\n')
reportingFile.write(f'\t\t\t<p>Report generated at {reportingTime}</p>\n')
reportingFile.write(f'\t\t</div>\n')

reportingFile.write(f'\t\t<table class="roundTable">\n')

reportingFile.write(f'\t\t<thead>\n')
reportingFile.write(f'\t\t\t<tr>\n')
reportingFile.write(f'\t\t\t\t<th>Recording Date</th>\n')
reportingFile.write(f'\t\t\t\t<th>Average</th>\n')
reportingFile.write(f'\t\t\t\t<th>90th Percentile</th>\n')
reportingFile.write(f'\t\t\t\t<th>75th Percentile</th>\n')
reportingFile.write(f'\t\t\t\t<th>50th Percentile</th>\n')
reportingFile.write(f'\t\t\t</tr>\n')
reportingFile.write(f'\t\t</thead>\n')

reportingFile.write(f'\t\t<tbody>\n')

for individualRecord in spreadsheetRecords:
    
    recordingDate = datetime.strftime(individualRecord[0], "%d/%m/%Y")
    individualRecord = numericLibrary.delete(individualRecord, 0)

    recordingAverage = numericLibrary.mean(individualRecord)
    recordingPercentile90 = numericLibrary.percentile(individualRecord, 90)
    recordingPercentile75 = numericLibrary.percentile(individualRecord, 75)
    recordingPercentile50 = numericLibrary.percentile(individualRecord, 50)  

    if recordingPercentile90 >= 7.0:
        uncontrolledRecords.append(individualRecord)
        reportingFile.write(f'\t\t\t<tr>')        
        reportingFile.write(f'<td>{recordingDate}</td><td>{recordingAverage:.2f}</td><td>{recordingPercentile90:.2f}</td><td>{recordingPercentile75:.2f}</td><td>{recordingPercentile50:.2f}</td>')
        reportingFile.write(f'</tr>\n')
    else:
        controlledGlucose = True
        controlledRecords.append(individualRecord)
        reportingFile.write(f'\t\t\t<tr>')
        reportingFile.write(f'<td>{recordingDate}</td><td>{recordingAverage:.2f}</td><td>{recordingPercentile90:.2f}</td><td>{recordingPercentile75:.2f}</td><td>{recordingPercentile50:.2f}</td>')
        reportingFile.write(f'</tr>\n')

reportingFile.write(f'\t\t</tbody>\n')
reportingFile.write(f'\t\t</table>\n')

controlledPercentage = (len(controlledRecords) / len(spreadsheetRecords) * 100)
uncontrolledPercentage = (len(uncontrolledRecords) / len(spreadsheetRecords) * 100)

reportingFile.write(f'\t\t<div>\n')
reportingFile.write(f'\t\t\t<p>The final score is {int(controlledPercentage)}/100; there were {len(uncontrolledRecords)} bad records out of {len(spreadsheetRecords)}</p>\n')
reportingFile.write(f'\t\t</div>\n')

reportingFile.write(f'\t\t<div>\n')
reportingFile.write('\t\t\t<p>© 2020 Seagull Holdings. All Rights Reserved.</p>\n')
reportingFile.write(f'\t\t</div>\n')

reportingFile.write(f'\t</body>\n')
reportingFile.write(f'\t</html>\n')
