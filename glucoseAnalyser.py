import sys as systemLibrary
import numpy as numericLibrary
import pandas as spreadsheetLibrary
import scipy.stats as scientificLibrary

from datetime import datetime
from glucoseFormatter import formatter

patientName = systemLibrary.argv[1]
patientBlood = systemLibrary.argv[2]
patientIdentifier = systemLibrary.argv[3]

reportingTime = 0
recordingDate = 0
recordingAverage = 0
recordingPercentile90 = 0
recordingPercentile75 = 0
recordingPercentile50 = 0
recordingKurtosis = 0

controlledGlucose = False

controlledRecords = []
uncontrolledRecords = []

controlledPercentage = 0
uncontrolledPercentage = 0

spreadsheetEntries = spreadsheetLibrary.read_excel('glucoseReadings.xlsx')
spreadsheetRecords = spreadsheetEntries.to_numpy()

reportingTime = datetime.strftime(datetime.now(), "%d/%m/%Y %H:%M:%S")
reportingFile = open("glucoseReporting.html", "w")

reportingFile.write(f'<html>')
reportingFile.write(formatter.newLine())

reportingFile.write(f'<head>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'<title>Glucose Reporting @ {reportingTime}</title>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'<style>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'* {{color: #523547}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'* {{font-size: 13pt}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'* {{font-family: Abadi, Futura, Arial, Helvetica}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'body {{background-color: #B2BEB5}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'div {{padding: 6px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'div {{border-style: 4px #5E7BBD}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'div {{border-style: double none double none}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'div {{border-width: 4px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'div {{border-color: #5E7BBD}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'h6 {{margin: 3px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'h6 {{text-transform: uppercase}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'ul {{list-style-type: square}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'ul {{margin: 0px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'div, table {{margin: 9px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'td, th {{border: 1px solid #5E7BBD}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'th, td {{padding: 3px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'th, td {{width: 130px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'.roundTable thead th:first-child {{border-radius: 6px 0 0 0}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'.roundTable thead th:last-child {{border-radius: 0 6px 0 0}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'.roundTable tbody tr:last-child td:first-child {{border-radius: 0 0 0 6px}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'.roundTable tbody tr:last-child td:last-child {{border-radius: 0 0 6px 0}}')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'</style>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.singleIndentation())
reportingFile.write(f'</head>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.singleIndentation())
reportingFile.write(f'<body>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'<div>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<h6>Personal Information</h6>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<ul>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'<li>Full Name: {patientName}</li>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'<li>Blood Group: {patientBlood}</li>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'<li>Health Identifier: {patientIdentifier}</li>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'</ul>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'</div>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'<table class="roundTable">')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<thead>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'<tr>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quintupleIndentation())
reportingFile.write(f'<th>Recording Date</th>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quintupleIndentation())
reportingFile.write(f'<th>Average</th>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quintupleIndentation())
reportingFile.write(f'<th>90th Percentile</th>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quintupleIndentation())
reportingFile.write(f'<th>75th Percentile</th>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quintupleIndentation())
reportingFile.write(f'<th>50th Percentile</th>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quintupleIndentation())
reportingFile.write(f'<th>Kurtosis</th>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'</tr>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'</thead>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<tbody>')
reportingFile.write(formatter.newLine())

for individualRecord in spreadsheetRecords:
    
    recordingDate = datetime.strftime(individualRecord[0], "%d/%m/%Y")
    individualRecord = numericLibrary.delete(individualRecord, 0)

    recordingAverage = numericLibrary.mean(individualRecord)
    recordingPercentile90 = numericLibrary.percentile(individualRecord, 90)
    recordingPercentile75 = numericLibrary.percentile(individualRecord, 75)
    recordingPercentile50 = numericLibrary.percentile(individualRecord, 50) 
    recordingKurtosis = scientificLibrary.kurtosis(individualRecord)

    if recordingPercentile90 >= 7.0 or recordingKurtosis >= 3.0:
        uncontrolledRecords.append(individualRecord)

        reportingFile.write(formatter.quadrupleIndentation())
        reportingFile.write(f'<tr>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingDate}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingAverage:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingPercentile90:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingPercentile75:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingPercentile50:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingKurtosis:.2f}</td>')
        reportingFile.write(formatter.newLine())

        reportingFile.write(formatter.quadrupleIndentation())
        reportingFile.write(f'</tr>')
        reportingFile.write(formatter.newLine())

    else:
        controlledGlucose = True
        controlledRecords.append(individualRecord)
        
        reportingFile.write(formatter.quadrupleIndentation())
        reportingFile.write(f'<tr>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingDate}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingAverage:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingPercentile90:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingPercentile75:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingPercentile50:.2f}</td>')
        reportingFile.write(formatter.newLine())
        
        reportingFile.write(formatter.quintupleIndentation())
        reportingFile.write(f'<td>{recordingKurtosis:.2f}</td>')
        reportingFile.write(formatter.newLine())

        reportingFile.write(formatter.quadrupleIndentation())
        reportingFile.write(f'</tr>')
        reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'</tbody>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'</table>')
reportingFile.write(formatter.newLine())

controlledPercentage = (len(controlledRecords) / len(spreadsheetRecords) * 100)
uncontrolledPercentage = (len(uncontrolledRecords) / len(spreadsheetRecords) * 100)

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'<div>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<h6>Observations</h6>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<ul>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'<li>The final score is {int(controlledPercentage)}/100</li>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'<li>There were {len(uncontrolledRecords)} bad records out of {len(spreadsheetRecords)}</li>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'</ul>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'</div>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'<div>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<h6>Notes</h6>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'<ul>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write(f'<li>Report Generated {reportingTime}</li>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.quadrupleIndentation())
reportingFile.write('<li>© 2020-2021 Seagull Holdings. All Rights Reserved.</li>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.tripleIndentation())
reportingFile.write(f'</ul>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.doubleIndentation())
reportingFile.write(f'</div>')
reportingFile.write(formatter.newLine())

reportingFile.write(formatter.singleIndentation())
reportingFile.write(f'</body>')
reportingFile.write(formatter.newLine())

reportingFile.write(f'</html>')
reportingFile.write(formatter.newLine())
