import os, sys, subprocess
subprocess.call(['pip', 'install', '--upgrade','pip'])
subprocess.call(['pip', 'install', 'selenium'])
subprocess.call(['pip', 'install', 'openpyxl'])

from selenium import webdriver as wd
from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Color, PatternFill, Font, Border, NamedStyle
from openpyxl.styles import Alignment, Side
from openpyxl.formatting import Rule
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter, rows_from_range
from openpyxl.utils import units
from pprint import pprint
import argparse


##==============================================================================
##      NO PY FILE EDITING NEEDED
##
##      The reports will be saved as an XLSX file. Both reports will be made
##      after the script runs.
##
##      1   Login to the VPN
##      2   Execute the runner.sh script
##      2.1      $bash runner.sh
##        OR
##      2.1 give exec permission and run using $./runner.sh
##==============================================================================
##       (WINDOWS)
##      A.1 on commandline run C:\>  python3 -m venv myenv
##      A.2 C:\> <path to script>\Scripts\activate.bat
##      A.3 C:\> <path to script> python3 'ccp_daily_automation.py' --smoke
##        OR
##      A.3 C:\> <path to script> python3 'ccp_daily_automation.py' --regression
##==============================================================================

#Global Values
LINE_WRAP_LENGTH    = 60
DURATION_COL_LENGTH = 18
DURATION_COL        = 5
FAIL_SC_COLS        = 6
FAIL_ST_COLS        = 7
automationReport    = ''
reportFileName      = ''
suite               = {}
devices             = []
nodes               = []

def init():
    '''process CLI arguments'''
    global automationReport 
    global reportFileName
    global suite
    global devices
    global nodes

    parser = argparse.ArgumentParser(description='Scrape Jenkins report' +
                                                 ' and create XLSX report.')
    parser.add_argument('-s', '--smoke', help="Generate Smoke Report",
                        action="store_true")
    parser.add_argument('-r','--regression', help="Generate Regression Report",
                        action="store_true")
    args = parser.parse_args()
    pprint(args)

    if args.smoke: 
        automationReport = 'RCO Smoke Tests'
        reportFileName = 'RCO_Smoke_Report'
    elif args.regression:
        automationReport = 'RCO Regression Tests'
        reportFileName = 'RCO_Smokeless_Regression_Report'

    suite = {'nodes': [], 'duration': '', 'total': 0, 'failed': 0,
             'name': automationReport}
    pprint(suite)
    devices = ['Desktop', 'Tablet', 'Mobile']
    nodes = ['Authentication', 'Confirmation', 'Delivery', 'Payment', 'Review']
    #nodes = ['Confirmation']

def open_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    else:
        opener ="open" if sys.platform == "darwin" else "xdg-open"
        subprocess.call([opener, filename])

def getFeatureFileName(string):
    if string[0] == '(':
        cutoff = 0
        for char in string:
            if char == ')':
                return string[1:cutoff]
            cutoff += 1

def remUnderscore(string):
    final = ''
    for char in string:
        if char != '_':
            final += char
    return final

def cellLineBreak(stringList):
    if len(stringList) == 0:
        return ''
    if len(stringList) == 1:
        return stringList[0]
    final = '"' + stringList[0]
    for i in range(len(stringList)):
        if i == 0:
            continue
        final += '\n' + stringList[i]
    return final + '"'

def duration(string):
    draft = ''
    digits = '1234567890'
    letters = 'hmis'
    for char in string:
        if char in digits:
            draft += char
        if char in letters:
            draft += char
    final = ''
    lastChar = ''
    for char in draft:
        if char == 's':
            if lastChar == 'm':
                final += char
            if lastChar in digits:
                final += char
        elif char in digits:
            if lastChar not in digits:
                final += ' ' + char
            else:
                final += char
        else:
            final += char
        lastChar = char
    return final

def mergeCells(ws, cell, length):
    mergeStart = cell.row
    mergeEnd   = mergeStart + length
    for col in range(1,6):
        ws.merge_cells(start_row = mergeStart,
                start_column     = col,
                end_row          = mergeEnd,
                end_column       = col)

def mergeSheet(ws):
    mergeLen   = 0
    newGroup   = False
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=1):
        for cell in row:
            if cell.value and not newGroup:
                start = cell
            elif cell.value and newGroup:
                mergeCells(ws, start, mergeLen)
                mergeLen = 0
                start    = cell
                newGroup = False
            elif not cell.value:
                if cell.row < ws.max_row:
                    mergeLen += 1
                    newGroup  = True
                else:
                    mergeCells(ws, start, mergeLen + 1)

def takeAllSuites(ws):
    return [cell for row in ws.iter_rows(min_row=1,
                                         max_row=ws.max_row,
                                         max_col=1)
            for cell in row if str(cell.value).startswith('Suite:') ]

def resizeToFitColumn(ws, columnIndex):
    """ one-based index """
    width = max(list(map(len,
        [str(cell.value) for row in ws.iter_rows(min_row = 1,
                                            max_row = ws.max_row,
                                            min_col = columnIndex,
                                            max_col = columnIndex)
                                            for cell in row])))
    columnName = get_column_letter(columnIndex)
    ws.column_dimensions[columnName].width = width

def resizeRow(ws, rowIndex, multiplier):
    """ one-based index """
    size = units.DEFAULT_ROW_HEIGHT * multiplier
    ws.row_dimensions[rowIndex].height = size

def resizeColumn(ws, columnIndex, width):
    """ one-based index """
    columnName = get_column_letter(columnIndex)
    ws.column_dimensions[columnName].width = width

#Red Fill
redFill = PatternFill(start_color = colors.RED,
                      end_color   = colors.RED,
                      fill_type   = 'solid')

#Blue Fill
blueFill = PatternFill(start_color = colors.BLUE,
                       end_color   = colors.BLUE,
                       fill_type   = 'solid')
#Green Fill
greenFill = PatternFill(start_color = colors.GREEN,
                          end_color = colors.GREEN,
                          fill_type = 'solid')

#vertical alignment centered 
vertCenterAl =  Alignment(vertical="center")

#horizontal centered alignment
horCenterAl =  Alignment(horizontal="center")

#centered everything
centerAl =  Alignment(vertical="center", horizontal="center")

#thin border
bd         = Side(style='thin', color="000000")
thinBorder = Border(left=bd, top=bd, right=bd, bottom=bd)

#all cells
defaultStyle           = NamedStyle(name="default")
defaultStyle.alignment = vertCenterAl
defaultStyle.border    = thinBorder


#blue highlight with white text style
blueHighlight           = NamedStyle(name="blueHighlight")
blueHighlight.font      = Font(color=colors.WHITE)
blueHighlight.alignment = centerAl
blueHighlight.border    = thinBorder
blueHighlight.fill      = blueFill

#suite header style
suiteStyle            = NamedStyle(name="Suite")
suiteStyle.font       = Font(color=colors.WHITE)
suiteStyle.alignment  = Alignment(vertical="center", horizontal="center")
suiteStyle.border     = thinBorder
suiteStyle.fill       = blueFill

#failed cells style
failedStyle            = NamedStyle(name="Failed")
failedStyle.alignment = Alignment(vertical="center", wrap_text=True)


def scrapeInfo():
    global suite
    global devices
    global nodes
    pprint(suite)
    for node in nodes:
        for device in devices:
            newNode = {'name': 'Tags=Checkout_' + node + '_Responsive_' +
                       device, 'features': [], 'duration': '', 'total': 0,
                       'failed': 0}
            suite['nodes'].append(newNode)

    driver=wd.Chrome()
    driver.get("http://jenkins.ccp.tm.tmcs/view/RCO/")
    driver.find_element_by_link_text(suite['name']).click()
    for node in suite['nodes']:
        features = []
        driver.find_element_by_link_text(node['name']).click()
        cucumberReport = driver.find_element_by_link_text('Cucumber Reports').click()
        node['duration'] = driver.find_element_by_id('stats-total-duration').text
        featureElements = driver.find_elements_by_partial_link_text('Checkout_')
        featureTexts = []
        for element in featureElements:
            featureTexts.append(element.text)
        for featureText in featureTexts:
            featureElement = driver.find_element_by_link_text(featureText)
            feature = {}
            feature['name'] = featureText
            feature['total'] = int(driver.find_element_by_id('stats-number-scenarios-' +
                                                             featureText).text)
            node['total'] = node['total'] + feature['total']
            suite['total'] = suite['total'] + feature['total']
            feature['duration'] = driver.find_element_by_id('stats-duration-' +
                                                            featureText).text
            featureElement.click()
            failures = driver.find_elements_by_xpath("//div[@class='failed']/span[@class='scenario-keyword']")
            feature['failed'] = len(failures)
            node['failed'] = node['failed'] + feature['failed']
            suite['failed'] = suite['failed'] + feature['failed']
            scenarios = driver.find_elements_by_xpath("//div[@class='failed']/span[@class='scenario-name']")
            steps = driver.find_elements_by_xpath("//div[@class='failed']/span[@class='step-name']")
            feature['failureList'] = []
            feature['failedSteps'] = []
            feature['failures'] = {}
            count = 0
            for num in range(len(failures)):
                if failures[num].text == 'Background:':
                    feature['failureList'].append(failures[num].text)
                    feature['failures'][failures[num].text] = {'scenario': failures[num].text ,
                                                               'step': steps[num].text}
                else:
                    feature['failureList'].append(failures[num].text +
                                                  scenarios[count].text)
                    feature['failures'][failures[num].text +
                                        scenarios[count].text] = {'scenario': failures[num].text +
                                                                  scenarios[count].text,
                                                                  'step': steps[num].text}
                    count += 1
                feature['failedSteps'].append(steps[num].text)
            features.append(feature)
            driver.back()
        node['features'] = features
        driver.back()
        driver.back()
    driver.quit()

def writeToTextFile(suite, reportFileName):
    reportFileName += '.tsv'
    report = open(reportFileName, 'w')
    report.write('Tests' + '\t' + 'Success Rate' + '\t' + '"#\nTests"'
                 + '\t' + '"#\nFailed"' + '\t' + 'Duration' + '\t'
                 + 'Failed Scenarios' + '\t' + 'Failed Steps' + '\n')
    report.write(suite['name'] + '\t' +
                 str(round(100.0*(suite['total']-suite['failed'])/suite['total'],
                           1)) + '%' + '\t' + str(suite['total']) + '\t' +
                 str(suite['failed']) + '\n')
    for node in suite['nodes']:
        if node['total'] == 0:
            percent = 'No Tests'
        else:
            percent = str(round(100.0*(node['total']-node['failed'])/
                                node['total']),1) + '%'
        report.write('"Suite:\n' + node['name'][5:] + '"' + '\t' + percent +
                     '\t' + str(node['total']) + '\t' + str(node['failed']) +
                     '\t' + duration(node['duration']) + '\n')
        for feature in node['features']:
            if feature['total'] == 0:
                percent = 'No Tests'
            else:
                percent = (str(round(100.0*(feature['total'] - feature['failed'])
                               / feature['total'],1)) + '%')
            report.write(getFeatureFileName(feature['name']) + '\t'
                        + percent + '\t' + str(feature['total'])
                        + '\t' + str(feature['failed']) + '\t'
                        + duration(feature['duration']) + '\t')
            count = True
            for failure in feature['failureList']:
                if count:
                    report.write(feature['failures'][failure]['scenario'] +
                                 '\t' + feature['failures'][failure]['step'] +
                                 '\n')
                    count = False
                else:
                    report.write('\t\t\t\t\t' +
                                 feature['failures'][failure]['scenario'] +
                                 '\t' + feature['failures'][failure]['step'] +
                                 '\n')
            if feature['failed'] == 0:
                report.write('\n')
    report.close()
    return


def writeToExcelFile(suite, excelFileName):
    excelFileName += '.xlsx'
    wb             = Workbook()
    ws             = wb.active
    header         = ('Tests','Success Rate',"# Tests","# Failed", 'Duration',
                      'Failed Scenarios','Failed Steps')
    ws.append(header)
    subHeader      = (suite['name'],
            str(round(100.0*(suite['total'] - suite['failed']) /
                      suite['total'],1)) + '%', str(suite['total']),
            str(suite['failed']))
    ws.append(subHeader)

    for node in suite['nodes']:
        if node['total'] == 0:
            percent = 'No Tests'
        else:
            percent = str(round(100.0*(node['total'] - node['failed']) /
                                node['total'],1)) + '%'

        row = ("Suite:  " + node['name'][5:],
                percent ,
                str(node['total']),
                str(node['failed']),
                duration(node['duration']))

        ws.append(row)

        for feature in node['features']:
            if feature['total'] == 0:
                percent = 'No Tests'
            else:
                percent = (str(round(100.0*(feature['total'] -
                                            feature['failed']) /
                                     feature['total'],1)) + '%')
                row = (getFeatureFileName(feature['name']), percent,
                        str(feature['total']), str(feature['failed']),
                        duration(feature['duration']), '')
                ws.append(row)

            count = True
            for failure in feature['failureList']:
                if count:
                    ws.cell(row=ws.max_row,
                            column=FAIL_SC_COLS,
                            value=feature['failures'][failure]['scenario'])
                    ws.cell(row=ws.max_row,
                            column=FAIL_ST_COLS,
                            value=feature['failures'][failure]['step'])
                    count = False
                else:
                    ws.append(
                        {FAIL_SC_COLS: feature['failures'][failure]['scenario'],
                        FAIL_ST_COLS: feature['failures'][failure]['step']})

    #merging
    mergeSheet(ws)

    #styling
    #all cells
    for row in ws.rows:
        for cell in row:
            cell.style = defaultStyle

    #first 3 rows
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=3):
        for cell in row:
            cell.style = blueHighlight

    #test suite rows
    suiteCells = takeAllSuites(ws)
    for suiteCell in suiteCells:
        for col in ws.iter_cols(min_row = suiteCell.row,
                                max_col = ws.max_column,
                                max_row = suiteCell.row):
            for cell in col:
                cell.style = suiteStyle
                resizeRow(ws, cell.row, 2)

    #style the rate column
    rateColumn = 'B2:B' + str(ws.max_row)
    ws.conditional_formatting.add(rateColumn,
            CellIsRule(operator='==', formula=['"100.0%"'], fill=greenFill))
    ws.conditional_formatting.add(rateColumn,
            CellIsRule(operator='!=', formula=['"100.0%"'], fill=redFill))

    for row in ws[rateColumn]:
        for cell in row:
            cell.font = Font(color=colors.WHITE)

    for row in ws['B:E']:
        for cell in row:
            cell.alignment = centerAl

    #style the failed scenarios and failed steps columns
    failedColumns = 'F4:G' + str(ws.max_row)
    for row in ws[failedColumns]:
        for cell in row:
            cell.alignment = failedStyle.alignment

    #set failed scenarios and failed steps column widths
    ws.column_dimensions['F'].width = 60
    ws.column_dimensions['G'].width = 60

    #size all the columns
    for  i in range(1, 6):
        resizeToFitColumn(ws, i)

    #resize duration column
    resizeColumn(ws, 5, DURATION_COL_LENGTH)

    #resize first row
    resizeRow(ws, 1, 2)

    #resize rows with merged cells
    for row in ws.iter_rows(min_row=3, max_col=1, max_row=ws.max_row):
        for cell in row:
            failedScenarioCell = cell.offset(column=FAIL_SC_COLS-1)
            failedStepCell = cell.offset(column=FAIL_ST_COLS-1)
            valueLength = max(map(len,
                                  map(str,(failedScenarioCell.value,
                                           failedStepCell.value))))
            if(failedScenarioCell.value):
                size   = 1 + (valueLength // LINE_WRAP_LENGTH)
                resizeRow(ws, cell.row, size)
    #save file
    wb.save(excelFileName)

if __name__ == '__main__':
    pprint(suite)
    init()
    pprint(suite)
    scrapeInfo()
    writeToExcelFile(suite, reportFileName)
    open_file(reportFileName + '.xlsx')
