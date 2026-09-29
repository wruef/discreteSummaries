#!/usr/bin/env python3
"""Build the RCA discrete sample summary, README notes, and lab sample logs for a cruise.

Everything that used to be hardcoded per year is either derived from the cruise data
directory name or exposed as a command line argument.  A cruise directory is expected
to be named  Cabled-NN_<CRUISE>_<YYYY-MM-DD>  and to contain:

    CTD Data/                                   .btl and .hdr files
    ROV Data/                                   .ct2 (JASON) or *_ctd_dive_export.csv (ROPOS)
    <prefix>_DiscreteCastLogs.xlsm              shipboard cast/sample logs
    <prefix>_Oxygen_Sample_Data/                Winkler workbooks, one per analysis day
    <prefix>_<Type>_Sample_Data*.xls[x|m]       lab returns, discovered as they arrive

Examples:
    ./buildDiscreteSummary.py ../CruiseData/Cabled-17_RR2607_2026-08-09
    ./buildDiscreteSummary.py ../CruiseData/Cabled-17_RR2607_2026-08-09 --sample-logs --worktag GR041160
    ./buildDiscreteSummary.py ../CruiseData/Cabled-17_RR2607_2026-08-09 --rov-columns temp=10,cond=9,press=11,sal=13
"""

import argparse
import csv
import datetime as dt
import glob
import operator
import os
import re
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd

fillValue = '-9999999'
scriptDir = os.path.dirname(os.path.abspath(__file__))

### ROV CT2 / CTD export column indices.  These change with whatever is plumbed into the
### vehicle CTD, so verify them each year and override with --rov-columns if needed.
ROVcolumnDefaults = {'JASON': {'temp': 4, 'cond': 5, 'press': 6, 'sal': 7},
                     'ROPOS': {'temp': 10, 'cond': 9, 'press': 11, 'sal': 13, 'oxy': 12}}

### Lab return workbooks, discovered as <prefix>_<glob key>_Sample_Data*.xls*
sampleDataGlobs = {'nut': 'Nutrients', 'sal': 'Salinity', 'fluor': 'Chlorophyll', 'dic': 'DIC'}


###############################################################################
### cruise directory inspection
###############################################################################

def parseArgs():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('cruiseDir', help='cruise data directory, e.g. ../CruiseData/Cabled-17_RR2607_2026-08-09')
    p.add_argument('--outdir', help='where to write outputs (default: the cruise data directory)')
    p.add_argument('--rov', choices=['JASON', 'ROPOS'], help='override ROV auto-detection')
    p.add_argument('--rov-columns', help='CT2/CTD column indices, e.g. temp=10,cond=9,press=11,sal=13')
    p.add_argument('--mean-window', type=float, default=2,
                   help='minutes of ROV CTD data averaged before each bottle closure (default: 2)')
    p.add_argument('--sample-logs', action='store_true', help='also generate lab sample logs')
    p.add_argument('--worktag', default='TBD', help='UW worktag printed on the sample logs')
    p.add_argument('--version', default='ver-1-00', help='version string for sample log file names')
    p.add_argument('--cast-logs', help='override the DiscreteCastLogs workbook')
    p.add_argument('--oxy-dir', help='override the Winkler workbook directory')
    p.add_argument('--nut-file', nargs='+', help='override nutrient workbook(s)')
    p.add_argument('--sal-file', nargs='+', help='override salinity workbook(s)')
    p.add_argument('--fluor-file', nargs='+', help='override chlorophyll workbook(s)')
    p.add_argument('--dic-file', nargs='+', help='override DIC workbook(s)')
    p.add_argument('--customer', default='Wendi Ruef')
    p.add_argument('--email', default='wruef@uw.edu')
    p.add_argument('--office', default='206-221-6760')
    p.add_argument('--pi', default='D. Kelley')
    p.add_argument('--budget-contact', default='Jenny E')
    p.add_argument('--budget-email', default='jenny9@uw.edu')
    p.add_argument('--budget-office', default='Office: 206-542-5279')
    p.add_argument('--maps-dir', default=scriptDir, help='directory holding bottleMap.csv, flags.csv, etc.')
    return p.parse_args()


def cruisePrefix(cruiseDir):
    """Cabled-17_RR2607_2026-08-09 -> ('Cabled-17_RR2607', '2026')"""
    dirName = os.path.basename(os.path.normpath(cruiseDir))
    m = re.match(r'(.+)_(\d{4})-\d{1,2}-\d{1,2}$', dirName)
    if not m:
        raise SystemExit('cruise directory must be named <prefix>_<YYYY-MM-DD>, got: ' + dirName)
    return m.group(1), m.group(2)


def findOne(pattern, label):
    hits = [f for f in glob.glob(pattern) if not os.path.basename(f).startswith('~$')]
    if len(hits) > 1:
        raise SystemExit('multiple %s files match %s: %s' % (label, pattern, hits))
    return hits[0] if hits else None


def findMany(pattern):
    """Labs sometimes send the same workbook twice in different formats; keep one copy per name."""
    ranked = {'.xlsx': 0, '.xlsm': 1, '.xls': 2}
    hits = {}
    for f in sorted(glob.glob(pattern)):
        if os.path.basename(f).startswith('~$'):
            continue
        stem, ext = os.path.splitext(f)
        if stem not in hits or ranked.get(ext, 9) < ranked.get(os.path.splitext(hits[stem])[1], 9):
            hits[stem] = f
    return sorted(hits.values())


def rovFiles(cruiseDir, ROV):
    """JASON logs one .ct2 per dive; ROPOS exports one *_ctd_dive_export.csv.  Extensions vary in case."""
    match = (lambda f: f.endswith('.ct2')) if ROV == 'JASON' else (lambda f: f.endswith('ctd_dive_export.csv'))
    found = []
    for rootdir, dirs, files in os.walk(os.path.join(cruiseDir, 'ROV Data')):
        found += [os.path.join(rootdir, f) for f in files if match(f.lower())]
    return sorted(found)


def detectROV(cruiseDir):
    for ROV in ['JASON', 'ROPOS']:
        if rovFiles(cruiseDir, ROV):
            return ROV
    raise SystemExit('could not detect ROV type from files in ' + os.path.join(cruiseDir, 'ROV Data'))


def parseColumnOverride(spec, defaults):
    columns = dict(defaults)
    for pair in spec.split(','):
        key, value = pair.split('=')
        columns[key.strip()] = int(value)
    return columns


###############################################################################
### parsers
###############################################################################

def parseNMEA(line, castDict, cruiseYear=None):
    """Pull latitude, longitude, and UTC start time out of a Seabird header line."""
    if '* NMEA Latitude' in line:
        # * NMEA Latitude = 45 49.81 N
        m = re.search(r'.* NMEA Latitude\s=\s(\d*)\s(\d*.\d*)\s.*', line)
        if m:
            castDict['Start Latitude [degrees]'] = float(m.group(2)) / 60 + float(m.group(1))
    if '* NMEA Longitude' in line:
        # * NMEA Longitude = 129 44.77 W
        m = re.search(r'.* NMEA Longitude\s=\s(\d*)\s(\d*.\d*)\s.*', line)
        if m:
            castDict['Start Longitude [degrees]'] = -(float(m.group(2)) / 60 + float(m.group(1)))
    if '* NMEA UTC' in line:
        # * NMEA UTC (Time) = Jul 30 2017 11:01:22  ->  2017-07-30T11:01:22.000Z
        m = re.search(r'.*NMEA\sUTC.*=.*([a-zA-Z]{3}).*(\d{2}).*(\d{4}).*(\d{2}:\d{2}:\d{2}).*', line)
        if m:
            monthInt = '%02d' % dt.datetime.strptime(m.group(1), '%b').month
            castDict['Start Time [UTC]'] = '%s-%s-%sT%s.000Z' % (m.group(3), monthInt, m.group(2), m.group(4))


def parseBottleFile(btlFile, bottleMap_dict):
    castDict = {}
    bottleTimes = []
    bottleData = []
    bottleHeader = []

    with open(btlFile, 'r') as f:
        btlLines = f.readlines()

    for line in btlLines:
        if line.startswith('*') or line.startswith('#') or re.search(r'.*Bottle.*Date.*', line) or re.search(r'.*Position.*Time.*', line):
            parseNMEA(line, castDict)
            if re.search(r'.*Bottle.*Date.*', line):
                bottleHeader = line.split()
        else:
            dataLines = line.split()
            if re.search(r'.*\d{2}:\d{2}:\d{2}.*', dataLines[0]):
                bottleTimes.append(dataLines[0])
            elif re.search(r'.*[1-9]|1[1-9]|2[1-4].*', dataLines[0]):
                bottleData.append(dataLines)

    bottleHeader = bottleHeader[2:]
    castDict.setdefault('BottleData', {})
    for i in range(len(bottleData)):
        bottle = int(bottleData[i][0])
        castDict['BottleData'].setdefault(bottle, {})
        monthInt = '%02d' % dt.datetime.strptime(bottleData[i][1], '%b').month
        timeString = '%s-%s-%sT%s.000Z' % (bottleData[i][3], monthInt, bottleData[i][2], bottleTimes[i])
        castDict['BottleData'][bottle]['CTD Bottle Closure Time [UTC]'] = timeString
        dataList = bottleData[i][4:-1]
        for j in range(len(bottleHeader)):
            if bottleHeader[j] in bottleMap_dict:
                castDict['BottleData'][bottle][bottleMap_dict[bottleHeader[j]]] = dataList[j]
        if 'Ph' not in bottleHeader:
            castDict['BottleData'][bottle]['CTD pH'] = fillValue

    return castDict


def parseHeaderFile(hdrFile):
    hdrDict = {}
    with open(hdrFile, 'r') as f:
        for line in f:
            parseNMEA(line, hdrDict)
    return hdrDict


def parseROVfile(ROVfile, columns, ROV, cruiseYear):
    ROVdata = []
    with open(ROVfile, 'r') as f:
        ROVlines = f.readlines()

    for line in ROVlines:
        if ROV == 'ROPOS':
            if not line.startswith(cruiseYear):
                continue
            dataLine = line.split(',')
            date = dataLine[0].split(' ')
            timeString = date[0] + 'T' + date[1] + '.000Z'
        else:
            if not line.startswith('CT2'):
                continue
            dataLine = line.split(',')
            timeString = dataLine[1].replace('/', '-') + 'T' + dataLine[2] + 'Z'
        ROVdata.append([timeString,
                        float(dataLine[columns['press']].strip(',')),
                        float(dataLine[columns['temp']].strip(',')),
                        float(dataLine[columns['cond']].strip(',')),
                        float(dataLine[columns['sal']].strip(','))])

    return pd.DataFrame(ROVdata, columns=['ROV Time', 'CTD Pressure [db]', 'CTD Temperature 1 [deg C]',
                                          'CTD Conductivity 1 [S/m]', 'CTD Salinity 1 [psu]'])


def parseOxygenFile(oxyFile):
    ### Winkler workbooks: cast # in column B, niskin # in column G, bottle # in column AB, [O2] ml/l in column AK
    ### data rows are the ones with a numeric [O2]
    df_oxy = pd.read_excel(oxyFile, sheet_name='Winkler Calculations', header=None, usecols=[1, 6, 27, 36],
                           names=['Cast #', 'Niskin #', 'Sample Bottle #', 'Discrete Oxygen [mL/L]'])
    df_oxy['Discrete Oxygen [mL/L]'] = pd.to_numeric(df_oxy['Discrete Oxygen [mL/L]'], errors='coerce')
    ### the cast # is entered once per cast; replicate rows beneath it are left blank
    df_oxy['Cast #'] = df_oxy['Cast #'].ffill()
    df_oxy = df_oxy.dropna()

    ### the lab labels niskin positions fwd/aft; convert to the sample log convention
    df_oxy['Niskin #'] = df_oxy['Niskin #'].replace({'fwd': 'Forward', 'aft': 'Aft'})

    return df_oxy


###############################################################################
### helpers
###############################################################################

def flagBits(flagStrings, column, flagHeadersMap_dict, flagBitMap_dict):
    bits = []
    flagBitMap = list('*0000000000000000')
    flagColumns = flagBitMap_dict[flagHeadersMap_dict[column]]
    for flagStr in flagStrings.split(','):
        for k, v in flagColumns.items():
            if flagStr.strip() in v:
                bits.append(k)
    for bit in bits:
        flagBitMap[16 - int(bit)] = '1'
    return ''.join(flagBitMap)


def meanROVdata(sampleTime, meanWindow_mins, df_ROV):
    try:
        bottleTime = dt.datetime.strptime(sampleTime, '%Y-%m-%dT%H:%M:%S.%fZ')
    except ValueError:
        bottleTime = dt.datetime.strptime(sampleTime, '%Y-%m-%dT%H:%M:%SZ')
    windowStart = bottleTime - dt.timedelta(minutes=meanWindow_mins)
    if not pd.api.types.is_datetime64_any_dtype(df_ROV['ROV Time']):
        ### timestamps are already UTC, so drop the tz to keep the window comparison naive;
        ### the vehicle occasionally logs an invalid second (:60), which coerces to NaT and falls out of the window
        df_ROV['ROV Time'] = pd.to_datetime(df_ROV['ROV Time'], format='ISO8601', utc=True,
                                            errors='coerce').dt.tz_localize(None)
    df_window = df_ROV.loc[(df_ROV['ROV Time'] > windowStart) & (df_ROV['ROV Time'] <= bottleTime)]
    meanData = df_window.mean(numeric_only=True)
    ### the cast log no longer carries a dive Start Time, so report the first ROV record in the averaging window
    meanData['Start Time [UTC]'] = df_window['ROV Time'].min().isoformat(timespec='milliseconds') + 'Z'
    return meanData


def bottleDigits(label):
    """Bottle labels drift year to year (OXY-010, RCA010, SAL-0313); match on the number alone."""
    return re.sub(r'\D', '', str(label)).lstrip('0')


def limitDecimals(value, places=6):
    """Round over-precise values, keeping plain decimal notation and leaving flags/fillValue alone."""
    valueString = str(value)
    if valueString == fillValue or not re.fullmatch(r'-?\d+\.?\d*([eE][-+]?\d+)?', valueString):
        return value
    decimalValue = Decimal(valueString)
    if abs(decimalValue.as_tuple().exponent) > places:
        decimalValue = decimalValue.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    return format(decimalValue, 'f')


###############################################################################
### loaders
###############################################################################

def loadCTD(cruiseDir, bottleMap_dict):
    bottleDict = {}
    headerDict = {}
    for rootdir, dirs, files in os.walk(os.path.join(cruiseDir, 'CTD Data')):
        for file in files:
            m = re.search(r'.*(CTD-\d*)\.(btl|hdr)$', file)
            if not m:
                continue
            if m.group(2) == 'btl':
                bottleDict[m.group(1)] = parseBottleFile(os.path.join(rootdir, file), bottleMap_dict)
            else:
                headerDict[m.group(1)] = parseHeaderFile(os.path.join(rootdir, file))
    return bottleDict, headerDict


def loadROV(cruiseDir, ROV, ROVcolumns, cruiseYear):
    ROVdict = {}
    diveRegex = r'(J2-\d+)' if ROV == 'JASON' else r'(R\d+)_ctd_dive_export'
    for ROVfile in rovFiles(cruiseDir, ROV):
        m = re.search(diveRegex, os.path.basename(ROVfile), re.IGNORECASE)
        if m:
            ROVdict[m.group(1)] = parseROVfile(ROVfile, ROVcolumns, ROV, cruiseYear)
        else:
            print('error retrieving dive number from file: ', ROVfile)
    return ROVdict


def loadCastLogs(castLogFile):
    """Read the shipboard DiscreteCastLogs workbook."""
    sheets = pd.read_excel(castLogFile, sheet_name=None)

    df_casts = sheets['CastList']
    df_samples = sheets['SampleList']
    df_samples = df_samples[~df_samples['Cast'].isnull()]
    df_samples = df_samples.fillna(fillValue)

    ### OxygenLog_all is the shipboard titration record; summary oxygen comes from the lab Winkler workbooks
    df_chlorophyll = sheets['ChloroLog_all']
    df_chlorophyll = df_chlorophyll[~df_chlorophyll['Unnamed: 3'].isnull()].reset_index()
    new_header = df_chlorophyll.iloc[0]              # grab the first row for the header
    df_chlorophyll = df_chlorophyll.fillna(fillValue)
    df_chlorophyll = df_chlorophyll[1:].astype(str)  # take the data less the header row
    df_chlorophyll.columns = new_header

    return {'df_casts': df_casts,
            'df_samples': df_samples,
            'df_chlorophyll': df_chlorophyll.filter(['Cast #', 'Niskin #', 'Sample Bottle #',
                                                     'Chlorophyll Vial', 'Volume Filtered', 'Acetone Volume']),
            'df_CastLog_ROV': sheets['CastLog_ROV'],
            'df_CTDflags': sheets['CTDflags'],
            'df_CTDflags_ROV': sheets['CTDflags_ROV']}


def buildMetadata(df_casts):
    metadataDict = {}
    castColumns = [c for c in df_casts.columns if c != 'Cast']
    df_casts = df_casts.fillna(fillValue)
    for index, row in df_casts.iterrows():
        if not isinstance(row['Cruise'], float):
            metadataDict[row['Cast']] = {col: row[col] for col in castColumns}
    return df_casts, metadataDict


def loadNutrients(files):
    """All sheets of all nutrient workbooks are merged; the labs return them in batches."""
    df_list = []
    nut_header = ['index', 'bottle #', 'Discrete Phosphate [uM]', 'Discrete Silicate [uM]',
                  'Discrete Nitrate [uM]', 'Discrete Nitrite [uM]', 'Discrete Ammonium [uM]']
    for file in files:
        for k, v in pd.read_excel(file, sheet_name=None).items():
            ### data rows are the ones with a bottle number; skips the header, blank, and check-standard rows
            filteredData = v[pd.to_numeric(v['Unnamed: 1'], errors='coerce').notnull()][
                ['Unnamed: 1', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']].reset_index()
            filteredData.columns = nut_header
            df_list.append(filteredData)
    return pd.concat(df_list, axis=0, ignore_index=True).fillna(fillValue)


def loadOxygen(oxyDir):
    df_list = []
    for rootdir, dirs, files in os.walk(oxyDir):
        for file in files:
            if file.endswith('.xls') and not file.startswith('~$'):
                df_list.append(parseOxygenFile(os.path.join(rootdir, file)))
    return pd.concat(df_list, axis=0, ignore_index=True).astype(str)


def loadSalinity(files):
    df_list = []
    sal_header = ['index', 'bottle #', 'Discrete Salinity [psu]']
    for file in files:
        for k, v in pd.read_excel(file, sheet_name=None).items():
            ### data rows are the ones with both a bottle number and a salinity; skips header, standard, check-standard rows
            filteredData = pd.DataFrame({'bottle #': pd.to_numeric(v['Unnamed: 3'], errors='coerce'),
                                         'Discrete Salinity [psu]': pd.to_numeric(v['Unnamed: 5'], errors='coerce')}).dropna().reset_index()
            filteredData.columns = sal_header
            ### bottle numbers must stay integers, the sample log looks them up as '313', not '313.0'
            filteredData['bottle #'] = filteredData['bottle #'].astype(int)
            df_list.append(filteredData)
    return pd.concat(df_list, axis=0).fillna(fillValue)


def loadChlorophyll(files):
    df_list = []
    fluor_header = ['index', 'bottle #', 'Discrete Chlorophyll [ug/L]',
                    'Discrete Phaeopigment [ug/L]', 'Discrete Fo/Fa Ratio']
    for file in files:
        for k, v in pd.read_excel(file, sheet_name=None).items():
            ### vial number in column A, Chl a / phaeopigment / Fo-Fa ratio in columns H, I, J
            ### data rows are the ones with a numeric vial number; header names change year to year
            vials = pd.to_numeric(v.iloc[:, 0], errors='coerce')
            filteredData = v[vials.notnull()].iloc[:, [0, 7, 8, 9]].reset_index()
            filteredData.columns = fluor_header
            df_list.append(filteredData)
    return pd.concat(df_list, axis=0).fillna(fillValue)


def loadDIC(files):
    df_list = []
    dic_header = ['index', 'bottle #', 'pCO2 Analysis Temp [deg C]', 'Calculated Alkalinity [umol/kg]',
                  'Discrete DIC [umol/kg]', 'Discrete pCO2 [uatm]', 'Calculated CO2aq [umol/kg]',
                  'Calculated Bicarb [umol/kg]', 'Calculated CO3 [umol/kg]', 'Calculated pH',
                  'Calculated Omega-C', 'Calculated Omega-A']
    for file in files:
        for k, v in pd.read_excel(file, sheet_name=None).items():
            bottleColumn = v.columns[0]
            filteredData = v[(v[bottleColumn].str.contains('DIC')) & (v['AnalysisT'].apply(isinstance, args=(float,)))][
                [bottleColumn, 'AnalysisT', 'alk (µeq/kg)', 'TCO2 (µmol/kg)', 'pco2_in situ (µatm)',
                 'co2aq (µmol/kg)', 'bicarb (µmol/kg)', 'co3 (µmol/kg)', 'pHt', 'omega-C', 'omega-A']].reset_index()
            filteredData.columns = dic_header
            df_list.append(filteredData)
    return pd.concat(df_list, axis=0).fillna(fillValue)


###############################################################################
### sample logs
###############################################################################

sampleLogSpecs = {
    'nut':    {'file': 'Nutrients',   'title': 'Nutrient Analyses',              'units': 'uM',
               'bottle': 'Nutrient Bottle Number',    'extra': []},
    'sal':    {'file': 'Salinity',    'title': 'Salinity Analyses',              'units': 'salinity',
               'bottle': 'Salinity Bottle Number',    'extra': []},
    'chloro': {'file': 'Chlorophyll', 'title': 'Chlorophyll Analyses',           'units': 'ug/L',
               'bottle': 'Chlorophyll Bottle Number', 'extra': ['Volume Filtered (L)', 'Extraction Volume (L)']},
    'dic':    {'file': 'DIC',         'title': 'DIC Analysis for pCO2 and TCO2', 'units': None,
               'bottle': 'DIC Bottle Number',         'extra': ['In-Situ Temperature (avg)', 'In-Situ Salinity (avg)']},
}

standardColumns = ['CTD Station', 'CTD ID', 'Niskin Bottle Number', 'Target Depth (m)', 'Sample Bottle Number']


def sampleLogRows(sampleType, spec, compileDict, divePrefix, meanWindow):
    """One row per drawn sample: station, CTD id, niskin, target depth, bottle, plus per-type extras."""
    df_samples = compileDict['df_samples']
    df_chloro = compileDict['df_chlorophyll']
    bottleString = spec['bottle']
    df_sub = df_samples[~df_samples[bottleString].str.match(fillValue, na=False)]

    rows = []
    for index, row in df_sub.iterrows():
        sampleRow = [str(compileDict['metadataDict'][row['Cast']]['Station']),
                     str(compileDict['metadataDict'][row['Cast']]['CTD File']),
                     str(row['Niskin/Bottle Position']),
                     str(row['Target Depth'])]
        if sampleType == 'chloro':
            match = ((df_chloro['Sample Bottle #'].str.match(row[bottleString])) &
                     (df_chloro['Cast #'].str.match(row['Cast'])))
            for column in ['Chlorophyll Vial', 'Volume Filtered', 'Acetone Volume']:
                sampleRow.append(str(df_chloro.loc[match, column].values[0]))
        elif sampleType == 'dic':
            sampleRow.append(str(row[bottleString]))
            if divePrefix not in row.Cast:
                bottleData = compileDict['bottleDict'][row.Cast]['BottleData'][row['Niskin/Bottle Position']]
                sampleRow.append(str(bottleData['CTD Temperature 1 [deg C]']))
                sampleRow.append(str(bottleData['CTD Salinity 1 [psu]']))
            else:
                df_rov = compileDict['df_CastLog_ROV']
                sampleTime = df_rov.loc[(df_rov['Dive'].str.match(row['Cast'])) &
                                        (df_rov['Niskin'].str.match(row['Niskin/Bottle Position'])),
                                        'CTD Bottle Closure Time [UTC]'].values[0]
                ROVctdData = meanROVdata(sampleTime, meanWindow, compileDict['ROVdict'][row['Cast']])
                sampleRow.append(str(ROVctdData['CTD Temperature 1 [deg C]']))
                sampleRow.append(str(ROVctdData['CTD Salinity 1 [psu]']))
        else:
            sampleRow.append(str(row[bottleString]))
        rows.append(sampleRow)
    return rows


def writeSampleLogs(args, compileDict, prefix, outdir, divePrefix, today):
    """Sample logs for nutrients, salinity, chlorophyll, and DIC, for distribution to the analysis labs.

    All four share a header block; chlorophyll and DIC each add two columns.
    """
    written = []
    for sampleType, spec in sampleLogSpecs.items():
        rows = sampleLogRows(sampleType, spec, compileDict, divePrefix, args.mean_window)
        logFile = os.path.join(outdir, '%s_%s_Sample_Log_%s_%s.csv' % (prefix, spec['file'], today, args.version))
        with open(logFile, 'w') as f:
            f.write(spec['title'] + '\n')
            f.write('\n' if spec['units'] is None else ',,Units Required,%s\n' % spec['units'])
            f.write('\n')
            f.write('Customer and Data Recipient,%s\n' % args.customer)
            f.write('Email,%s\n' % args.email)
            f.write('Office,%s\n' % args.office)
            f.write('\n')
            f.write('Worktage,%s,Budget Contact,%s\n' % (args.worktag, args.budget_contact))
            f.write('PI,%s,,%s\n' % (args.pi, args.budget_email))
            f.write(',,,%s\n' % args.budget_office)
            f.write('Total Samples,%d\n' % len(rows))
            f.write('\n')
            f.write('\n')
            f.write(','.join(standardColumns + spec['extra']) + '\n')
            for sampleRow in rows:
                f.write(','.join(sampleRow) + '\n')
        written.append(logFile)
    return written


###############################################################################
### summary rows
###############################################################################

def dataOnlyRows(compileDict, headers, headerMap_dict, flagMaps):
    """Casts with no Niskins triggered still get a row, carrying metadata and header position only."""
    dataRows = []
    for index, row in compileDict['df_casts'].iterrows():
        if 'Data cast only, no Niskins triggered' not in row['CTD File Flag']:
            continue
        dataRow = []
        for column in headers:
            source = headerMap_dict[column]
            if ',' in source:
                source = source.split(',')[0 if 'CTD' in row.Cast else 1]
            if 'metadataDict' in source:
                dataCell = compileDict[source][row.Cast][column]
                if 'Flag' in column:
                    dataCell = fillValue if fillValue in dataCell else flagBits(dataCell, column, *flagMaps)
            elif 'Cast' in column:
                dataCell = row.Cast
            elif column in ['Start Latitude [degrees]', 'Start Longitude [degrees]', 'Start Time [UTC]']:
                dataCell = compileDict['headerDict'][row.Cast][column]
            else:
                dataCell = fillValue
            dataRow.append(dataCell)
        dataRows.append(dataRow)
    return dataRows


def discreteSampleRows(compileDict, headers, headerMap_dict, flagMaps, meanWindow):
    """One summary row per discrete sample, pulling each column from the source named in the header map."""
    dataRows = []
    ### the DIC lab does not report these; they stay filled
    NA_DIC_vars = ['Discrete Alkalinity [umol/kg]', 'Discrete pH [Total scale]', 'pH Analysis Temp [deg C]',
                   'Calculated DIC [umol/kg]', 'Calculated pCO2 [uatm]']

    for index, row in compileDict['df_samples'].iterrows():
        dataRow = []
        for column in headers:
            source = headerMap_dict[column]
            if ',' in source:
                source = source.split(',')[0 if 'CTD' in row.Cast else 1]

            if 'df_samples' in source:
                dataCell = row[column]
                if 'Flag' in column:
                    dataCell = fillValue if fillValue in dataCell else flagBits(dataCell, column, *flagMaps)

            elif any(s in source for s in ['metadataDict', 'bottleDict']):
                if 'bottleDict_bottle' in source:
                    dataCell = compileDict['bottleDict'][row.Cast]['BottleData'][row['Niskin/Bottle Position']][column]
                else:
                    dataCell = compileDict[source][row.Cast][column]
                    if 'Flag' in column:
                        dataCell = fillValue if fillValue in dataCell else flagBits(dataCell, column, *flagMaps)

            elif any(s in source for s in ['df_CTDflags', 'df_CTDflags_ROV']):
                df = compileDict[source]
                cell = df.loc[df['Parameter'] == column, 'Parameter Flag']
                dataCell = flagBits(cell.values[0], column, *flagMaps) if len(cell) > 0 else fillValue

            elif 'df_CastLog_ROV' in source:
                df = compileDict[source]
                match = (df['Dive'] == row.Cast) & (df['Niskin'] == row['Niskin/Bottle Position'])
                if column not in df:
                    ### 'Start Time [UTC]' was dropped from the cast log; take it from the ROV averaging window
                    bottleTime = df.loc[match, 'CTD Bottle Closure Time [UTC]'].values[0]
                    dataCell = meanROVdata(bottleTime, meanWindow, compileDict['ROVdict'][row.Cast])[column]
                else:
                    dataCell = df.loc[match, column].values[0]

            elif 'ROVdict' in source:
                df = compileDict['df_CastLog_ROV']
                bottleTime = df.loc[(df['Dive'] == row.Cast) & (df['Niskin'] == row['Niskin/Bottle Position']),
                                    'CTD Bottle Closure Time [UTC]'].values[0]
                dataCell = meanROVdata(bottleTime, meanWindow, compileDict['ROVdict'][row.Cast])[column]

            elif 'fill' in source:
                dataCell = fillValue

            elif 'df_oxygen' in source:
                df = compileDict[source]
                dataCell = fillValue
                if row['Oxygen Bottle Number'] != fillValue and not isinstance(df, str):
                    cell = df.loc[(df['Sample Bottle #'].map(bottleDigits) == bottleDigits(row['Oxygen Bottle Number'])) &
                                  (df['Cast #'] == str(row.Cast)), column]
                    if len(cell) > 0:
                        dataCell = cell.values[0]
                    else:
                        print('error retrieving oxygen values...', row['Oxygen Bottle Number'], row.Cast,
                              row['Niskin/Bottle Position'])

            elif 'df_fluor' in source:
                df = compileDict[source]
                dataCell = fillValue
                if row['Chlorophyll Bottle Number'] != fillValue and not isinstance(df, str):
                    df_chloro = compileDict['df_chlorophyll']
                    vialNumber = df_chloro.loc[(df_chloro['Sample Bottle #'].str.match(row['Chlorophyll Bottle Number'])) &
                                               (df_chloro['Cast #'] == str(row.Cast)) &
                                               (df_chloro['Niskin #'] == str(row['Niskin/Bottle Position'])),
                                               'Chlorophyll Vial'].values[0]
                    df['bottle #'] = df['bottle #'].astype(str)
                    cell = df.loc[df['bottle #'] == vialNumber, column]
                    if len(cell) > 0:
                        dataCell = cell.values[0]
                    else:
                        print('error retrieving chlorophyll values...', vialNumber)

            elif 'df_nuts' in source:
                df = compileDict[source]
                dataCell = fillValue
                if row['Nutrient Bottle Number'] != fillValue and not isinstance(df, str):
                    cell = df.loc[df['bottle #'].map(bottleDigits) == bottleDigits(row['Nutrient Bottle Number']), column]
                    if len(cell) > 0:
                        dataCell = cell.values[0]
                    else:
                        print('error retrieving nutrient values...', row['Nutrient Bottle Number'])

            elif 'df_sal' in source:
                df = compileDict[source]
                dataCell = fillValue
                if row['Salinity Bottle Number'] != fillValue and not isinstance(df, str):
                    cell = df.loc[df['bottle #'].map(bottleDigits) == bottleDigits(row['Salinity Bottle Number']), column]
                    if len(cell) > 0:
                        dataCell = cell.values[0]
                    else:
                        print('error retrieving salinity values...', row['Salinity Bottle Number'])

            elif 'df_DIC' in source:
                df = compileDict[source]
                dataCell = fillValue
                if column not in NA_DIC_vars and row['DIC Bottle Number'] != fillValue and not isinstance(df, str):
                    cell = df.loc[df['bottle #'].str.contains(row['DIC Bottle Number']), column]
                    if len(cell) > 0:
                        dataCell = cell.values[0]
                    else:
                        print('error retrieving carbon values...', row['DIC Bottle Number'])

            dataRow.append(dataCell)
        dataRows.append(dataRow)
    return dataRows


def flagDiscreteMatches(dataRows, headers):
    """Set the 'discrete sample available' bit on the CTD parameter flag wherever a discrete value exists."""
    discreteCTDmatch = {'Discrete Salinity [psu]': ['CTD Conductivity 1 Flag', 'CTD Conductivity 2 Flag'],
                        'Discrete Oxygen [mL/L]': ['CTD Oxygen Flag'],
                        'Discrete Chlorophyll [ug/L]': ['CTD Fluorescence Flag'],
                        'Discrete pH [Total scale]': ['CTD pH Flag'],
                        'Calculated pH': ['CTD pH Flag']}

    for row in dataRows:
        for key, flags in discreteCTDmatch.items():
            discreteIndex = headers.index(key)
            if fillValue in str(row[discreteIndex]):
                continue
            for flag in flags:
                flagIndex = headers.index(flag)
                if fillValue not in row[flagIndex]:
                    newFlag = list(row[flagIndex])
                    newFlag[9] = '1'
                    row[flagIndex] = ''.join(newFlag)


def writeREADME(readmeFile, compileDict):
    """CTD file list for the README re-naming table, plus every note from the CastList and SampleList."""
    ctdFiles = [v['CTD File'] for k, v in compileDict['metadataDict'].items() if fillValue not in k]

    notes = []
    for index, row in compileDict['df_samples'].iterrows():
        if fillValue not in row.Notes:
            notes.append('%s, %s, Niskin %s: %s' % (compileDict['metadataDict'][row.Cast]['Cruise'], row.Cast,
                                                    row['Niskin/Bottle Position'], row.Notes))
    for key, value in compileDict['metadataDict'].items():
        if fillValue not in value['Notes']:
            notes.append('%s, %s, %s' % (value['Cruise'], key, value['Notes']))

    with open(readmeFile, 'w') as f:
        f.write('File Mapping:\n')
        f.writelines(line + '\n' for line in ctdFiles)
        f.write('Summary Notes:\n')
        f.writelines(line + '\n' for line in notes)


###############################################################################
### main
###############################################################################

def main():
    args = parseArgs()
    cruiseDir = os.path.normpath(args.cruiseDir)
    prefix, cruiseYear = cruisePrefix(cruiseDir)
    outdir = args.outdir or cruiseDir
    today = dt.date.today().isoformat()

    ROV = args.rov or detectROV(cruiseDir)
    ROVcolumns = ROVcolumnDefaults[ROV]
    if args.rov_columns:
        ROVcolumns = parseColumnOverride(args.rov_columns, ROVcolumns)
    divePrefix = 'J2' if ROV == 'JASON' else 'R'

    castLogFile = args.cast_logs or findOne(os.path.join(cruiseDir, prefix + '_DiscreteCastLogs.xls*'), 'cast log')
    if not castLogFile:
        raise SystemExit('no DiscreteCastLogs workbook found in ' + cruiseDir)
    oxyDir = args.oxy_dir or findOne(os.path.join(cruiseDir, '*Oxygen_Sample_Data'), 'oxygen data')
    labFiles = {key: getattr(args, key + '_file') or
                     findMany(os.path.join(cruiseDir, '%s_%s_Sample_Data*.xls*' % (prefix, name)))
                for key, name in sampleDataGlobs.items()}

    print('cruise:      %s (%s, %s)' % (prefix, ROV, cruiseYear))
    print('cast logs:   %s' % os.path.basename(castLogFile))
    print('oxygen:      %s' % (os.path.basename(oxyDir) if oxyDir else 'not yet received'))
    for key, files in labFiles.items():
        print('%-12s %s' % (key + ':', ', '.join(os.path.basename(f) for f in files) or 'not yet received'))

    ### maps: summary headers to data sources, bottle file headers to summary columns, flag strings to bit positions
    mapPath = lambda name: os.path.join(args.maps_dir, name)
    headerMap_dict = pd.read_csv(mapPath('discreteSummaryHeaderMap.csv'), index_col=0).squeeze('columns').to_dict()
    bottleMap_dict = pd.read_csv(mapPath('bottleMap.csv'), index_col=0).squeeze('columns').to_dict()
    flagHeadersMap_dict = pd.read_csv(mapPath('flagMap.csv'), index_col=0).squeeze('columns').to_dict()
    flagBitMap_dict = pd.read_csv(mapPath('flags.csv'), index_col=0).squeeze('columns').to_dict()
    flagMaps = (flagHeadersMap_dict, flagBitMap_dict)

    compileDict = loadCastLogs(castLogFile)
    compileDict['df_casts'], compileDict['metadataDict'] = buildMetadata(compileDict['df_casts'])
    compileDict['ROVdict'] = loadROV(cruiseDir, ROV, ROVcolumns, cruiseYear)
    compileDict['bottleDict'], compileDict['headerDict'] = loadCTD(cruiseDir, bottleMap_dict)

    compileDict['df_nuts'] = loadNutrients(labFiles['nut']) if labFiles['nut'] else fillValue
    compileDict['df_sal'] = loadSalinity(labFiles['sal']) if labFiles['sal'] else fillValue
    compileDict['df_fluor'] = loadChlorophyll(labFiles['fluor']) if labFiles['fluor'] else fillValue
    compileDict['df_DIC'] = loadDIC(labFiles['dic']) if labFiles['dic'] else fillValue
    compileDict['df_oxygen'] = loadOxygen(oxyDir) if oxyDir else fillValue

    if args.sample_logs:
        for logFile in writeSampleLogs(args, compileDict, prefix, outdir, divePrefix, today):
            print('wrote ' + logFile)

    headers = list(headerMap_dict)
    dataRows = dataOnlyRows(compileDict, headers, headerMap_dict, flagMaps)
    dataRows += discreteSampleRows(compileDict, headers, headerMap_dict, flagMaps, args.mean_window)

    castIndex, niskinIndex = headers.index('Cast'), headers.index('Niskin/Bottle Position')
    dataRows = sorted(dataRows, key=operator.itemgetter(castIndex, niskinIndex))

    flagDiscreteMatches(dataRows, headers)

    ### lat/long are pinned to 5 decimal places, everything else is trimmed to 6
    latIndex, lonIndex = headers.index('Start Latitude [degrees]'), headers.index('Start Longitude [degrees]')
    for row in dataRows:
        row[latIndex] = '{:.5f}'.format(row[latIndex])
        row[lonIndex] = '{:.5f}'.format(row[lonIndex])

    summaryFile = os.path.join(outdir, prefix + '_Discrete_Summary.csv')
    with open(summaryFile, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows([limitDecimals(cell) for cell in row] for row in dataRows)
    print('wrote %s (%d rows)' % (summaryFile, len(dataRows)))

    readmeFile = os.path.join(outdir, prefix + '_README_notes')
    writeREADME(readmeFile, compileDict)
    print('wrote ' + readmeFile)


if __name__ == '__main__':
    main()
