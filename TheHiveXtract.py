#!/usr/bin/env python2
# coding: utf-8
#Author: Florian PERRET @Cyber_Pescadito
#Date: 31.07.2019
#Description: Extract data from TheHive to CSV & formated XLSX

import datetime
import os
import requests
import json
import csv
import shutil as sl
from pyexcel.cookbook import merge_all_to_a_book,merge_two_files
import pyexcel
import glob
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')


thehive_api_url='***YourHiveURL***'
thehive_key='***YourAPIKey***'

def humanToEpoch(day,month,year):
      epoch = int((datetime.datetime(year,month,day,0,0) - datetime.datetime(1970,1,1)).total_seconds() * 1000)
      return epoch

def mkstmp(ts,tfmt='%m/%d/%Y %H:%M CDT'):
        if not type(ts) is int:
                ts=int(ts)
        return datetime.datetime.fromtimestamp(ts/1000).strftime(tfmt)

def GetCases(thehive_api_url,thehive_key):
        h = {'content-type': 'application/json'}
        token = 'Bearer ' + thehive_key
        h['Authorization'] = token
        payload= {'sort': '-startDate'}
        searchParams={"query":{"_and":[{"_or":[{"_field":"severity","_value":1},{"_field":"severity","_value":2},{"_field":"severity","_value":3}]},{"_and":[{"_not":{"status":"Deleted"}},{"_not":{"_in":{"_field":"_type","_values":["dashboard","data","user","analyzer","caseTemplate","reportTemplate","action"]}}}]}]}}
        search_url = thehive_api_url + '/_search?range=0-9999&nparent=10'
        response=requests.post(search_url, params=payload,json=searchParams, headers=h, verify='ca-certificates.crt')
        thecase = json.loads(response.text)

        return thecase

def PutCasesOnFile(fileName,listCase):
        j=listCase
        columns=[]
        for k in j[0]:
                columns.append(k)
        for case in j:
                for f in case['customFields']:
                        columns.append(f)
        if 'createdAt' in columns:
                columns.remove('createdAt')
        if 'pap' in columns:
                columns.remove('pap')
        if 'flag' in columns:
                columns.remove('flag')
        if 'updatedAt' in columns:
                columns.remove('updatedAt')
        if 'metrics' in columns:
                columns.remove('metrics')
        if '_type' in columns:
                columns.remove('_type')
        if '_version' in columns:
                columns.remove('_version')
        if '_routing' in columns:
                columns.remove('_routing')
        if '_id' in columns:
                columns.remove('_id')
        if '_parent' in columns:
                columns.remove('_parent')
        if 'customFields' in columns:
                columns.remove('customFields')
        columns=list(set(columns))
        columns.insert(0,'createdAt')
        csv=','.join(columns).strip().strip(",")+"\n"
        for case in j:
                row=""
                case['createdAt']=mkstmp(case['createdAt'])
                case['startDate']=mkstmp(case['startDate'])
                for c in columns:
                        if c in case:
                                cell=str(case[c])
                                cell=cell.replace("\r\n\r-","|")
                                cell=cell.replace("\r\n","|")
                                cell=cell.replace("\n","|")
                                cell=cell.replace(",",";")
                                row+=cell[:512]+","
                        elif c in case['customFields']:
                                f=case['customFields'][c]
                                if 'boolean' in f:
                                        f=f['boolean']
                                elif 'number' in f:
                                        f=f['number']
                                elif 'string' in f:
                                        f=f['string']
                                str(f).encode('utf-8', 'ignore')
                                cell=str(f)
                                cell=cell.replace("\r\n\r-","|")
                                cell=cell.replace("\r\n","|")
                                cell=cell.replace("\n","|")
                                cell=cell.replace(",",";")
                                row+=cell[:512]+","
                        else:
                                row+=","
                row=row.strip().strip(",")+"\n"
                row=row.decode('latin-1', 'ignore')
                csv+=row
        #response['Content-Type']="text/csv"
        #response.write(csv)
        with open(fileName,"w+") as f:
                f.write(csv+"\n")

def formatAsTable(sitexlsx):
        workbook = xlsxwriter.Workbook(sitexlsx)
        worksheet1.add_table()

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def csvToList():
        with open(sitecsv) as f:
                reader = csv.reader(f)
                data = []
                for row in reader:
                        data.append(row)
                return data

#create fake site list...
siteList = ['export']
idsite = 0

#Create a CSV for each site
for site in siteList:
        listCase = GetCases(thehive_api_url,thehive_key)
        fileName= 'export' + '.csv'
        PutCasesOnFile(fileName,listCase)
        idsite += 1

#Create a xlsx for each site
for site in siteList:
        sitecsv = site + '.csv'
        sitexlsx = site + '.xlsx'
        merge_all_to_a_book(glob.glob(sitecsv),sitexlsx)
        sheet = pyexcel.get_sheet(file_name=sitexlsx)
        dataList = csvToList()
        rawHeaders = dataList[0]
        theHeaders=[]
        for item in rawHeaders:
                theHeaders.append({'header': item})
        del dataList[0]
        colCount = len(list(sheet.columns()))
        colName = colnum_string(colCount)
        rowCount = len(list(sheet.rows())) - 1
        tableDelimiters = 'A1:' + str(colName) + str(rowCount)
        workbook = xlsxwriter.Workbook(sitexlsx)
        worksheet1 = workbook.add_worksheet(sitecsv)
        worksheet1.add_table(tableDelimiters,{'data': dataList, 'columns': theHeaders})
        workbook.close()
