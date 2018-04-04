#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Mon Feb 19 22:21:12 2018

@author: yaniv.kusuma@gmail.com
"""

import json, urllib3, time
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def data_updater(ri, clist, l, c):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Cryptoweigh_key.json', scope)
    client = gspread.authorize(creds)
    sh = client.open('Rank Analysis')
    wk = sh.worksheet('API-CoinRank')

    for i in [0,1]:  # two times updates
        cell_list = wk.range(3,ri+1+i+c,l+2,ri+1+i+c) #(start,end)
        for cell in range(len(cell_list)):
            cell_list[cell].value = ''
        wk.update_cells(cell_list)
        for cell in range(len(cell_list)):
            cell_list[cell].value = clist[cell][i]
        wk.update_cells(cell_list)
        
def add_new_coin(new_coins):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Cryptoweigh_key.json', scope)
    client = gspread.authorize(creds)
    sh = client.open('Rank Analysis')
    wk = sh.worksheet('Data Analysis')

    nrow = wk.col_values(1)
    nrow = [i for i in nrow if i != '']
    nrow = len(nrow)
    cell_list = wk.range(nrow+1,1,nrow+len(new_coins)+1,1)
    print (cell_list)
#    exit(1)
    print (new_coins)
    for i in range(len(new_coins)):
        
        print ('Added {} in Data Analysis sheet' .format(new_coins[i]))
        cell_list[i].value = new_coins[i]
    wk.update_cells(cell_list)
    
    
def coin_compare_data_updater(ri, clist, l):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Cryptoweigh_key.json', scope)
    client = gspread.authorize(creds)
    sh1 = client.open('Rank Analysis')
#    ls = sh1.worksheets()
#    print (ls[6])
#    exit(1)    
    wk1 = sh1.worksheet('API-Coincompare')

#    wk1 = sh1.get_worksheet(6)
#    print (wk1)
#    exit(1)


    cell_list = wk1.range(2,2,1000,2) #'[ row col start, row col end]' B2:B1000
    for cell in range(len(cell_list)):
        cell_list[cell].value = ''
    wk1.update_cells(cell_list)

    cell_list = wk1.range(2,2,l+1,2) #'[ row col start, row col end]' B2:B1000
    
    for cell in range(len(cell_list)):
        cell_list[cell].value = clist[cell][0]
    wk1.update_cells(cell_list)
    
    time.sleep(6)
    cmpr_result = wk1.col_values(3)[1:]
    new_coins = []
    for i in range(len(cmpr_result)):
        if cmpr_result[i] == 'No Match Found':
            new_coins.append(clist[i][0])
    add_new_coin(new_coins)
        
def dashboard_data_updater(ri):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Cryptoweigh_key.json', scope)
    client = gspread.authorize(creds)
    sh3 = client.open('Rank Analysis')
    wk3 = sh3.worksheet('Dashboard')

    time.sleep(4)
    values_list = wk3.row_values(1)[2:13]    
    values_list[0]=datetime.today().strftime("%d-%m-%Y %H:%M:%S")
    values_list.insert(0,'Successful')
    cell_list = wk3.range(4+ri,2,4+ri,13)
    print (ri)
    print (cell_list)
    print (values_list)
#    print len(cell_list)
#    print len(values_list)
    for cell in range(len(values_list)):
        cell_list[cell].value = values_list[cell]
    wk3.update_cells(cell_list)

def data_migrate(x):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Cryptoweigh_key.json', scope)
    client = gspread.authorize(creds)
    sh = client.open('Rank Analysis')

#    exit(1)
    wk = sh.worksheet('API-CoinRank')
    c = 0
    for ri in range(0,30): #0,30
        c+=1
      
        for i in [0,1]:  # two times updates
            ccol = wk.col_values(ri+i+1+c)[2:]
            
            l = len(ccol)+2  
            print (3,ri+1+i+c,l,ri+1+i+c)
            cell_list1 = wk.range(3,ri+1+i+c,l,ri+1+i+c) #(start,end)
            for cell in range(len(cell_list1)):
                cell_list1[cell].value = ''
            wk.update_cells(cell_list1)                
            cell_list2 = wk.col_values(ri+1+i+c+2)[2:]
#            cell_list2.append('')
#            cell_list2.append('')
            if len(cell_list2) > len(cell_list1):
                d = len(cell_list2) - len(cell_list1)
                l1 = len(cell_list1) + d + 2
                cell_list1 = wk.range(3,ri+1+i+c,l1,ri+1+i+c) #(start,end)
            elif len(cell_list1) > len(cell_list2):
                d = len(cell_list1) - len(cell_list2) + 2
                cell_list2 = cell_list2 + ['']*d 
                
                
            print (len(cell_list1))
            print (len(cell_list2))
                
            for cell in range(len(cell_list1)):
                cell_list1[cell].value = cell_list2[cell]
            wk.update_cells(cell_list1)
    ri = 30; c = 31
    for i in [0,1]:
        cell_list1 = wk.range(3,ri+1+i+c,l,ri+1+i+c) #(start,end)
        for cell in range(len(cell_list1)):
            cell_list1[cell].value = ''
        wk.update_cells(cell_list1)             
    last_column_data(x)

def last_column_data(x):
    data = requests.get('https://api.coinmarketcap.com/v1/ticker/?limit=3000')
    data = json.loads(data.text)
    clist = []
    for stat in data:
        if stat['market_cap_usd'] and stat['24h_volume_usd'] and float(stat['24h_volume_usd']) > 50000.0000 :
            clist.append([stat['name'],stat['rank']])
    l = len(clist)
    ri = 30  #30
    c= 31   #31

    coin_compare_data_updater(ri, clist,l)
    time.sleep(8)
    print ('migrated data update')
    data_updater(ri, clist, l, c)
    dc = ri
    time.sleep(8)
    print (x)
    print ('migrated dashboard')
    dashboard_data_updater(dc+x)
    
import requests    
def data_timer():    
    print ('data analysis update started')
    c = 0
    for ri in range(0,31):  #0,31

        c+=1
        clist = []
        data = requests.get('https://api.coinmarketcap.com/v1/ticker/?limit=3000')
#        print (data.text)
        data = json.loads(data.text)

        for stat in data:

            if stat['market_cap_usd'] and stat['24h_volume_usd'] and float(stat['24h_volume_usd']) > 50000.0000 :

                clist.append([stat['name'],stat['rank']])

        l = len(clist)
        print (l)
        print ('run %s' %str(ri))
        print ('-> coincompare started')
        coin_compare_data_updater(ri, clist, l)
        time.sleep(5)
        data_updater(ri, clist, l, c)
        time.sleep(5)
        print ('-> dashboard started')
        dashboard_data_updater(ri)
        print ('==Waiting===')
        time.sleep(1800)
    x = 0    
    while True:
        x +=1        
        data_migrate(x)
        print ('========')
        print ('data migrated')
        print ('===Waiting====')
        
        time.sleep(1600)

data_timer()
