# -*- coding: utf-8 -*-
"""
Created on Thu Apr  9 11:25:29 2020

@author: lecam
"""
import os,sys
import pyodbc
import pandas as pd
import time, calendar,locale
import openpyxl
import datetime as dt
import shutil
import dateutil.parser

from pandas import DataFrame
from time import strftime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from datetime import datetime
from datetime import timedelta

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders

#######################################################
#  variables temporelles
locale.setlocale(locale.LC_ALL, '')
a =datetime.now()
y = a.year
m= a.month
d= a.day
datefile = a.strftime('%Y%m%d')
datefile2 = a.strftime('%Y%m%d_%H%M')

#######################################################
#autres variables
global Mois_Trt,nb,nb_maj_avant,nb_maj_apres
Mois_Trt,nb,nb_maj_avant,nb_maj_apres = "",0,0,0

#######################################################
#fonction d'execution de procedure stockée
def exec_procedure(nomproc, param,  typeproc):
    global nb
    nb = 0
    conn = pyodbc.connect('DSN=EXTENSIONS__RW_64;')
    cursor = conn.cursor()
    cursor.execute( nomproc, param )
    
    if typeproc == 'S':
        row = cursor.fetchone()
        while row:
            nb = str(row[0]) 
            row = cursor.fetchone()
    
    conn.commit() 
    cursor.close()
    del cursor
    conn.close()
##########################################################################
#  fonction de traitement
def traitement(log):
    global Mois_Trt,nb,nb_maj_avant,nb_maj_apres
    Mois_Trt = a.strftime('%m')
    ######################################################COUNT AVANT UPDATE
    exec_procedure("Exec sp_FG_Construction @mois = ?", (str(Mois_Trt)),  "S")
    nb_maj_avant = nb
    ###########################################################UPDATE
    exec_procedure("Exec sp_FG_Construction_update @mois = ? , @codeFG = ?", (str(Mois_Trt), "RCB19"), "U")
    
    ###########################################################COUNT APRES UPDATE
    exec_procedure("Exec sp_FG_Construction @mois = ?", (str(Mois_Trt)),  "S")
    nb_maj_apres = nb

traitement("log.txt")