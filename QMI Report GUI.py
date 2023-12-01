from __future__ import print_function
import pandas as pd
import os  
from pathlib import Path
import re
from datetime import date
import numpy as np
import glob
import json
from argparse import ArgumentParser
from gooey import Gooey, GooeyParser

@Gooey(program_name="Create Weekly QMI Report")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Create Weekly QMI Report')
    parser.add_argument('RVR_IOPD',
                        action='store',
                        default=stored_args.get('RVR_IOPD'),
                        widget='FileChooser',
                        help='RVR Inter Org Past Due')
    parser.add_argument('JEF_IOPD',
                        action='store',
                        default=stored_args.get('JEF_IOPD'),
                        widget='FileChooser',
                        help='RVR Inter Org Past Due')
    parser.add_argument('TOP_750',
                        action='store',
                        default=stored_args.get('TOP_750'),
                        widget='FileChooser',
                        help='TOP_750 on SAS')
    parser.add_argument('RVR_Firm_Order_Report',
                        action='store',
                        default=stored_args.get('RVR_Firm_Order_Report'),
                        widget='FileChooser',
                        help='RVR Firm Order Report')
    parser.add_argument('JEF_Firm_Order_Report',
                        action='store',
                        default=stored_args.get('JEF_Firm_Order_Report'),
                        widget='FileChooser',
                        help='JEF Firm Order Report')
    parser.add_argument('output_directory',
                        action='store',
                        widget='DirChooser',
                        default=stored_args.get('output_directory'),
                        help="Output directory to save QMI report")
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args


def combineIOPD(RVR: str, JEF: str):
    #rvrDues = pd.read_csv("rvr1.csv", delimiter = ",", index_col = 0)
    rvrDues = pd.read_excel(RVR)
    jefDues = pd.read_excel(JEF)

    #remove header for RVR
    new_header = rvrDues.iloc[0] #grab the first row for the header
    rvrDues = rvrDues[1:] #take the data less the header row
    rvrDues.columns = new_header #set the header row as the df header

    #remove header for JEF
    new_header = jefDues.iloc[0] #grab the first row for the header
    jefDues = jefDues[1:] #take the data less the header row
    jefDues.columns = new_header #set the header row as the df header
    print(rvrDues.head())
    print(jefDues.head())
    
    rvrDues = rvrDues[['Order Number', 'Item', 'Item Description', 'Quantity Ordered', 'Ship ORG', 'Past Due Status']]
    jefDues = jefDues[['Order Number', 'Item', 'Item Description', 'Quantity Ordered', 'Ship ORG', 'Past Due Status']]
    rvrDues = rvrDues.loc[~(rvrDues['Past Due Status'] == 'Future order') & ~(rvrDues['Ship ORG'] == 'JEF')
                           & ~(rvrDues['Ship ORG'] == 'RVR') & ~(rvrDues['Ship ORG'] == 'DB6') & ~(rvrDues['Ship ORG'] == 'KLC')
                           & ~(rvrDues['Ship ORG'] == 'SLC') & ~(rvrDues['Ship ORG'] == 'WAL') & ~(rvrDues['Ship ORG'] == 'NAP')]
    jefDues = jefDues.loc[~(jefDues['Past Due Status'] == 'Future order') & ~(jefDues['Ship ORG'] == 'JEF')
                           & ~(jefDues['Ship ORG'] == 'RVR') & ~(jefDues['Ship ORG'] == 'DB6') & ~(jefDues['Ship ORG'] == 'KLC')
                           & ~(jefDues['Ship ORG'] == 'SLC') & ~(jefDues['Ship ORG'] == 'WAL') & ~(jefDues['Ship ORG'] == 'NAP')]
    print(rvrDues)
    print(jefDues)
    

    pastDues = pd.concat([rvrDues,jefDues]).reset_index(drop=True)
    print(pastDues)
    print("Quantity Ordered Total: " + str(rvrDues['Quantity Ordered'].sum() + jefDues['Quantity Ordered'].sum()))

    

    #use groupby to determine the best statistics and insights 
    return pastDues

def externalPD(RVR:str, JEF:str):
    rvrDues = pd.read_excel(RVR)
    jefDues = pd.read_excel(JEF)

    #remove header for RVR
    new_header = rvrDues.iloc[0] #grab the first row for the header
    rvrDues = rvrDues[1:] #take the data less the header row
    rvrDues.columns = new_header #set the header row as the df header

    #remove header for JEF
    new_header = jefDues.iloc[0] #grab the first row for the header
    jefDues = jefDues[1:] #take the data less the header row
    jefDues.columns = new_header #set the header row as the df header
    print(rvrDues.head())
    print(jefDues.head())
    jefDues = jefDues.iloc[:-6]
    pastDues = pd.concat([jefDues,rvrDues]).reset_index(drop=True).iloc[:-6].replace(np.nan, '')
    pastDues = pastDues.loc[~(pastDues['Line Type'] == "Shipment")] #filters out shipments
    pastDues = pastDues.loc[~pastDues['Item Number'].str.contains('^PP[a-z]*', flags=re.I, regex=True)] #filters out prepack
    pastDues['Due Date'] = pd.to_datetime(pastDues['Due Date'], format='%Y-%m-%d')
    pastDues = pastDues[pd.notnull(pastDues['Due Date'])] #gets rid of blanks
    pastDues = pastDues.loc[(pastDues['Due Date'] >= '2020-01-01')
                     & (pastDues['Due Date'] <= date.today().strftime("%Y-%m-%d"))].reset_index() #filters only dates between 2020 and the day you are running the file
    pastDues['NEED_BY_DATE'] = ''
    pastDues['Intransit Quantity'] = pastDues['Intransit Quantity'].astype(int)
    pastDues['QTY_OPEN'] = pastDues['Quantity Ordered'].subtract(pastDues['Intransit Quantity']).subtract(pastDues['Received Quantity']).astype(int)
    pastDues = pastDues[['PO Number', 'Release','PO Line Number', 'Shipment Number', 'Vendor Name', 'Ship From', 'Buyer', 'Planner Code',
                         'Item Number', 'Description', 'Supplier Item', 'Due Date', 'NEED_BY_DATE', 'Quantity Ordered', 'Intransit Quantity', 'Received Quantity', 'QTY_OPEN']]
    print(pastDues)
    #pastDues.to_excel("EPD3.xlsx", index=False)

    return pastDues

def BObySupplier(file:str): 
    df = pd.read_excel(file)
    df['TOTAL_BO_PIECES'] = df['TRUE_BO'].subtract(df['BO_PP']).astype(int)
    df['TOTAL_BO_LINES'] = ''
    df['TECH_BO_PIECES'] = df['TRUE_TECH_BO'].subtract(df['TECH_BO_PP']).astype(int)
    df['TECH_BO_LINES'] = ''
    df = df[['VENDOR_NAME', 'VENDOR_SITE', 'ITEM_ID', 'DESCRIPTION', 'COMP_NAME','TOTAL_BO_PIECES', 'TOTAL_BO_LINES', 'TECH_BO_PIECES', 'TECH_BO_LINES']].iloc[:-1].replace(np.nan, '')
    
    print(df)
    return df


#VISION: After combining all three documents, I want it to append a row with various statistics on total BO and PD from previous weeks to a running excel file that updates
#some statistics could include 
if __name__ == '__main__':
    conf = parse_args()
    IOPD = combineIOPD(conf.RVR_IOPD, conf.JEF_IOPD) #IOPD Sheet
    BBS = BObySupplier(conf.TOP_750)
    EPD = externalPD(conf.RVR_Firm_Order_Report, conf.JEF_Firm_Order_Report)

    today = date.today()
    filepath = conf.output_directory+"/"
    with pd.ExcelWriter(filepath + today.strftime("%m.%d.%Y") + " QMI Report.xlsx") as writer:
        EPD.to_excel(writer, sheet_name = 'External Past Due', index=False)
        BBS.to_excel(writer, sheet_name= 'BO By Supplier', index=False)
        IOPD.to_excel(writer, sheet_name= 'Internal Past Due', index=False)
