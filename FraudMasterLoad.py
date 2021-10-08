#!/usr/bin/env python
#Import the libraries
import pandas as pd
import pyodbc
from sqlalchemy import create_engine
import logging
import sys

logging_format = '%(levelname)s: %(asctime)s: %(message)s'
logging.basicConfig(level=logging.DEBUG, format=logging_format)
logger = logging.getLogger()

handler = logging.FileHandler('test1.log')
# create a logging format
formatter = logging.Formatter(logging_format)
handler.setFormatter(formatter)

logger.addHandler(handler)

#Give the connection info
sqlcon = create_engine('mssql+pyodbc://MIUser:B3nn3tt5!@bendb02.ukfast.bennetts.co.uk/Bennetts_MI?driver=SQL+Server')

with sqlcon.begin() as conn:
         conn.execute("""TRUNCATE TABLE Bennetts_MI.Fraud.T_Fraud_Master""")

logger.info("TRUNCATE TABLE Bennetts_MI.Fraud.T_Fraud_Master")

df = pd.read_excel (r'\\benfile01\bennetts\Fraud\Fraud Capture Master.xlsx', sheet_name='Master')

logger.info("Fraud Capture Master.xlsx read")

dfnew = df[["Policy Number", "Date Added"]]

logger.info("Inserting Data to T_Fraud_Master")

dfnew.to_sql(name='T_Fraud_Master',index=False,schema='Fraud',con=sqlcon,if_exists='append')

logger.info("Completed Inserting data to T_Fraud_Master")
