# Python-OracleDB
Python script to connect oracle database and run the record count query to check the volume received in ETL job.

## Pre-requirements:

Table creation - pip install pandas
Oracle Db connection - pip install cx_Oracle
Sending Mail - pip install pypiwin32

## Details :
The volume-check.py file  builds a connection with Orcale Database using credentials, then the query is excetued for fetching the record count in database tables.
The volume count obtained is then shared as email to the receipents with baseline volume and a HIGH/LOW status indication.
The mail is formatted with tables with pandas library.
