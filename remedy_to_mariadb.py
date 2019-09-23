# remedy_to_mariadb.py
# -*- coding: utf-8 -*-
"""
Query de tareas de las GOA enviadas a MariaDB

SqlAlchemy necesita instalar mysqlclient, pero la manera
tradicional no siempre funciona. La alternativa es descargar el archivo de
https://www.lfd.uci.edu/~gohlke/pythonlibs/#mysqlclient
y luego instlarlo con pip
pip install mysqlclient-1.4.4-cp37-cp37m-win32.whl


"""

import pyodbc
import pandas as pd
import sqlalchemy as sql

# En el administrador de bases de datos, ya tengo configurado el servidor
# de conexion a remedy, con ip, usuario, pass, puerto

conxn = pyodbc.connect('DSN=AR System ODBC Data Source',autocommit=True) 
cursor = conxn.cursor()

# Query de ejemplo
query = "SELECT HPD_Help_Desk.Incident_Number,\
HPD_Help_Desk.Service_Type, \
HPD_Help_Desk.Description \
FROM HPD_Help_Desk HPD_Help_Desk \
WHERE (HPD_Help_Desk.Reported_Date>{ts '2016-11-01 00:00:00'}) \
AND (HPD_Help_Desk.Status <> 'Closed') \
ORDER BY HPD_Help_Desk.Incident_Number DESC"

data = pd.read_sql(query,conxn)

conxn.close()

connect_string = 'mysql://user:password@#.#.#.#/dbname'

sql_engine = sql.create_engine(connect_string)

data.to_sql('db_tabla', con=sql_engine, if_exists='replace',
                index_label='id')




