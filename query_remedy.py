import pyodbc
import pandas as pd

# En el administrador de origenes de bases de datos ODBC 32bit, 
# ya tengo configurado el servidor de conexion a AR System ODBC Data Source

conxn = pyodbc.connect('DSN=AR System ODBC Data Source',autocommit=True) 
cursor = conxn.cursor()

#aca puedo cambiar las querys por la que se necesiten

query_inc = "SELECT HPD_Help_Desk.Incident_Number,\
HPD_Help_Desk.Service_Type, \
HPD_Help_Desk.Description \
FROM HPD_Help_Desk HPD_Help_Desk \
WHERE (HPD_Help_Desk.Reported_Date>{ts '2016-11-01 00:00:00'}) \
AND (HPD_Help_Desk.Status <> 'Closed') \
ORDER BY HPD_Help_Desk.Incident_Number DESC"


query_tas = "SELECT TMS_Task.RootRequestID,\
TMS_Task.Task_ID,\
TMS_Task.Create_Date,\
TMS_Task.Status,\
TMS_Task.Assignee,\
TMS_Task.Assignee_Group,\
TMS_Task.Assignee_Organization \
FROM TMS_Task TMS_Task \
WHERE (TMS_Task.Create_Date >= {ts '2016-01-01 00:00:00'}) \
AND (TMS_Task.Status <> 'Closed') \
ORDER BY TMS_Task.Task_ID DESC"

#almacena las los datos consultados en un DataFrame de Pandas
#con esto ya puedo manipular los datos

data_inc = pd.read_sql(query_inc,conxn)
data_tas = pd.read_sql(query_tas,conxn)

#cierro la conexion a remedy
conxn.close()


# cruzo los datos segun el numero de incidente (une las dos consultas)
data = pd.merge(data_tas, data_inc, left_on=data_tas.iloc[:,0], right_on=data_inc.iloc[:,0], how='left')
#unir las consultas crea columnas duplicadas, por eso las borro
data = data.drop(['key_0','Incident_Number'], 1)


# creo un archivo excel
writer = pd.ExcelWriter('remedy.xlsx')
# guardo los datos en hojas distintas
data_inc.to_excel(writer, sheet_name='inc',index=False)
data_tas.to_excel(writer, sheet_name='tas',index=False)
data.to_excel(writer, sheet_name='remedy',index=False)
writer.save()
writer.close()
