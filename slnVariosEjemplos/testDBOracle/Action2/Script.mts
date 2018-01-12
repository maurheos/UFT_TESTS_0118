Dim objConnection 
'Set Adodb Connection Object
Set objConnection = CreateObject("ADODB.Connection")     
Dim objRecordSet 
 
'Create RecordSet Object
Set objRecordSet = CreateObject("ADODB.Recordset")     
 
Dim DBQuery 'Query to be Executed
DBQuery = "Select * from DBEASYCALCULATION.TBLAGECALCULATOR"
 
'Connecting using SQL OLEDB Driver
'objConnection.Open "Provider = {Microsoft ODBC for Oracle};Server =localhost:1521/XE;User Id = SYSTEM;Password=Sophos.2017;Database = DBEASYCALCULATION"
'objConnection.Open "Provider = MSDAORA;Server =localhost:1521/XE;User Id = SYSTEM;Password=Sophos.2017;Database = DBEASYCALCULATION"
'objConnection.Open "Provider={Oracle in XE};Data Source=(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = SBMEDLAPBANAR7)(PORT = 1521))(CONNECT_DATA =(SID=xe)(SERVER = DEDICATED)(SERVICE_NAME = XE)));User Id=SYSTEM;Password=Sophos.2017"
'objConnection.Open "Driver={Oracle in XE};Server=xe;Uid=SYSTEM;Pwd=Sophos.2017"
objConnection.Open "Driver={Oracle in XE};dbq=localhost:1521/xe;Uid=SYSTEM;Pwd=Sophos.2017"
 
'Execute the Query
objRecordSet.Open DBQuery,objConnection
 
'Return the Result Set
'Value = objRecordSet.fields.item(0)				
'msgbox Value
 objRecordSet.MoveFirst
i=0
Do While Not (objRecordSet.BOF Or objRecordSet.EOF)
    i = i + 1
    With objRecordSet
        ' Asignar a las variables el contenido del registro
        sId = .Fields("IDDATA").Value & ""
        sDia = .Fields("DATEDAY").Value & ""
        sMes = .Fields("DATEMONTH").Value & ""
        sAnno = .Fields("DATEYEAR").Value & ""
        sTuEdad = .Fields("YOURAGEIS").Value & ""
		sTuEdadDias = .Fields("YOURAGEINDAYS").Value & ""
		sTuEdadHoras = .Fields("YOURAGEINHOURS").Value & ""
		sTuEdadMinutos= .Fields("YOURAGEINMINUTES").Value & ""
        
	Print "No. Registro: " + sId + vbCrLf _
		   + "Día: " + sDia + vbCrLf _
		   + "Mes: " + sMes + vbCrLf _
		   + "Año: " + sAnno + vbCrLf _ 
		   + "Tu edad es: " + sTuEdad + vbCrLf _
		   + "Tu edad en días: " + sTuEdadDias + vbCrLf _
		   + "Tu edad en horas: " + sTuEdadHoras + vbCrLf _
		   + "Tu edad en minutos: " + sTuEdadMinutos
    End With
    '
    ' Mostrar el siguiente registro
    objRecordSet.MoveNext
Loop

If i = 0 Then
    Print "No se encontró ningún registro"
End If
 
 
' Release the Resources
objRecordSet.Close        
objConnection.Close		
 
Set objConnection = Nothing
Set objRecordSet = Nothing
