'DbTable("DbTable").Check CheckPoint("DbTable")

Dim objConnection 
'Set Adodb Connection Object
Set objConnection = CreateObject("ADODB.Connection")     
Dim objRecordSet 
 
'Create RecordSet Object
Set objRecordSet = CreateObject("ADODB.Recordset")     
 
Dim DBQuery 'Query to be Executed
DBQuery = "Select * from DBEASYCALCULATION.TBLAGECALCULATOR WHERE IDDATA=1"
 
'Connecting using SQL OLEDB Driver
 'objConnection.Open "Driver={Oracle in XE};dbq=localhost:1521/xe;Uid=SYSTEM;Pwd=Sophos.2017"
 objConnection.Open "DSN=XE;UID=SYSTEM;PWD=Sophos.2017;DBQ=SBMEDLAPBANAR7:1521/XE"
 
'Execute the Query
objRecordSet.Open DBQuery,objConnection
 
'Return the Result Set
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
'Do Until objRecordSet.EOF
'MsgBox "" & objRecordSet. 
'objRecordSet.MoveNext
'Loop
'Print objRecordSet.RecordCount
'Print objRecordSet.Fields.Count
'Value = objRecordSet.fields.item(0)
'msgbox Value
 
' Release the Resources
objRecordSet.Close        
objConnection.Close		
 
Set objConnection = Nothing
Set objRecordSet = Nothing

'RunAction "Action3", oneIteration, sDia, sMes, sAnno, age, days, hours, minutes
RunAction "Action3", oneIteration, "1", "1", "1991", age, days, hours, minutes
Print "Tu edad actual es: " + age + vbCrLf _ 
	  + "Tu edad actual en días: " + days + vbCrLf _
	  + "Tu edad actual en horas: " + hours + vbCrLf _
	  + "Tu edad actual en minutos: " + minutes
