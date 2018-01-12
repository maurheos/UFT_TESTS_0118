Dim oExtern
Set oExtern = Extern
oExtern.Declare micInteger,"GetPrivateProfileString", "kernel32.dll","GetPrivateProfileStringA", _
				micString, micString, micString,micString+micByRef, micInteger, micString

call oExtern.GetPrivateProfileString("ProgramInformation", "ConfigFileVersion", "", RetVal, 255, "C:\Program Files (x86)\HPE\Unified Functional Testing\bin\wrls_ins.ini")
print RetVal
set oExtern = nothing			
				
