SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe", "https://www.easycalculation.com/"
wait(5)
'Browser("Free Online Math Calculator").Page("Free Online Math Calculator").Link("Age Calculator").WaitProperty "name", "Age Calculator", 10000 @@ hightlight id_;_Browser("Free Online Math Calculator").Page("Free Online Math Calculator").Link("Age Calculator")_;_script infofile_;_ZIP::ssf7.xml_;_
Browser("Free Online Math Calculator").Page("Free Online Math Calculator").Link("Age Calculator").Click @@ hightlight id_;_Browser("Free Online Math Calculator").Page("Free Online Math Calculator").Link("Age Calculator")_;_script infofile_;_ZIP::ssf6.xml_;_

Dim contenedorIngreso 
Set contenedorIngreso= Browser("Free Online Math Calculator").Page("Age Calculator Online")
With contenedorIngreso
	.WebEdit("WebEdit").Set parameter("Dia")
	.WebEdit("WebEdit_2").Set parameter("Mes")
	.WebEdit("WebEdit_3").Set parameter("Anno")
	.WebButton("Go").Click
End With

Dim contenedorSalida
Set contenedorSalida= Browser("Age Calculator Online").Page("Age Calculator Online")
With contenedorSalida
	parameter("AgeYears")= .WebEdit("val").GetROProperty("value")
	parameter("AgeDays")= .WebEdit("val1").GetROProperty("value")
	parameter("AgeHours")= .WebEdit("val2").GetROProperty("value")
	parameter("AgeMinutes")= .WebEdit("val3").GetROProperty("value")
End With
'Browser("Free Online Math Calculator").Page("Age Calculator Online").WebEdit("WebEdit").Set parameter("Dia") @@ hightlight id_;_Browser("Free Online Math Calculator").Page("Age Calculator Online").WebEdit("WebEdit")_;_script infofile_;_ZIP::ssf2.xml_;_
'Browser("Free Online Math Calculator").Page("Age Calculator Online").WebEdit("WebEdit_2").Set parameter("Mes") @@ hightlight id_;_Browser("Free Online Math Calculator").Page("Age Calculator Online").WebEdit("WebEdit 2")_;_script infofile_;_ZIP::ssf3.xml_;_
'Browser("Free Online Math Calculator").Page("Age Calculator Online").WebEdit("WebEdit_3").Set parameter("Anno") @@ hightlight id_;_Browser("Free Online Math Calculator").Page("Age Calculator Online").WebEdit("WebEdit 3")_;_script infofile_;_ZIP::ssf4.xml_;_
'Browser("Free Online Math Calculator").Page("Age Calculator Online").WebButton("Go").Click @@ hightlight id_;_Browser("Free Online Math Calculator").Page("Age Calculator Online").WebButton("Go")_;_script infofile_;_ZIP::ssf5.xml_;_
'parameter("AgeYears")=Browser("Age Calculator Online").Page("Age Calculator Online").WebEdit("val").GetROProperty("value")
'parameter("AgeDays")=Browser("Age Calculator Online").Page("Age Calculator Online").WebEdit("val1").GetROProperty("value")
'parameter("AgeHours")=Browser("Age Calculator Online").Page("Age Calculator Online").WebEdit("val2").GetROProperty("value")
'parameter("AgeMinutes")=Browser("Age Calculator Online").Page("Age Calculator Online").WebEdit("val3").GetROProperty("value")

