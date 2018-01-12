'Print Environment("TestDir")
'wait(10)
' 
'SystemUtil.Run "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome Beta.lnk", "https://www.diffchecker.com/diff"
''Browser("Computed Diff - Diff Checker").Page("Computed Diff - Diff Checker").WebEdit("xpath:=//DIV/DIV[@class='CodeMirror-lines' and @role='presentation']/DIV[1]/TEXTAREA[1]").highlight
'wait(5)
''Grabe the WebElement
'Set oWebEdit = Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("WebEdit_3")
''Grabe the WebElement's Runtime-Object's CurrentStyle object
'Set oStyle = oWebEdit.Object.style
'
''Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("WebEdit_3").Object.set 
'wait(10)
''Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("xpath:=//*[@id='__next']/div/div/div/div/div[2]/form/div/div[1]/div/div[6]/div[1]/div/div/div/div[5]/div/pre").Set "Ok"
''Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("xpath:=//div[@id='__next']/div/div/div/div/div[2]/form/div/div/div/div/textarea").Set "Ok"
''Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("css:=div.CodeMirror-lines").highlight
''Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("xpath:=//*[@id='__next']/div/div/div/div/div[2]/form/div/div[2]/div/div[6]/div[1]/div/div/div/div[5]/div/pre").Set "ready"
''msgbox Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("WebEdit_3").GetTOProperties
'msgbox Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("WebEdit_3").GetTOProperty("id")
'msgbox Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("WebEdit_3").GetROProperty("id")
'wait(10)
'Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("WebEdit_3").Set "ready"
'Browser("Diff Checker - Online").Page("Diff Checker - Online").WebEdit("WebEdit_4").Set "ready"
'Browser("Diff Checker - Online").Page("Diff Checker - Online").WebButton("Find Difference!").Click
'Browser("Diff Checker - Online").Page("Diff Checker - Online").WebList("expiry").Select "Store for 1 hour"
'Browser("Diff Checker - Online").Page("Diff Checker - Online").WebList("expiry").Select "Don’t store diff"
