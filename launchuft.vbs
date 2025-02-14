Dim qtApp, qtTest
Set qtApp = CreateObject("QuickTest.Application")
 
qtApp.Launch
qtApp.Visible = True
 
' Open the test
qtApp.Open "C:\Users\HP\OneDrive\Desktop\ASNM1-UFT-Final\Test.tsp", True
 
' Run the test
Set qtTest = qtApp.Test
qtTest.Run
qtTest.Close
 
qtApp.Quit
 
Set qtTest = Nothing
Set qtApp = Nothing