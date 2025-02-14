' Define variables
Dim city, flightDate, flightClass, passengers, passengerName, agentName, password,tickets
Dim screenshotPath
Dim i
' Define the folder path for screenshots
screenshotPath = "D:\Screenshots\"
' Loop through each row in the Data Table
For i = 1 To DataTable.GetRowCount
' Set the current row for each iteration
    DataTable.SetCurrentRow(i)
' Fetch values from the Data Table
    agentName = DataTable.Value("agentName", dtGlobalSheet)
    password = DataTable.Value("password", dtGlobalSheet)
    city = DataTable.Value("City", dtGlobalSheet)
    flightDate = DataTable.Value("Date", dtGlobalSheet)
    flightClass = DataTable.Value("Class", dtGlobalSheet)
    passengers = DataTable.Value("Passengers", dtGlobalSheet)
    passengerName = DataTable.Value("Passenger Name", dtGlobalSheet)
    tickets = DataTable.Value("Tickets", dtGlobalSheet)
    
' Open the application
    SystemUtil.Run "C:\Users\HP\OneDrive\Desktop\Flight GUI.lnk"
    
' Screenshot before login
    WpfWindow("OpenText MyFlight Sample").CaptureBitmap screenshotPath & "Before_Login_" & i & ".png", True
    
' Login to the application
    WpfWindow("OpenText MyFlight Sample").WpfEdit("agentName").Set agentName
    WpfWindow("OpenText MyFlight Sample").WpfEdit("password").SetSecure password
    WpfWindow("OpenText MyFlight Sample").WpfButton("OK").Click
    
'Text Checkpoint 1
  WpfWindow("OpenText MyFlight Sample_2").WpfObject("Tickets").Check CheckPoint("Tickets")
  
' Screenshot after login
    WpfWindow("OpenText MyFlight Sample").CaptureBitmap screenshotPath & "After_Login_" & i & ".png", True
    
' Select flight details
    WpfWindow("OpenText MyFlight Sample").WpfComboBox("toCity").Select city
    WpfWindow("OpenText MyFlight Sample").WpfImage("WpfImage").Click 3,7
    WpfWindow("OpenText MyFlight Sample").WpfCalendar("Su").SetDate flightDate
    WpfWindow("OpenText MyFlight Sample").WpfImage("WpfImage").Click 2,14
    WpfWindow("OpenText MyFlight Sample").WpfComboBox("Class").Select flightClass
    
  ' Added the manual object
    WpfWindow("OpenText MyFlight Sample_2").WpfComboBox("flightTktManual").Select tickets
    
  ' Screenshot before select flight
    WpfWindow("OpenText MyFlight Sample").CaptureBitmap screenshotPath & "Before_Find_Flights_" & i & ".png", True
    
    WpfWindow("OpenText MyFlight Sample").WpfButton("FIND FLIGHTS").Click
   
 

' Intentional Fail Checkpoint (FAIL only once) 3
If i=2 Then    
 WpfWindow("OpenText MyFlight Sample_2").WpfObject("Flight").Check CheckPoint("Bus")
End If

' Take screenshot after selecting flight details
    WpfWindow("OpenText MyFlight Sample").CaptureBitmap screenshotPath & "After_Find_Flight_" & i & ".png", True
    
' Select flight and enter passenger details
    WpfWindow("OpenText MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,2
    
' Standard Checkpoint (PASS) - Checking if BOOK button is enabled before clicking 4
    WpfWindow("OpenText MyFlight Sample_2").WpfButton("SELECT FLIGHT").Check CheckPoint("SELECT FLIGHT_2")
    
    WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").Click
    WpfWindow("OpenText MyFlight Sample").WpfEdit("passengerName").Set passengerName

' Screenshot before booking
 WpfWindow("OpenText MyFlight Sample").CaptureBitmap screenshotPath & "Before_Booking_" & i & ".png", True
 
' Click ORDER to complete booking
    WpfWindow("OpenText MyFlight Sample").WpfButton("ORDER").Click
    
    wait(5)
    
' Take screenshot after booking
    WpfWindow("OpenText MyFlight Sample").CaptureBitmap screenshotPath & "After_Booking_" & i & ".png", True
    
' Close the application
    WpfWindow("OpenText MyFlight Sample").Close
Next
' Display final results in the test report
  Reporter.ReportEvent micPass, "Automation Execution", "Test completed successfully"
