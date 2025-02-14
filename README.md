# UFT-One-Automation
This is my INFO-6255 Project- Flight GUI

This project is an implementation of a Data-driven Automation Framework using UFT (Unified Functional Testing). The script automates the process of logging into an application Flight GUI.exe, filling out forms, and validating the results using checkpoints. The project adheres to the requirements provided and includes additional features such as screenshots and Windows scheduler integration.

---

## Project Requirements

1. **Data-driven Automation Framework**: 
   - No hard-coded values are used in the script. All data is fetched from an external data file (Excel).
   
2. **Object Repository**:
   - The script uses the Object Repository. One object(The Tickets) was manually added to the repository, and the associated code was manually inserted into the script.

3. **Checkpoints**:
   - The script includes a total of 3 checkpoints:
     - 2 checkpoints pass.
     - 1 checkpoint fails (only once).
     - Checkpoint types include:
       - 1 Text checkpoint.

4. **Comments**:
   - Adequate comments are added throughout the code for clarity and better understanding.

5. **Test Data and Screenshots**:
   - Legible and meaningful test data values are used.
   - Screenshots are taken:
     - BEFORE and AFTER filling out each form.
     - After the completion of the script.

6. **Repetition**:
   - The script runs 4 times for 4 different data rows in the data sheet (for 4 different reservations).

7. **Login**:
   - The script logs into the application for each data sheet row (4 times).

8. **Live Execution**:
   - The script will be executed live during the class presentation, and the results will be shown in UFT.

9. **Test Report**:
   - The test report shows all tests passed/failed.

---


  - The UFT automation is configured to run using the Windows scheduler.

---

## Project Structure
UFT_Project/
│
├── Data/ # Contains the data file for data-driven testing
│ └── TestData.xlsx # Excel file with test data
│
├── ObjectRepository/ # Contains the Object Repository files
│ └── OR.tsr # Object Repository file
│
├── Screenshots/ # Contains screenshots taken during execution
│ ├── BeforeFormFill_1.png # Screenshot before filling form for row 1
│ ├── AfterFormFill_1.png # Screenshot after filling form for row 1
│ └── ... # Additional screenshots
│
├── Scripts/ # Contains the UFT script
│ └── ReservationScript.vbs # Main UFT script
│
├── TestResults/ # Contains test execution reports
│ └── TestReport.html # Test execution report
│
└── README.md # Project documentation (this file)





---

## How to Run the Script

1. **Prerequisites**:
   - UFT (Unified Functional Testing) installed.
   - Test application accessible.
   - Data file (`TestData.xlsx`) updated with valid test data.

2. **Steps**:
   - Open UFT and load the `ReservationScript.vbs` script.
   - Ensure the Object Repository (`OR.tsr`) is linked.
   - Run the script from UFT or configure it to run via Windows Scheduler.

3. **Windows Scheduler Setup**:
   - Create a new task in Windows Task Scheduler.
   - Set the action to run the UFT executable with the script as an argument.
   - Schedule the task to run at the desired time.

---

## Screenshots

I have Uploaded the screenshots in Github inside the screenshots folder

---

## Test Results

The test report generated after execution shows the status of each checkpoint and overall test results. Below is a summary:

| Checkpoint Type | Status  |
|-----------------|---------|
| Text            | Pass    |
| Standard        | Pass    |
| Standard        | Fail    |

---

## Notes

- The script is designed to handle 4 iterations, one for each row in the data file.
- The failing checkpoint is intentionally included to demonstrate failure handling.
- Screenshots and test reports are saved for each iteration.

---

## Author

[Hemin Dhamelia]  
[heminpatel2002@gmail.com | +1 551 384 6855]

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
