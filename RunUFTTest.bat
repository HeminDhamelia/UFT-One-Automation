@echo off
set UFT_PATH="C:\Program Files (x86)\OpenText\UFT One\bin\UFT.exe"
set TEST_PATH="C:\Users\HP\OneDrive\Desktop\ASNM1-UFT-Final\Test.tsp"
 
:: Start UFT and wait for it to fully load
start "" %UFT_PATH%
timeout /t 25 /nobreak >nul
 
:: Use UFT Automation Object Model (AOM) to open and run the test
cscript //nologo "%~dp0launchuft.vbs"