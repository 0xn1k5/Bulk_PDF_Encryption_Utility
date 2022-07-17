@echo off
echo.
echo ====================================================================
echo ---------------------: Bulk PDF Encryption v0.1 :-------------------
echo ---------------------: By Nikhil Raj (0xn1k5) - 07/2022 :-------------------
echo ====================================================================
echo.
echo.
echo Step 1: Copy the PDF files in 'OriginalPDF' folder for Encryption
echo Step 2: Paste the PDF Filename (Col A) and password (Col B) in excel file under 'PasswdList' folder
echo Step 3: Choose 'Y' to  proceed 
echo.
set /P INPUT=Would you like to proceed (Y/N): %=%
If /I "%INPUT%"=="y" goto yes
If /I "%INPUT%"=="n" goto no
:yes
Powershell.exe -executionpolicy Bypass -File  pdfEncrypt_v1.0.ps1 
echo "File Encryption Completed! 
echo "Check the encrypted PDF files under 'EncryptedPDF' folder" 

:no
echo See you later!
@pause
