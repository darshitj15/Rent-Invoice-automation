# Rent-Invoices-Automation
## Summary
This is a python program to generate monthly rent invoices in word file and upload it on google drive using rest API. 
The program uses tkinter library for the GUI, docx to create the word file, and pydrive library to upload on google drive. The program is error proof and user-friendly. It is useful for all the property owners who need to generate a bill for their tenants, calculate different expenses, and then upload it on drive sharing permissions with the rental. This is a very common use as most of property owners have to do this every month and this program will help them automate their process.
## Prerequisites
* Before you continue, ensure you have installed the Python version 3.7.
## Installation
If you are running your program on cmd or tools which requires additional installation of Python libraries, here is the list of libraries that you should install to run this program.
* docx: This library add inputs into a word file in an existing one or by creating a new one. Use the command 'pip install python-docx' to install this library using your command prompt.
* PyDrive: This library is used to make changes in the google drive using rest APIs. Use the command 'pip install PyDrive' to install this library using your command prompt.
* tkinter: This library enables a GUI with which the user will interact and give inputs. Its the standars library which comes with Python installation. If you are using PyCharm, you might need to install the future module to use tkinter.
* client_secrets.json file: Put this file in the same folder as your python file if you are using PyCharm or in your current directory where the word file will be saved if you use the terminal. You can enable your own google drive API from https://developers.google.com/drive/api/v3/enable-drive-api or use the file I have provided in this project's repository.


 
