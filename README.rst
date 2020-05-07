.. image:: https://i.pinimg.com/originals/31/6c/eb/316ceb2b81248f951926e806ecb6e8a9.gif
   :align: center
   :target: https://powerbi.microsoft.com/en-us/
   :alt: Power BI Logo
============
Power BI refresher
============
Script for automation of refreshing Power BI workbooks. Built on Python 3.7 and pyautogui.

Developed for Power BI Desktop May 2019 Update (.69.5467.2151) on Windows 10 64-bit with English locale.

============
Installation
============
Install using pip

pip install opencv-python <br/>                                                                                                          pip install pyautogui

you can find in requirements.txt file


============
Usage  
============
python sql_stored_proc_update.py    <br/>                                                                                                      
WHERE<br/>
python <py file> is file which es refreshing each 5 minutes sql_stored procedure and creates new table removing old one.<br/>
If there are new loan applications found, then it automatically runs pbix. file, refresh and save it, then publish and close.
  And this process is going and going recursively till it again finds new applications for loan.

Keep in mind that that user of pbix should be logged in order to be able to publish it.     <br/>                                        Please keep in mind that this script uses GUI of Power BI Desktop and it needs that a user is logged in Windows session. 

# how it works

See how it works
https://drive.google.com/file/d/1jpfOyHJR-MH7CNLJtec6hDsxo4zpRHWR/view

# Bug reporting
Create Github issue. Please write version of your Power BI Desktop, OS and attach command line result and screenshot.
