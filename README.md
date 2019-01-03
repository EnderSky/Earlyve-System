# Earlyve-System

Repository containing the Earlyve system.

Tracks students who take early leave from school or are late for school.
	

## How to Install
1. Download Python 3.7.0 from www.python.org
2. Download the required files from this Github Repository by clicking "Download" and "Download ZIP"
3. Right-click the ZIP Folder and select "Extract All"
4. Open Command Prompt
5. Copy and Paste this line into Command Prompt: _pip3 install openpyxl wxpython --user_

## How to Operate
1. Always ensure that the application is in the "Latest Version" folder, together with the "Assets" folder.
2. Navigate to the "Assets" folder, followed by "Email Details.txt". 
Change the email details to the email which you will be using for this program.
    The format for that text file is as follows:
	(Email address)
	(Password)
	(smtp email server)
3. Ensure that the student database is saved as "Student Database.csv"
4. Ensure that the log file is saved as "Data Log.xlsx"
5. Ensure that all the required variables in the "Email Template" files (denoted by the curly brackets "{}") are present.
The format of the email can be changed within the corr as long as the required variables are present.
List of Required Variables:
- {formTeacher_one} --> Name of first form teacher
- {formTeacher_two} --> Name of second form teacher
- {name} --> Name of student
- {student_class} --> Student's class
- {date} & {time} --> Date & Time when student took early leave / was late for school
- {reason} --> Reason student gave for taking early leave / coming late for school

## How to Run
1. Open the Folder containing the project "Earlyve-System-master" followed by "Latest Version".
2. Double click the EarlyveSystemApp application and you're good to go!
3. Press "Home" followed by "Help" on the menu bar to understand how to utilise this program.