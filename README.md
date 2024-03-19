# URI-Database-Developer-Internship
Custom application for URI Pharmacy Outreach Group - this is a Beta / Proof of Concept version <br>
This project was created in response to the URI Pharmacy Outreach Group's need for an application to record and analyze patient outreach in the local community.  The Pharmacy department goes on site to local Adult Daycare Centers to provide services to low income individuals to aid their healthcare. <br>
I performed requirement elicitations and created a front end in VBA in Excel. <br>
There is sample data in the Excel workbook associated with the application.  The database was not developed at the time.
# Insructions for usage:
Necesary files / folders:
| Object | Description |
|--------|-------------|
| Pharm-Outreach-FrontEnd (2).xlsm | executable file |
| "Drop Down Lists Files" | folder with .csv files to populate drop-down lists |

# Run the Excel solution:
Download the .xlsm and the folder with the related .csv files <br>
Open the .xlsm <br>
Click the Developer Tab <br>
Click Visual Basic <br>
Open the frmWelcome object - it is the entry point to the program <br>
Click Run in the top ribbon <br>
When the application opens, it needs the path for the drop down files,  <br>
copy and paste the path for "Drop Down List Files" into the dialog box <br>
You will be taken to the welcome dialog box to start your session
# Note
I wrote a custom class to import .csv files <br>
I was working on a custom class to I/O to a MySql database for patient records <br>
Application colors were for URI
# Known Bugs
The "save to database" dialogs don't perform any actions - project not completed
