# BSV Job Search Final Metrics



## Foreword

This document outlines the correct uses for a script written in Python meant to streamline the process in which data analysis on job search demographics/statistics is conducted.

This script was developed during my time at Breakthrough Silicon Valley and is intended for continued use by their organization.



## Function

* Reads data from first tab of spreadsheet listing info about all applicants
* Analyzes data and outputs metrics to a “Final Metrics” tab that is created
* Uses coding language Python 
* 2 versions of the code are available:
  - 2 stage interview process
  - 3 stage interview process
  - Each code performs the same processes with one version adding an extra round of interviews taken into account

**Note:** A lot of the code is hard coded. This means that the code only operates under very specific parameters that have been set and will be detailed below. It is important that every parameter is followed for the code to work properly.



## Important Setup

### Downloading programs to run Python

**1. Download Anaconda 3 :** https://www.anaconda.com/distribution/
 - Follow instructions to install anaconda on computer
 - Note: Anaconda automatically installs Python - coding language 
 
**2. Download Sublime Text :** https://www.sublimetext.com/ 
 - Follow instructions to install program on computer

**3. Open terminal :** 
 - Click search bar at top of personal computer and search “terminal”

**4. Type : pip install gspread**
 - Wait for download and “successfully installed” prompt

**5. Type : pip install oauth2client**
 - Wait for download and “successfully installed” prompt

## Allowing code to access google sheets
### Google account 
 - In order for the code to access google sheets, I created a new gmail account which has already been configured to access the data sheets that we fill manually fill out with applicant info. The account info is below but there’s no need to login and change anything.
To run the code on a respective google sheet, share the document with the email address: final-metrics@final-metrics.iam.gserviceaccount.com

### Gmail Account Info:

**Email:** BTSVcode@gmail.com

**Password:** BTSVpython



## Running the code

### Code files
 - The code and other necessary files are stored within the folder shared with you called “BTSVcode”
 - This entire file must be saved anywhere on your computer.

### Running code
**1. Open code in sublime**
   - To access the code, open sublime and open the file called “finalmetrics.py” within sublime. Opening it any other place will not work. 

**2. Change line 15**
   - Line 15 contains the line of code which tells the code which spreadsheet to access in google docs. Type the name of the spreadsheet which you want to run the code on. The input is case/space,character sensitive so the name of the spreadsheet must be typed in exactly. Insert the name of the spreadsheet between the ‘ ‘ marks.
   - Ex.  ->   file = client.open('Dev/Comm') 
   - Ex.  ->   file = client.open(‘SDPI- Recruitment Metrics’)

**3. Run the code**
 - At the top click: tools -> build
 - The code should run. It will take roughly 5 minutes to fully complete. The small text box underneath the code in sublime will display “success” if the code finishes correctly. You can also view the code running in real time in the spreadsheet. A new tab will be created called “Final Metrics” and info will be printing.



## Intricacies and Limitations of this code

### Format of Applicant Info Tab
 - The code pools the data that it analyzes from only the first tab. The first tab in the spreadsheet should be correctly formatted (detailed below) with the proper applicant data inputted.
 - These are links to format templates of how the applicant info tab should be set up:
   - 2 Interview Rounds: https://docs.google.com/spreadsheets/d/1wObWThywQbw2HEzSX1MVX5guXa3gRqejYcpkH1or_bo/edit#gid=0
   - 3 Interview Rounds: https://docs.google.com/spreadsheets/d/1JSEwX1i_JBneS8Ms6uW_G08NJu_TJnWBVorbXNSoF6c/edit#gid=0
 - The sections below detail the applicant info tab should be formatted to ensure the code can read it correctly.
 - Also below details the specific options for specific columns that offer various answers

### Overall 2 Round Format:

A: Date of Application

B: First Name

C: Last Name

D: Email

E: Number

F: Source of Hire

G: Last Contact

H: Phone Interview Conducted?

I: Phone Screen Status

J: Phone Screen Notes

K: 1st Round Conducted?

L: 1st Round Status

M: 1st Round Notes

N: 2nd Round Conducted?

O: 2nd Round Status

P: Updates/Notes

Q: Rejection Sent?

R: Keep in contact?

S: Extra Notes

T: Gender

U: Race/Ethnicity

V: Veteran

W: Disability

X: Current Location

Y: Desired Salary

Z: Languages


### Overall 3 Round Format:

A: Date of Application

B: First Name

C: Last Name

D: Email

E: Number

F: Source of Hire

G: Last Contact

H: Phone Interview Conducted?

I: Phone Screen Status

J: Phone Screen Notes

K: 1st Round Conducted?

L: 1st Round Status

M: 1st Round Notes

N: 2nd Round Conducted?

O: 2nd Round Status

P: 2nd Round Notes

Q: 3rd Round Conducted

R: 3rd Round Status

S: Updates/Notes

T: Rejection Sent?

U: Keep in contact?

V: Extra Notes

W: Gender

X: Race/Ethnicity

Y: Veteran

Z: Disability

AA: Current Location

AB: Desired Salary

AC: Languages


### Gender options:

- Male
- Female
- Nonbinary
- Decline to Specify
- No Data

### Race/Ethnicity Options: 

- Hispanic/Latino
- White
- Black/African American
- Native Hawaiian/Other Pac Islander
- Asian
- American Indian/Alaska Native
- Two or More Races
- Decline to Specify
- No Data

### Veteran Options:

- No Data
- Disabled Veteran
- Armed Forces Service Medal Veteran
- Veteran who served on active duty during a war or in a campaign or expedition for which a campaign badge has been authorized
- Recently Separated Veteran
- Veteran (but not one of the four protected classes above)
- Decline to state
- Not a Veteran

### Disability Options:

- No Data
- Not
- Yes 
- Decline to Answer

**Note:**
These above drop down list options can be created by clicking the top of the row (to format the entire row) and selecting the “data validation” button. Then select list of options and manually type in each option separated by commas.

**Important:** The code has specific columns that it reads from to get certain info, therefore if any extra columns are added/deleted the code won’t work.

### Rules for inputting data into certain columns

**Source of Hire:** If Trinet says “Our job board” → input “Breakthrough”

**Desired Salary:**
 - Input “$” followed by desired salary amount of triNet.
   - Ex. input $60,000  not  60,000
 - If TriNet has a $/per hour value, calculate the yearly salary this would be and input this into the desired salary column/cell
   - Ex. $40/hr  →  $83,200

**Phone Screen/ 1st Round/ 2nd Round Status/ 3rd Round Status:**
 - Use this column to indicate whether an email needs to be sent, interview scheduled, or any other indication of the current step or steps needed to be taken for this applicant at the respective stage. The code doesn’t use these columns for everything but this is just what this column should be used for, for the reader.



## Functions of the Code

This code pools data from the first applicant tab, as outlined above, performs functions on this data, and outputs this data onto a tab that it creates within the existing spreadsheet. 

When run, the code will create a tab called “Final Metrics” in the spreadsheet. An error will occur if a tab already exists with this name so don’t create the tab yourself, the code will make it.

### Functions

**Hiring Stats:**

Counts of number of applicants that applied through each method of applying and prints them in a table

**Applicant / Gender Stats:**

Counts the number of applicants that belong to each respective gender as well as the total number of applicants that applied and prints them to a table.

**Ethnicity/Race Stats:**

Counts the number of applicants that belong to each respective ethnicity/race and prints them to a table.

**Veteran Stats:**

Counts the number of applicants that belong to each respective category of veteran and prints them to a table.

**Disability Stats:**

Counts the number of applicants that belong to each respective category of disability and prints them to a table.

**Mean Salary:**

Adds the total desired reported salaries of all applicants the finds the average reported value.

**Gender by Stages:**

Counts number of applicants that identify as each gender with respect to each interview stage

**Ethnicity by Stages:**

Counts number of applicants that identify as each ethnicity with respect to each interview stage

**Gender and Ethnicity by Stages:**

Counts the number of applicants that belong to each respective category of ethnicity, gender, and each interview stage and prints them to a table.


**Time to Hire:**

This is a manual value that must be computed by the user. The code adds the label but can’t figure this value out on its own. The user must take the post date of the job posting and count the number of days it was visible until the role was filled.
