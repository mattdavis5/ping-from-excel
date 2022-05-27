# ping-from-excel
Python script that iterates through an Excel file and performs a ping command for each host/IP address.
<br />
<br />
<br />
## Get Started
Before running this script, ensure Python version 3.6 or newer is installed on your machine.
<br />
<br />
### Install Libraries
To install libraries required to run this script, perform the following instructions:
1. Open Command Prompt or other terminal 
2. Install the [openpyxl library](https://openpyxl.readthedocs.io/en/stable/) by entering-
  ```cmd
  pip install openpyxl
  ```
3. Install the [pythonping library](https://pypi.org/project/pythonping/) by entering-
  ```cmd
  pip install pythonping
  ```
  


## User Input

ping-from-excel.py requires the user to input the following:

1. A valid Excel file path
2. The file path for a new or existing .txt file to log ping details
3. The column which lists hosts/IPs to ping
4. The starting row which lists hosts/IPs (e.g. with a table header in row 2, row 3 is the starting row)

<br />

## Script Output

ping-from-excel.py will output:
   
* A text file with each ping test's details as well as a final summary of how many hosts/IPs are pingable
* Modify the Excel file so that the hosts/IPs that are reachable will be in green text, and those unreachable will be in italicized red text
