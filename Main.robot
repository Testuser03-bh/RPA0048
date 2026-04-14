*** Settings ***
Library           OperatingSystem
Resource          ../RPA0048/Flows/Flows.robot
Library           ../RPA0048/adapters/Library/RobotProcessLibrary.py



*** Tasks ***

robot file 
    Initialize Robot Process
    Extract the data from the excel
    Read Directory And Execute Company Process
    End Robot Process

