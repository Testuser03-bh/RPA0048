*** Settings ***
Library           OperatingSystem
Resource          ../Domain/VHP.robot
Resource          ../Domain/VPP.robot

*** Variables ***

${BASE_PATH}     C:/Temp/RPA0048
${filePath}      C:/UiPath/RPA0048/Processado


*** Keywords ***

Extract the data from the excel
    Phase 1 - Build Internal Memory (Step 2.1, 2.1.1)


Read Directory And Execute Company Process
    Log To Console      this is the basepath used in the companyfile: ${BASE_PATH}/Data/FinalTable
    ${dirs}=    List Directories In Directory    ${BASE_PATH}/Data/FinalTable
    ${dirs_lower}=    Evaluate    [d.lower() for d in ${dirs}]
    Clean up



    # -------- VHP PROCESS --------
    # IF    'vhpfinaltable' in ${dirs_lower}
    #     Log To Console    ===== Starting VHP Process =====
    #     Run Keyword And Continue On Failure     Connect To SAP- VHP
    #     Create the purchase order--VHP
    #     Terminate the SAP process
    #     Log To Console    ===== Finished VHP Process =====
    # END

    # -------- VPP PROCESS --------
    IF    'vppfinaltable' in ${dirs_lower}
        Log To Console    ===== Starting VPP Process =====
        Run Keyword And Continue On Failure    Connect To SAP-VPP
        Create the purchase order--VPP
        Terminate the SAP process 
        Log To Console    ===== Finished VPP Process =====
    END

    # -------- DELETE DATA FOLDER AFTER BOTH PROCESSES --------
     Clean up
    ${data_path}=    Set Variable    ${BASE_PATH}/Data
    ${folder_exists}=    Run Keyword And Return Status    Directory Should Exist    ${data_path}
    IF    ${folder_exists}
        Remove Directory    ${data_path}    recursive=True
        Log To Console    ===== Data folder deleted: ${data_path} =====
    END



    

