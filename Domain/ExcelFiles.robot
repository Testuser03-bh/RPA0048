*** Settings ***
Library    ../adapters/helperFileExcel.py
Library    Process
Library    OperatingSystem

*** Variables ***
${CMD_EXECUTABLE}  cmd.exe

${DIR}             C:/Temp/RPA0048
${FileDir}         C:/UiPath/RPA0048

*** Keywords ***

Phase 1 - Build Internal Memory (Step 2.1, 2.1.1)
    ${memory}=    Build Internal Memory     ${FileDir}     ${DIR}
    Log To Console    Internal memory created and JSON updated

    Log To Console  phase 2 - Generate Sorted JSON(step 4)
   ${sorted_data}=    Generate Sorted Json Common   ${DIR}
   Log To Console    sort sheet is created 

    Log To Console     Phase 3 - Extract Filenames From Sorted JSON
    ${files}=    Extract Filenames Per Sorted Json    ${DIR}
    
    ${data}=    Extract Company Costs Filewise    ${DIR}

    Log To Console      Phase 5 - Company Header (Step 5.1.2)
    ${aux_table}=    Build Header Aux Table Filewise   ${DIR}
    Log To Console    Header auxiliary JSON created




    Log To Console      Phase 6 - read the column E4(5.1.3)
    ${currency}=    Extract Currency Filewise   ${DIR}   

    Log To Console  phase 7 - Extract the employyes (Steps - 5.1.4)
    ${Employees}=    Extract Employees Filewise  ${DIR}


    Log To Console     Phase 5.3.1 - Build Final Table (EMPRESA == COUNTER_STEP)
    ${final_table}=    Build Final Table Filewise   ${DIR}



   

# phase 4 - Open each file from the path
#     FOR    ${file_path}     IN    @{files} 
#         Log To Console    OpeningFile : ${file_path}
#         File Should Exist    ${file_path}
#         Run Process    explorer.exe    ${file_path}
#     END
    
