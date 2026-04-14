*** Settings ***
Library            SapGuiLibrary    screenshots_on_error=False
Library            Process
Library            OperatingSystem
Library            Collections
Resource           ExcelFiles.robot
Library            RPA.Desktop
Library            String
Library            DateTime
Library            RPA.Email.ImapSmtp
Library            RPA.Excel.Application
Library            ../adapters/Library/InitAllSettingsSQL.py
Resource           ${EXECDIR}/Data/Resource/VHPlocators.resource


   
*** Keywords ***

Connect To SAP- VHP
   
    ${primary_fetched_config}=    Get All Settings    ${PRIMARY_PROCESS_NAME}
    ${secondary_fetched_config}=    Get All Settings    ${SECONDARY_PROCESS_NAME}
    Set Global Variable    ${primary_config}    ${primary_fetched_config}
    Set Global Variable    ${secondary_config}    ${secondary_fetched_config}
    
    Start Process      ${SAP_LOGON_PATH}    shell=True
    Wait Until Keyword Succeeds    20s    4s    Connect To Session
    Connect To Session
    Evaluate                __import__('dotenv').load_dotenv(".env")
    ${ENV}=    Get Environment Variable    ENV
    IF    '${ENV}' == 'UAT'
        ${SID}=    Set Variable    ${primary_config['SID_Test']}
    ELSE
        ${SID}=    Set Variable    ${primary_config['SID']}
    END
    ${system_code}=       Evaluate       "${SID}".split("|")[1].strip()
    Log To Console    ${system_code}
    Open Connection    ${system_code} 
    Wait Until Keyword Succeeds    20s    4s    Element should Be Present    wnd[0]/usr/txtRSYST-BNAME
    Input Text         wnd[0]/usr/txtRSYST-BNAME    ${primary_config['SAP_User']}
    Wait Until Keyword Succeeds    30s    3s    Element should Be Present    wnd[0]/usr/pwdRSYST-BCODE
    Evaluate           __import__('dotenv').load_dotenv(".env")
    ${pass}=                Get Environment Variable    VHPPASSWORD
    Input Password          wnd[0]/usr/pwdRSYST-BCODE    ${pass}
    Send VKey          0
    ${IMG_PATH}=    Set Variable    image:${EXECDIR}//Data//Images//tick.png 
    Run keyword and ignore error  RPA.Desktop.Click  image:${IMG_PATH}tick.png 
  

Create the purchase order--VHP  
# Read data From the Header_aux_table
    ${empresa}=     Evaluate     "${primary_config['Empresa']}".split("|")[1].split(";")[0].strip()
    ${company_type}=     Evaluate     "${primary_config['Empresa']}".split("|")[1].split(";")[0].strip()
    Log To Console    ${company_type}

    ${header_data}=    Extract All Header Fields    ${DIR}       ${Empresa}
    FOR    ${item}    IN    @{header_data}

        ${fileName}=      Get From Dictionary    ${item}    file
        ${file}           Evaluate               "${fileName}"[:16]
        ${headers}=       Get From Dictionary    ${item}    headers
        ${titulo}=        Get From Dictionary    ${headers}    Título do serviço:
        ${num}=           Get From Dictionary    ${headers}    Núm. participantes:
        ${valor}=         Get From Dictionary    ${headers}    Valor do serviço:
        ${fornecedor}=    Get From Dictionary    ${headers}    Fornecedor:

        ${ValuationPrice}=      Get From Dictionary    ${headers}    Valor do serviço:   
        ${materialGrp}=         Get From Dictionary    ${headers}    Grupo de mercadorias:
        ${Quantity}=            Get From Dictionary    ${headers}    Qtd. solicitações para o ano:   
        ${ShortText}=           Get From Dictionary    ${headers}    Título do serviço:
        ${ShortText}            Evaluate                "${ShortText}"[:40]
        ${tipo_servico}=        Get From Dictionary    ${headers}    Título do serviço:
        ${final_text}=    Catenate    SEPARATOR=\n
        ...    Título do serviço: ${titulo}
        ...    Núm. participantes: ${num}
        ...    Valor do serviço: ${valor}
        ...    Fornecedor: ${fornecedor}

        
        Wait Until Keyword Succeeds    30s    2s    Element should Be Present    wnd[0]/tbar[0]/okcd
        Run Transaction    ${primary_config['SAP_Operation']}      

        #Step5.5.2.2.4
        Press Keys     ctrl    f2   
        sleep    1s
        Select From List By Label     ${DROPDOWN_IDD}    Purchase requisition
        Wait Until Keyword Succeeds    30s    2s    Element should Be Present    ${TEXT_EDITOR_IDD}


        
        Input Text    ${TEXT_EDITOR_IDD}    ${final_text}

        ${today}=    Get Current Date    result_format=%Y-%m-%d
        ${delivery_date}=    Add Time To Date     ${today}     ${primary_config['SAP_DeliveryDate']} days      result_format=%d.%m.%Y
        ${SAP_Plant}=             Extract Hydro Plant     ${primary_config['SAP_Plant']}       ${company_type}           
        ${SAP_Storage}=         Evaluate               "${primary_config['SAP_Storage']}".split("|")[1].strip()
        ${SAP_PurchaseGroup}=   Evaluate    "${primary_config['SAP_PurchaseGroup']}".split("|")[1].strip()
        Set Cell Value    ${GRID_IDD}    0    KNTTP    ${primary_config['SAP_Category']}  
        Set Cell Value    ${GRID_IDD}    0    TXZ01    ${ShortText}
        Set Cell Value    ${GRID_IDD}    0    MENGE    ${Quantity}
        Set Cell Value    ${GRID_IDD}    0    MEINS    ${primary_config['SAP_UM']}
        Set Cell Value    ${GRID_IDD}    0    WGBEZ    ${materialGrp}
        Set Cell Value    ${GRID_IDD}    0    EEIND    ${delivery_date}
        Set Cell Value    ${GRID_IDD}    0    NAME1    ${SAP_Plant}
        Wait Until Keyword Succeeds    30s    2s    Element should Be Present    wnd[0]/sbar 
        ${msg_type}=    Get Value    wnd[0]/sbar    
        # Step 5.3.2.2.10 — Invalid Plant Error
        IF    "${msg_type}" == "E"
            Create Invalid Plant Error
            ...    ${empresa}
            ...    ${tipo_servico}
            ...    ${ROOT_DIR}
            Continue For Loop
        END

        Set Cell Value    ${GRID_IDD}    0    EKGRP    ${SAP_PurchaseGroup}
        Set Cell Value    ${GRID_IDD}    0    AFNAM    ${primary_config['SAP_User']}
        Set Cell Value    ${GRID_IDD}    0    LGOBE    ${SAP_Storage}
        Wait Until Keyword Succeeds    30s    2s    Element should Be Present    wnd[0]/sbar 
        # Step 5.3.2.2.11 — Invalid Storage Error

                ${msg_type}=    Get Value    wnd[0]/sbar
                IF    "${msg_type}" == "E"
                    Create Invalid Storage Error
                    ...    ${empresa}
                    ...    ${tipo_servico}
                    ...    ${ROOT_DIR}
                    Continue For Loop
                END

        Wait For Element    image:${EXECDIR}//Data//Images//Desired_vendor.png    timeout=20
        Move Mouse          image:${EXECDIR}//Data//Images//Desired_vendor.png
        Sleep    0.5s
        Click               image:${EXECDIR}//Data//Images//Desired_vendor.png 
        Press Keys     down
        Wait For Element    image:${EXECDIR}//Data//Images//Inside_celll.png    timeout=20
        Move Mouse          image:${EXECDIR}//Data//Images//Inside_celll.png 
        Sleep    0.5s
        Click                image:${EXECDIR}//Data//Images//Inside_celll.png
        Sleep  1s
        Input Text    ${TaxNo_onee}    ${VENDOR_ID}
        Wait Until Keyword Succeeds    20s    1s    Element should Be Present    wnd[1]/tbar[0]/btn[0]
          # Step 5.3.2.2.15 — CNPJ Not Found Error
        # ============================================================

                ${sap_message}=    Get Value    wnd[0]/sbar
                IF    "${msg_type}" == "E"
                    Create Cnpj Error
                    ...    ${empresa}
                    ...    ${tipo_servico}
                    ...    ${sap_message}
                    ...    ${ROOT_DIR}
                    Send Vkey    12
                    Sleep        1s
                    Send Vkey    12
                    Continue For Loop
                END

        Click Element        wnd[1]/tbar[0]/btn[0]
        Wait Until Keyword Succeeds    20s    1s    Element should Be Present    wnd[1]/tbar[0]/btn[0]   
        Click Element    wnd[1]/tbar[0]/btn[0]
       
        Send Vkey    0
        Wait Until Keyword Succeeds    30s    2s    Element should Be Present    wnd[0]/sbar
        # Wait until eleemnt is visible  wnd[0]/sbar  timeout=10

      

        Wait Until Keyword Succeeds    20s    2s    Element should Be Present    ${CustomerDataa} 
        Click Element      ${CustomerDataa}
        Set Focus          ${ContactPersonn}
        Input Text        ${ContactPersonn}      ${primary_config['SAP_User']} 
        Wait Until Keyword Succeeds    30s    2s    Element should Be Present    wnd[0]/sbar

        # Step 5.3.2.2.16 — Vendor Blocked / Quantity / Service Approver
        # ============================================================

                ${sap_message}=    Get Value    wnd[0]/sbar
                IF    "${msg_type}" == "E"
                    Create Vendor Error
                    ...    ${empresa}
                    ...    ${tipo_servico}
                    ...    ${sap_message}
                    ...    ${ROOT_DIR}
                    Send Vkey      12
                    Sleep          1s
                    Click Element  wnd[1]/tbar[0]/btn[0]
                    Continue For Loop
                END

       #Assignment table      
        Wait Until Keyword Succeeds    20s    2s    Element should Be Present     ${AccountAssignmnett}
        Set Focus          ${AccountAssignmnett}
        Click Element      ${AccountAssignmnett}
        Log To Console     ${company_type}

        ${Final_sheetheader_data}=    Extract All Header Fields From Finaltable    ${DIR}    ${company_type}
        ${expected_title}=    Set Variable    ${file}

        FOR    ${Finalsheet_item}    IN    @{Final_sheetheader_data}

            ${full_file}=    Get From Dictionary    ${Finalsheet_item}    file
            ${filename_only}=    Evaluate    os.path.basename(r'''${full_file}''')    os

            IF    '${expected_title}' in '${filename_only}'
                
                ${EmpCount}=    Get From Dictionary    ${Finalsheet_item}    employeeCount

                IF    ${EmpCount} == 0
                    Create No Employees Error
                    ...    ${empresa}
                    ...    ${tipo_servico}
                    ...    ${ROOT_DIR}
                    Continue For Loop
                END

                Log To Console    tyhis is yhe empocount ${EmpCount}


                IF    ${EmpCount} >1
                    Sleep    0.5s
                    Select From List By Label    ${DISTRIBUTION_COMBOo}    Distribution by percentage
                ELSE
                    Sleep    0.5s
                    Select From List By Label    ${DISTRIBUTION_COMBOo}    Single account assignment
                END

                # fill the data into the assignmnet tab
                ${employees}=    Get From Dictionary    ${Finalsheet_item}    employees
                ${row}=     Set Variable     0
                
                FOR    ${emp}    IN    @{employees}

                    ${Perce}=      Get From Dictionary    ${emp}    Porcentagem
                    ${costCtr}=    Get From Dictionary    ${emp}    CDC
                    ${G_LAct}=     Get From Dictionary    ${emp}    Class. Contábil
                    ${PF_locator}=    Catenate        SEPARATOR=        ${PERCENT_FIELD}     [3,${row}]
                    Input Text    ${PF_locator}    ${Perce} 
                    ${CS_locator}=    Catenate        SEPARATOR=        ${COSTCENTER_FIELD}     [4,${row}]
                    Input Text       ${CS_locator}    ${costCtr}
                    ${GLA_locator}=    Catenate        SEPARATOR=        ${GLAcct}     [5,${row}]
                    Input Text       ${GLA_locator}              ${G_LAct}
                    ${Rc_locator}=    Catenate        SEPARATOR=        ${Recipient}     [8,${row}]
                    Input Text     ${Rc_locator}     ${primary_config['SAP_User']}
                    Set Focus       ${Rc_locator}
                     IF    ${row} < 5     
                        ${row}=    Evaluate    ${row} + 1
                    ELSE    
                        Press Keys     SHIFT       down
                        Sleep     0.1s
                    END

                    Send Vkey     0

                   
                
                    
                END
                Wait Until Keyword Succeeds    30s    2s    Element should Be Present    wnd[0]/sbar
                # Step 5.3.2.14 — Account Assignment Errors
                ${sap_message}=    Get Value    wnd[0]/sbar
                IF    "Purchasing across company codes" in "${sap_message}" or "Cost center" in "${sap_message}" or "account" in "${sap_message}" or "Sum of percentages" in "${sap_message}"
                    Create Account Assignment Error
                    ...    ${empresa}
                    ...    ${tipo_servico}
                    ...    ${sap_message}
                    ...    ${ROOT_DIR}
                    Continue For Loop
                END

                Exit For Loop

            END



        END
        
            Input Text      ${Valuation_FIELD}          ${ValuationPrice} 
            Unselect Checkbox   ${GdBoxChkBx}
            Unselect Checkbox    ${GRNonValuation}

            Wait Until Keyword Succeeds    20s    2s    Element should Be Present    ${Source_of_supply}
            Click Element      ${Source_of_supply}

            Wait Until Keyword Succeeds    20s    2s    Element should Be Present    ${PurchaseOrg}
            ${Porg}=    Evaluate     "${primary_config['SAP_SalesOrg']}".split("|")[1].strip()
            Input Text    ${PurchaseOrg}     ${Porg}  

            #Click on the save button
            Click Element       wnd[0]/tbar[0]/btn[11]
            Wait Until Keyword Succeeds    10s    2s    Element should Be Present        wnd[0]/sbar 
            ${PR_status}=    Get value    wnd[0]/sbar
            Log To Console     ${PR_status}
            ${sap_status}=    Get Value    wnd[0]/sbar

            # Step 5.3.2.19 / 5.3.2.23 — PR Created Successfully
            Create Pr Success Record
            ...    ${empresa}
            ...    ${tipo_servico}
            ...    ${sap_status}
            ...    ${ROOT_DIR}  
            log to console      works till here  
            ${pr_number}=    Update Excel And Report
            ...    ${fileName}
            ...    ${sap_status}
            ...    ${empresa}
            ...    ${tipo_servico}
            ...    ${file_path}
            ...    ${ROOT_DIR}

            Log To Console    PR Created: ${pr_number}

            Wait Until Keyword Succeeds    20s    2s    Element Should Be Present     wnd[0]/tbar[1]/btn[17]
            Click Element      wnd[0]/tbar[1]/btn[17]

            Wait Until Keyword Succeeds    20s    2s    Element Should Be Present     wnd[1]/tbar[0]/btn[0]
            Click Element      wnd[1]/tbar[0]/btn[0]
         
            # Step 1 - Click Service Object / GOS Toolbox
            Wait Until Keyword Succeeds    20s    2s    Element Should Be Present    wnd[0]/titl/shellcont/shell
            Click Toolbar Button           wnd[0]/titl/shellcont/shell    %GOS_TOOLBOX

            # Step 2 - Click Create → Create Attachment
            Wait Until Keyword Succeeds    20s    2s    Element Should Be Present    wnd[0]/titl/shellcont/shell
            Select Context Menu Item       wnd[0]/titl/shellcont/shell    %GOS_TOOLBOX    %GOS_PCATTA_CREA
            ${filePath}=    Normalize Path    ${file_path} 
            Log To Console       ${filePath}  

            Input Text      wnd[1]/usr/ctxtDY_PATH          ${filePath}     
            Input Text      wnd[1]/usr/ctxtDY_FILENAME      ${fileName}
            ${log_path}=    Set Variable   ${filePath}${/}${fileName}
            Log To Console        ${log_path}

            Wait Until Keyword Succeeds    20s    2s    Element Should Be Present    wnd[1]/tbar[0]/btn[0]

            Click Element      wnd[1]/tbar[0]/btn[0]

            ${sap_final}=    Get Value    wnd[0]/sbar
            Log To Console       ${sap_final}
            log to console       after the sap final variables

            # Step 5.3.2.31 / 5.3.2.32 — Attachment Success or Failure
            Create Attachment Record
            ...    ${empresa}
            ...    ${tipo_servico}
            ...    ${sap_final}
            ...    ${ROOT_DIR}

            IF    "successfully" not in $sap_final
                Continue For Loop
            END

            IF    "The attachment was successfully created" in $sap_final

                ${only_pr}=    Get From Dictionary    ${pr_number}    pr_number
                Send Email Final Report-prnumber  ${only_pr}  ${tipo_servico} 
                log to console     "Email is send succcesfully for the Pr number "
            END


            Handle Attachment Status     ${ROOT_DIR}  ${empresa}    ${tipo_servico}     ${sap_final}

            ${AttachmentFilePath}=     Set Variable     ${ROOT_DIR}${/}Report${/}Relatório Final_Analítico.xlsx
            Send Email Final Report   ${AttachmentFilePath}
   
            sleep     2s

            Press Keys        alt      f4
            sleep       1s

    END

Send Email Final Report-prnumber   [Arguments]     ${only_pr}     ${tipo_servico}
        log to console      this ins inside the email final report
        ${grupo}=    Convert To String    ${materialGrp}
        ${grupo}=    Strip String    ${grupo}

        IF    $grupo != "" and $grupo != "203801"
            ${SEmail}=    Set Variable    Udit.Kumar-extern@voith.com
            # ${SEmail}=    Set Variable    ${primary_config['Email_Report_Outros']}
            Log To Console    Using OUTROS email: ${SEmail}
        ELSE
            ${SEmail}=    Set Variable    Udit.Kumar-extern@voith.com
            # ${SEmail}=    Set Variable    ${primary_config['Email_Report_Treinamento']}
            Log To Console    Using TREINAMENTO email: ${SEmail}
        END
    # 2. Map Dynamic DB Parameters for the Email Configuration
    ${email_sender}=       Set Variable    ${primary_config['Email_Sender']}
    ${email_recipient}=    Set Variable    ${SEmail}
    ${email_subject}=      Set Variable    ${primary_config['Email_Subject']} - RPA0048
    
    # 3. Handle HTML Body Content
    # (Assuming Email_Body_FinalProcess contains the direct HTML string. 
    # If it's a file path instead, you would use: Get File ${primary_config['Email_Body_FinalProcess']})
    #${body_msg}=           Set Variable    ${primary_config['Email_Body_FinalProcess']}

    # 4. Authorize SMTP Connection
    Authorize SMTP    
    ...    account=${primary_config['Email_Sender']}  
    ...    password=${EMPTY}    
    ...    smtp_server=${secondary_config['SMTP_Server']}   
    ...    smtp_port=${secondary_config['SMTP_Port']}
    ${body_msg}    Catenate    SEPARATOR=\n
    ...    Olá,
    ...    A Requisição nº ${only_pr} foi aberta, por gentileza aprovar.  
    ...    Best Regards,
    ...    Voith RPA

    Log To Console    Preparing to send final report to ${email_recipient}...

    # 5. Send Message using the Mapped Parameters and Attach Log File
    Send Message    
    ...    sender= ${primary_config['Email_Sender']}  
    ...    recipients= ${SEmail}
    ...    subject= ${primary_config['Email_Subject']}
    ...    body= ${body_msg}
    # ...    html= True
    # ...    attachments=${log_path}


     # Create description
    ${descricao}=    Catenate    SEPARATOR=    E-mail foi enviado para:     ${SEmail}

    # Call helper
    Create Error Record
    ...    Enviar e-mail
    ...    ${descricao}
    ...    ${empresa}
    ...    ${tipo_servico}
    ...    ${DIR}
    ...    email_sent_memory.json

    Log To Console    📧 Reportfor the PR number successfully sent with attached log!


Send Email Final Report   [Arguments]   ${AttachmentFilePath}     
        ${grupo}=    Convert To String    ${materialGrp}
        ${grupo}=    Strip String    ${grupo}

        IF    $grupo != "" and $grupo != "203801"
            ${SEmail}=    Set Variable    Udit.Kumar-extern@voith.com
            # ${SEmail}=    Set Variable    ${primary_config['Email_Report_Outros']}
            Log To Console    📧 Using OUTROS email: ${SEmail}
        ELSE
            ${SEmail}=    Set Variable    Udit.Kumar-extern@voith.com
            # ${SEmail}=    Set Variable    ${primary_config['Email_Report_Treinamento']}
            Log To Console    📧 Using TREINAMENTO email: ${SEmail}
        END
    # 2. Map Dynamic DB Parameters for the Email Configuration
    ${email_sender}=       Set Variable    ${primary_config['Email_Sender']}
    ${email_recipient}=    Set Variable    ${SEmail}
    ${email_subject}=      Set Variable    ${primary_config['Email_Subject']} - RPA0048
    
    # 3. Handle HTML Body Content
    # (Assuming Email_Body_FinalProcess contains the direct HTML string. 
    # If it's a file path instead, you would use: Get File ${primary_config['Email_Body_FinalProcess']})
    #${body_msg}=           Set Variable    ${primary_config['Email_Body_FinalProcess']}

    # 4. Authorize SMTP Connection
    Authorize SMTP    
    ...    account=${primary_config['Email_Sender']}  
    ...    password=${EMPTY}    
    ...    smtp_server=${secondary_config['SMTP_Server']}   
    ...    smtp_port=${secondary_config['SMTP_Port']}

    Log To Console    Preparing to send final report to ${email_recipient}...
    
    # 5. Send Message using the Mapped Parameters and Attach Log File
    Send Message    
    ...    sender= ${primary_config['Email_Sender']}  
    ...    recipients= ${SEmail}
    ...    subject= ${primary_config['Email_Subject']}
    ...    body= ${Final_EMAIL_BODY}
    # ...    html= True
    ...    attachments=${AttachmentFilePath} 

    Log To Console    📧 Report successfully sent with attached log!

Clean up
    [Documentation]    Closes Excel and kills all SAP processes to ensure a clean state.
    Log To Console    Cleaning up environment...
    
    # 2. End SAP
    Run Keyword And Ignore Error    OperatingSystem.Run    taskkill /F /IM saplogon.exe /T
    Run Keyword And Ignore Error    OperatingSystem.Run    taskkill /F /IM sapgui.exe /T

Terminate the SAP process

    Press Keys     ALT  F4
    sleep    1s
    press keys     ctrl    left
    sleep    1s
    press keys      enter 
 