*** Settings ***

Documentation    Tests for the *SheetConnector* class

Library  lib/GoogleSheetsTestHelper.py
Library  Collections

Default Tags  Google Sheets

Suite Setup       Should Not Be Empty  ${SPREADSHEET ID}  # ID of the document in Google Sheets
Test Teardown     Remove Extra Sheets

*** Variables ***

${DRYRUN}  ${TRUE}

# Sheet names
${SHEET1}               Sheet1
${INVALID SHEETNAME}    InvalidSheetName
${SHEET CR1}            Created1
${SHEET CR2}            Created2

${RANGE1}    Sheet1!b2:c3
${RANGE2}    Sheet1!b5:c5


*** Test Cases ***

Check Existing Sheet
    Connect  ${SPREADSHEET ID}  ${DRYRUN}
    Sheet Should Exist  ${SHEET1}

Check Invalid Sheet
    connect  ${SPREADSHEET ID}  ${DRYRUN}
    Sheet Should Not Exist  ${INVALID SHEETNAME}

Check Get Values
    Connect  ${SPREADSHEET ID}  ${DRYRUN}
    ${RES_LIST} =  Get Values  ${RANGE1}
    ${L1} =  Create List
    ${L2} =  Create List
    ${REF_LIST} =  Create List
    Append To List  ${L1}  b2  c2
    Append To List  ${L2}  b3  c3
    Append To List  ${REF_LIST}  ${L1}  ${L2}
    Lists Should Be Equal  ${REF_LIST}  ${RES_LIST}

Update Values
    Connect  ${SPREADSHEET ID}
    # ${ORIG_LIST} =  Get Values  ${RANGE2}
    ${L1} =  Create List
    ${L2} =  Create List
    Append To List  ${L1}  b5  c5
    Append To List  ${L2}  ${L1}
    Update Values  ${RANGE2}  ${L2}
    ${RES_LIST} =  Get Values  ${RANGE2}
    Lists Should Be Equal  ${L2}  ${RES_LIST}

    ${L1} =  Create List
    ${L2} =  Create List
    Append To List  ${L1}  B.5  C.5
    Append To List  ${L2}  ${L1}
    Update Values  ${RANGE2}  ${L2}
    ${RES_LIST} =  Get Values  ${RANGE2}
    Lists Should Be Equal  ${L2}  ${RES_LIST}

    # Update Values  ${RANGE2}  ${ORIG_LIST}
    # ${RES_LIST} =  Get Values  ${RANGE2}
    # Lists Should Be Equal  ${ORIG_LIST}  ${RES_LIST}

Check Or Create Sheet
    Connect  ${SPREADSHEET ID}
    Sheet Should Not Exist  ${SHEET CR1}
    Check Or Create Sheet  ${SHEET CR1}
    Sheet Should Exist  ${SHEET CR1}

Delete Sheet
    Connect  ${SPREADSHEET ID}
    Check Or Create Sheet  ${SHEET CR1}
    Sheet Should Exist  ${SHEET CR1}
    Delete Sheet  ${SHEET CR1}
    Sheet Should Not Exist  ${SHEET CR1}

Ensure Headers
    Connect  ${SPREADSHEET ID}
    Sheet Should Not Exist  ${SHEET CR2}
    ${L1} =  Create List
    ${REF_LIST} =  Create List
    Append To List  ${L1}  h1  h2  h3
    Append To List  ${REF_LIST}  ${L1}
    Ensure Headers  ${SHEET CR2}  ${L1}
    Sheet Should Exist  ${SHEET CR2}
    ${RES_LIST} =  Get Values  ${SHEET CR2}!1:1
    Lists Should Be Equal  ${REF_LIST}  ${RES_LIST}

    ${L2} =  Create List
    ${REF_LIST} =  Create List
    Append To List  ${L2}  h2  h4  h5
    Append To List  ${L1}  h4  h5
    Append To List  ${REF_LIST}  ${L1}
    Ensure Headers  ${SHEET CR2}  ${L2}
    ${RES_LIST} =  Get Values  ${SHEET CR2}!1:1
    Lists Should Be Equal  ${REF_LIST}  ${RES_LIST}

Get Background Colors
    [Setup]       Connect  ${SPREADSHEET ID}
    [Template]    Check Background Color
    ${SHEET1}  A  ${1}  \#000000
    ${SHEET1}  B  ${1}  \#FF0000
    ${SHEET1}  C  ${1}  \#00FF00
    ${SHEET1}  D  ${1}  \#0000FF
    ${SHEET1}  E  ${1}  \#FFFFFF


Set Background Colors
    [Setup]       Connect  ${SPREADSHEET ID}
    [Template]    Set And Check Background Color
    ${SHEET1}  \#FFFFFF
    ${SHEET1}  \#000000
    ${SHEET1}  \#FF0000
    ${SHEET1}  \#00FF00
    ${SHEET1}  \#0000FF
    ${SHEET1}  \#FF7D0A
    ${SHEET1}  \#9482C9
    ${SHEET1}  \#0070DE
    ${SHEET1}  \#F58CBA
    ${SHEET1}  \#C79C6E
    ${SHEET1}  \#C41F3B
    ${SHEET1}  \#A330C9
    ${SHEET1}  \#00FF96
    ${SHEET1}  \#69CCF0
    ${SHEET1}  \#ABD473
    ${SHEET1}  \#FFF569


*** Keywords ***

Sheet Should Exist
    [Arguments]    ${sheetname}
    ${RESULT} =    Sheet Exists    ${sheetname}
    Should Be True    ${RESULT}

Sheet Should Not Exist
    [Arguments]    ${sheetname}
    ${RESULT} =    Sheet Exists    ${sheetname}
    Should Be Equal    ${RESULT}    ${FALSE}

Check Background Color
    [Arguments]    ${sheet}  ${col}  ${row}  ${ref_hex_color}
    ${CELL} =  Create List
    Append To List  ${CELL}  ${col}  ${row}
    ${HEX_COLOR} =  Get Cell Hex Color  ${sheet}  ${CELL}
    Should Be Equal  ${HEX_COLOR}  ${ref_hex_color}

Set And Check Background Color
    [Arguments]    ${sheet}  ${ref_hex_color}
    ${CELL} =  Get Random Cell Index
    Set Cell Color  ${sheet}  ${CELL}  ${ref_hex_color}
    ${HEX_COLOR} =  Get Cell Hex Color  ${sheet}  ${CELL}
    Should Be Equal  ${HEX_COLOR}  ${ref_hex_color}
