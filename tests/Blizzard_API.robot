*** Settings ***

Documentation    Tests for the *BlizzardConnector* class

Library  lib/BlizzardTestHelper.py
#Library  Collections

Default Tags  Blizzard API

Suite Setup    Should Not Be Empty  ${BNET API KEY}
#Test Teardown
Test Setup     Init


*** Variables ***

${A_POWER_UNBOUND}            ${11609}
${A_CHAMPIONS_OF_LEGIONFALL}  ${11846}
${A_BREACHING_THE_TOMB}       ${11546}
${A_LEGENDARY_RESEARCH}       ${11223}

*** Test Cases ***
Check Level
    [Template]    Level Should Be
    archimonde       ayonis   == 11
    archimonde       oxyr     == 60
    chants-eternels  kodyx    == 55
    voljin           oxyde    >= 110

Check ILevel
    [Template]    ILevel Should Be
    archimonde       ayonis   == 5
    archimonde       oxyr     == 63
    chants-eternels  kodyx    == 45
    voljin           oxyde    >= 932

Check Legendaries Info
    [Template]    Legendaries Info Should Be
    archimonde       ayonis   NO
    archimonde       oxyr     NO
    chants-eternels  kodyx    NO
    voljin           kodyx    940
    voljin           oxyde    970+970

Achievements Load
    Load Achievements
    Should Know Achievement  ${A_POWER_UNBOUND}
    Should Know Achievement  ${A_CHAMPIONS_OF_LEGIONFALL}
    Should Know Achievement  ${A_BREACHING_THE_TOMB}
    Should Know Achievement  ${A_LEGENDARY_RESEARCH}

*** Keywords ***

Init
    Init Test  ${BNET API KEY}

Level Should Be
    [Arguments]  ${server}  ${character name}  ${expected level ref}
    ${level} =  Get Level  ${server}  ${character name}
    Should Be True  ${level} ${expected level ref}

ILevel Should Be
    [Arguments]  ${server}  ${character name}  ${expected ilevel ref}
    ${ilevel} =  Get ILevel  ${server}  ${character name}
    Should Be True  ${ilevel} ${expected ilevel ref}

Legendaries Info Should Be
    [Arguments]  ${server}  ${character name}  ${expected leg info}
    ${leg info} =  Get Legendaries Info  ${server}  ${character name}
    Should Be Equal  ${leg info}  ${expected leg info}