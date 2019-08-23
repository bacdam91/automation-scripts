Gui, Add, Text, ,Doctor's Code (Not precriber number)
Gui, Add, Edit, Limit4 w50 Uppercase vDoctorCode
Gui, Add, Text, ,Number of owings 
Gui, Add, Edit, Number Limit3 w50 vNoOfOwings
Gui, Add, Text, ,Number of scripts to skip 
Gui, Add, Edit, Number Limit3 w50 vOffset
Gui, Add, Text, ,Pharmacist Initial
Gui, Add, Edit, Limit2 Uppercase w50 vPharmInit
Gui, Add, Button, Default w80 gRun, Start
Gui, Show
return 

timeBetweenActions := 1000

Run:
Gui, Submit, NoHide
run()
return

run() 
{
    
    showWarningMessage()
    runChanger()    
}

runChanger()
{
    global NoOfOwings
    activateFredDispense() 
    Loop, %NoOfOwings%
    {
        changeDoctor(A_Index)
    }
}

changeDoctor(index)
{
    openOwingScriptPanel()
    closeAllPopUpsTwice()
    editScript(index)
    closeAllPopUpsTwice()
    changeDr()
    closeAllPopUpsTwice()
    saveChanges()
    closeAllPopUpsTwice()
}

changeDr()
{
    global DoctorCode
    global timeBetweenActions
    Send, {Down}
    Send, %DoctorCode%
    Send, {Enter}
    Sleep, timeBetweenActions
    terminateProgram()
    Send, {End}
    Sleep, 4000
    ; Send, {Esc}
    ; Send, {End}
}

terminateProgram()
{
    if(WinExist("Doctor Selection"))
    {
        MsgBox, Doctor is NOT found. This process will now terminate.
        IfMsgBox, OK
            ExitApp
            Return
    }
}

openOwingScriptPanel()
{
    global timeBetweenActions
    Send, !{F3}
    Sleep, timeBetweenActions
}

activateFredDispense()
{
    if(WinExist("Fred Dispense"))
    {
        WinActivate, Fred Dispense
    }
}

showWarningMessage()
{
    MsgBox, ,Date of Birth, Please ensure date of birth is entered.
}

closeAuthorityOwingItemBox()
{
    global timeBetweenActions
    if(WinExist("Authority Owing Item"))
    {
        Send, {Enter}
        Sleep, timeBetweenActions
        Send, {Esc}
        Sleep, timeBetweenActions
        Send, {End}
    }
}

closeEntitlementDetailsChangedBox()
{
    global timeBetweenActions
    if(WinExist("Entitlement Details Changed"))
    {
        Send, n
        Sleep, timeBetweenActions
    }
}

closeScriptDateBox()
{
    if(WinExist("Script Date"))
    {
        Send, y
        Sleep, 4000
    }
}

closeEditExistingPatientDetailsBox()
{
    global timeBetweenActions
    if(WinExist("Edit Existing Patient Details"))
    {
        Send, {Esc}
        Sleep, timeBetweenActions
    }
}

closeAllPopUps()
{
    closeAuthorityOwingItemBox()
    closeEditExistingPatientDetailsBox()
    closeEntitlementDetailsChangedBox()
    closeScriptDateBox()
}

closeAllPopUpsTwice()
{
    closeAllPopUps()
    closeAllPopUps()
}

editScript(index)
{
    global DoctorCode
    global NoOfOwings
    global Offset
    global PharmInit
    global timeBetweenActions

    ; MsgBox, Offset %offset%
    noOfDownPresses := index - 1 + Offset
    Loop, %noOfDownPresses%
    {
        Send, {Down}
    }
    Send, {F4}
    Sleep, 600
    Send, e
    Sleep, timeBetweenActions
}

saveChanges()
{
    global PharmInit
    Send, %PharmInit%
    Sleep, 1000
}