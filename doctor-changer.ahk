ms := 0
offset := 0
timeBetweenActions := 1000

#Persistent
addSeconds()
{
    global ms := global ms + 1000
}

run() {
    drCode := getDoctorsCode()
    noOfOwings := getNoOfOwings()
    pharmacistCode := getPharmacistCode()
    global offset := getScriptOffset()
    result := confirmDetails(drCode, noOfOwings, pharmacistCode)
    if(result = true)
    {
        runChangeDoctor(drCode, pharmacistCode, noOfOwings)
    } 
    else 
    {
        MsgBox, 0, Error, Script terminated.
    }
    
}

getScriptOffset()
{
    InputBox, offset, No of scripts to skip, Enter the number of scripts to skip
    return offset
}

runChangeDoctor(drCode, pharmacistCode, noOfOwings)
{
    SetTimer, addSeconds, 1000
    activateFredDispense()
    Loop, %noOfOwings%
    {
        changeDoctor(drCode, A_Index, pharmacistCode)
    }
    onScriptFinish(drCode, noOfOwings)
}

onScriptFinish(drCode, noOfOwings)
{
    timeStr := createTimeStr()
    MsgBox, 0, Task Completed, Number of owing scripts changed: %noOfOwings%`nNew doctor code: %drCode%`nExecution time: %timeStr%. `n`nThank you for using BTD Technology.
    ExitApp
    Return
}

createTimeStr()
{
    global ms 
    seconds := ms / 1000
    minutes := seconds / 60
    remainingSeconds := Mod(seconds, 60)
    timeStr := Floor(minutes) . " minutes and " . Floor(remainingSeconds) . " seconds"
    return timeStr
}

confirmDetails(drCode, noOfOwings, pharmacistCode)
{
    global offset
    MsgBox, 4, Confirm details, Doctor's code: %drCode%`n`nNumber of owings: %noOfOwings%`n`nPharmacist Code: %pharmacistCode% `n`nStarting script position: %offset%`n`nEnsure DOB is entered!
    IfMsgBox, Yes
        Return True
    else
        Return False
}

activateFredDispense()
{
    if(WinExist("Fred Dispense"))
    {
        WinActivate, Fred Dispense
    }
}

getDoctorsCode() 
{
    InputBox, drCode, Doctor's Code, Enter Doctor's code
    return drCode
}

getPharmacistCode()
{
    InputBox, pharmacistCode, Pharmacist's Code, Enter Pharmacist's code
    return pharmacistCode
}

getNoOfOwings() 
{
    InputBox, noOfOwings, Number of owings, Enter the number of owings
    return noOfOwings
}

openOwingScriptPanel()
{
    global timeBetweenActions
    Send, !{F3}
    Sleep, timeBetweenActions
}

editScript(index)
{
    global timeBetweenActions
    global offset
    ; MsgBox, Offset %offset%
    noOfDownPresses := index - 1 + offset
    Loop, %noOfDownPresses%
    {
        Send, {Down}
    }
    Send, {F4}
    Sleep, 600
    Send, e
    Sleep, timeBetweenActions
}

changeDr(drCode)
{
    global timeBetweenActions
    Send, {Down}
    Send, %drCode%
    Send, {Enter}
    Sleep, timeBetweenActions
    terminateProgram()
    Send, {End}
    Sleep, 4000
    ; Send, {Esc}
    ; Send, {End}
}

saveChanges(pharmacistCode)
{
    Send, %pharmacistCode%
    Sleep, 1000
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
    ; waitForDOB()
}

closeAllPopUpsTwice()
{
    closeAllPopUps()
    closeAllPopUps()
}

waitForDOB()
{
    if(WinExist("  Monitored Vic S4 Drug "))
    {
        MsgBox, Unable to continue. Please fix pop up and press ok to continue.
    }
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


changeDoctor(drCode, index, pharmacistCode)
{
    openOwingScriptPanel()
    closeAllPopUpsTwice()
    editScript(index)
    closeAllPopUpsTwice()
    changeDr(drCode)
    closeAllPopUpsTwice()
    saveChanges(pharmacistCode)
    closeAllPopUpsTwice()

}

run()