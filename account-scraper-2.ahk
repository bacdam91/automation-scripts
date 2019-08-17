ms := 0

#Persistent
addSeconds()
{
    global ms := global ms + 1000
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

run()
{
    outputFolder := getOutputFolderDestination()
    outputFolderConfirmation := confirmOutputFolder(outputFolder)
    files := selectFiles()

    if (!outputFolderConfirmation or outputFolder = "" or !files)
    {
        MsgBox, ,Errors, Process is now terminating.
        ExitApp
    }
    
    readFiles(files, outputFolder)    
    MsgBox, ,Process Complete, Tasks finished. `n`nThank you for choosing BTD Technology.
    ExitApp
}

readFiles(files, destination)
{
    parentDir := ""
    Loop, parse, files, `n
    {
        if(A_Index = 1)
        {
            parentDir := A_LoopField
        }
        else 
        {
            fileReadConfirmation := confirmFileRead(A_LoopField)
            ; facilityAddress := getFacilityAddress()

            if(!fileReadConfirmation)
            {
                continue
            }
            else 
            {
                SetTimer, addSeconds, 1000
                activateNotepad()
                writeHeaders(A_LoopField)
                readPatientList(parentDir, A_LoopField, facilityAddress)
                activateNotepad()
                csvFilename := StrReplace(A_LoopField, ".txt", " Accounts.csv")
                saveFile(csvFilename, destination)
                openNewNotepad()
            }
        }
    }
}

getOutputFolderDestination()
{
    MsgBox, 0, Information, Select a destination folder for account output
    FileSelectFolder, outputFolder 
    return outputFolder
}

confirmOutputFolder(outputFolder)
{
    MsgBox, 4, Information, Confirming %outputFolder% as the destination folder. Continue?

    IfMsgBox, no
    {
        return false
    }
    else 
    {
        return true
    }
}

selectFiles() 
{
    FileSelectFile, files, M3, , Select files with lists of patient names, Text Document (*.txt)
    if (files = "")
    {
        return false
    }
    return files
}

confirmFileRead(filename)
{
    MsgBox, 4, Reading patient list, The file %filename% is now being read. Continue?
    IfMsgBox, No
    {
        return false
    }
    else 
    {
        return true
    }
}

readPatientList(parentDir, file, address) 
{
    numberOfPatients := 0
    Loop, read, %parentDir%\%file%
    {
        activateFredOffice()
        searchForPatient(A_LoopReadLine)
        
        itemIndex := 0
        ; opened := openPatientProfile(itemIndex)
        While, openPatientProfile(itemIndex)  
        {
            patientName := ""
            accNumber := ""
            closedDate := ""
            accBal := ""
            moneyVal := ""
            
            if(WinExist("Charge Account"))
            {
                patientName := getPatientName()
                accNumber := getAccountNumber()
                closedDate := getAccountClosedDate()
                accBal := getAccountBalance()
            
                moneyVal := ""
                moneyVal := convertBalanceStringToFloat(accBal)
                amount30Days := 0
                amount60Days := 0
                amount90Days := 0
                
                if(accBal != "") 
                {
                    clickOnStatements()
                    amount30Days := get30DaysAmount()
                    amount60Days := get60DaysAmount()
                    amount90Days := get90DaysAmount()
                }

                closeAccountWindow()
            }

            ; if (moneyVal <= 0)
            ; {
            ;     Sleep, 1000
            ;     activateFredDispense()
            ;     printPRF(A_LoopReadLine, address)
            ; }

            activateNotepad()

            fullname := patientName
            if(patientName = "")
            {
                fullname := A_LoopReadLine
            }

            writeToNotepad(fullname, accBal, moneyVal, amount30Days, amount60Days, amount90Days, accNumber, closedDate, A_LoopReadLine) 
            

            itemIndex := itemIndex + 1
            activateFredOffice()
        }
       
        numberOfPatients := A_Index
        
    }
    onFileReadFinish(numberOfPatients) 
}

onFileReadFinish(numberOfPatients)
{
    timeStr := createTimeStr()
    MsgBox, 0, Task Completed, Number of patients:%numberOfPatients%`nExecution time: %timeStr%.
    global ms = 0
    Return
}

get30DaysAmount()
{
    ; This click has no function, just acts as a visual cue
    Click, left, 640, 230, 2
    ControlGetText, balance, WindowsForms10.EDIT.app.0.12ab327_r6_ad119, Charge Account
    return balance
}

get60DaysAmount()
{
    ; This click has no function, just acts as a visual cue
    Click, left, 640, 260, 2
    ControlGetText, balance, WindowsForms10.EDIT.app.0.12ab327_r6_ad118, Charge Account
    return balance
}

get90DaysAmount()
{
    ; This click has no function, just acts as a visual cue
    Click, left, 640, 285, 2
    ControlGetText, balance, WindowsForms10.EDIT.app.0.12ab327_r6_ad117, Charge Account
    return balance
}

clickOnStatements()
{
    CoordMode, Mouse, Client
    Click, left, 100, 500, 1
}

activateFredOffice()
{
    IfWinExist, Fred Office for Pharmacy (soserver)
    {
        WinActivate
        CoordMode, Mouse, Client
    }
}

searchForPatient(name)
{
    ;Click clear button to cleaer search field
    Click, left, 800, 60
    ;Click on search bar
    Click, left, 510, 60
    Send, %name%
    Send, {Enter}
    Sleep, 2000
}

openPatientProfile(itemIndex)
{
    y := 130 + (20 * itemIndex)

    ; MsgBox, %y%
    Click, left, 500, %y%, 2
    Sleep, 1000
    if(WinExist("Charge Account") or itemIndex = 0)
    {
        return true
    }
    else 
    {
        return false
    }
}

getAccountNumber()
{
    ; This click has no function, just acts as a visual cue
    Click, left, 350, 95, 2
    ControlGetText, accNumber, WindowsForms10.EDIT.app.0.12ab327_r6_ad114, Charge Account
    return accNumber   
    
}

getPatientName()
{
    ; This click has no function, just acts as a visual cue
    Click, left, 365, 200, 2
    ControlGetText, name, WindowsForms10.EDIT.app.0.12ab327_r6_ad110, Charge Account
    return name   
}

getAccountClosedDate()
{
    ; This click has no function, just acts as a visual cue
    Click, left, 680, 250, 2
    ControlGetText, closedDate, WindowsForms10.EDIT.app.0.12ab327_r6_ad16, Charge Account
    return closedDate   
}


getAccountBalance()
{
    ; This click has no function, just acts as a visual cue
    Click, left, 820, 120, 2
    ControlGetText, balance, WindowsForms10.EDIT.app.0.12ab327_r6_ad12, Charge Account
    return balance
}

closeAccountWindow()
{
    Click, left, 865, -12
}

convertBalanceStringToFloat(balance)
{
    return StrReplace(balance, "$", "")
}

activateNotepad()
{
    IfWinExist, Untitled - Notepad 
    {
        WinActivate
    }
}

writeToNotepad(name, accBal, moneyVal, amount30Days, amount60Days, amount90Days, accNumber, closedDate, originalInput)
{
    accBalStr := StrReplace(accBal, ",", "")
    amount30DaysStr := StrReplace(amount30Days, ",", "")
    amount60DaysStr := StrReplace(amount60Days, ",", "")
    amount90DaysStr := StrReplace(amount90Days, ",", "")

    accStatus := "Active"
    if(closedDate != "Still Active")
    {
        accStatus := "Inactive"
    }

    Send, %name%`,%accNumber%`,%accStatus%`,%accBalStr%`,
    if (moneyVal = 0)
    {
        Send, Paid,
    }
    else if (moneyVal > 0)
    {
        Send, To be paid,
    }
    else if (moneyVal < 0 and accBal != "")
    {
        Send, Account in credit,
    }
    else if (accBal = "")
    {
        Send, Account not found,
    }
    Send, %amount30DaysStr%`,%amount60DaysStr%`,%amount90DaysStr%`,%originalInput%
    Send, {Enter}
}

saveFile(filename, destination)
{
    Sleep, 500
    Send, ^s
    Sleep, 1000

    ;Type out filename
    Send, %filename%
    Sleep, 1000

    ;Click on directory bar
    ; Click, left, 380, 15, 1
    Send, {F4}
    Sleep, 200
    Send, {Esc}
    Send, %destination%
    Send, {Enter}
    Sleep, 1000

    ; Click save button
    ; Click, left, 500, 445, 1
    Send, !s
    Sleep, 1000
    ; If Confirm Save As dialog appear
    if WinExist("Confirm Save As")
    {
        Click, left, 230, 80, 1
        Sleep, 500
    }
}

writeHeaders(filename)
{
    Send, %filename%{Enter}{Enter}
    Send, Total`,=SUM(D4:D1000){Enter}
    Send, Patient Name`,Acc No.`,Status`,Amount`,Acc Status,30 Days,60 Days,90 Days,Input{Enter}
}

openNewNotepad()
{
    Send, ^n
}

run()