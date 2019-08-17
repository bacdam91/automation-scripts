run() 

run() 
{
    files := selectFiles()
    if (!files)
    {
        MsgBox, 0, Information, No file was selected, 10
        return
    }
    parseFiles(files)
}

activateFredDispense()
{
    if WinExist("Fred Dispense")
    {
        WinActivate
        CoordMode, Mouse, Client
    }
}

selectFiles()
{
    FileSelectFile, files, M3, , Select patient lists, Text Document (*.txt)
    if (files = "")
    {
        return false
    }
    return files
}

parseFiles(files)
{
    parentDir = ""
    Loop, parse, files, `n
    {
        if(A_Index = 1) 
        {
            parentDir := A_LoopField
        }
        else
        {
            MsgBox, 4, Reading patient list, The file %A_LoopField% is now being read. Continue?
            ; InputBox, facilityAddress, Information Required, Please provide street address of facility for better search on Fred
            IfMsgBox, No
            {
                continue
            }
            else 
            {
                InputBox, facilityCode, Facility's Code, Please enter the facility's code
                filterFacility(facilityCode)
                InputBox, address, Facility's address, Please enter the new facility address 
                readPatientList(parentDir, A_LoopField, address)
            }
        }
    }
}

filterFacility(facilityCode)
{
    activateFredDispense()
    Loop 5
    {
        Send, {Esc}
        Sleep, 500
    }
    Send, =%facilityCode%{Enter}
    Sleep, 1000
}

readPatientList(parentDir, filename, address)
{
    Loop, read, %parentDir%\%filename%
    {
        activateFredDispense()
        Send, %A_LoopReadLine%
        Sleep, 2000
        Send, {Tab}
        Sleep, 1000
        Send, {Enter}
        Loop 5
        {
            Send, {Esc}
            Sleep, 500
        }
        Send, {F8}
        Sleep, 1000
        Send, {Down 3}
        Sleep, 2000
        Send, %address%
        Send, {End}
        Sleep, 500
        if WinExist("Medicare Card")
        {
            Send, y
            Sleep, 500
        }
    }
}