run() 

run() 
{
    outputFolder := selectOutputFolder()
    if (outputFolder = "")
    {
        MsgBox, 0, Information, No destination folder was selected.
        return   
    }

    files := selectFiles()
    if (files = "")
    {
        MsgBox, 0, Information, No file was selected.
        return 
    }

    readFiles(files, outputFolder)
}

selectOutputFolder()
{
    FileSelectFolder, destination,,,Select output folder
    return destination
}

selectFiles()
{
    FileSelectFile, files, M3, RootDir\Filename, Select files containing the item barcodes, Text Document (*.txt)
    return files
}

readFiles(files, destination)
{
    path := ""
    Loop, parse, files, `n
    {
        if (A_Index = 1)
        {
            path := A_LoopField
        }
        else
        {
            readFile(path, A_LoopField)
            filename := createCSVFilename(A_LoopField)
            saveFile(filename, destination)
        }
    }
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
    openNewNotepad()
}

readFile(path, file)
{
    Loop, Read, %path%\%file%
    {
        ; MsgBox, %A_LoopReadLine%
        activateFredOffice()
        clearSearchBar()
        searchItemOnFredOffice(A_LoopReadLine)
        openItemWindow()
        item := getItemName()
        closeItemWindow()
        activateNotepad()
        writeToNotepad(item, A_LoopReadLine)
    }
}

activateFredOffice()
{
    if WinExist("Fred Office for Pharmacy (soserver)")
    {
        WinActivate, Fred Office for Pharmacy (soserver)
    }
}

activateNotepad()
{
    if WinExist("Untitled - Notepad")
    {
        WinActivate, Untitled - Notepad
    }
}

searchItemOnFredOffice(itemBarcode)
{
    Click, left, 500, 60, 1
    Send, %itemBarcode%{Enter}
    Sleep, 500
}

writeToNotepad(item, barcode)
{
    if (item = "")
    {
        Send, Could NOT find item`,
    }
    else 
    {
        Send, %item%`,
    }
    Send, %barcode%{Enter}
}

createCSVFilename(file)
{
    return StrReplace(file, ".txt", " details.csv" )
}

clearSearchBar()
{
    CoordMode, Mouse, Client
    Click, left, 805, 60, 1
}

getItemName()
{
    ControlGetText, itemName, WindowsForms10.EDIT.app.0.12ab327_r6_ad112, Item 
    return itemName
}

closeItemWindow()
{
    Click, left, 820, -10, 1
}

openItemWindow()
{
    Click, left, 500, 130, 2
    Sleep, 1500
}

openNewNotepad()
{
    Send, ^n
}