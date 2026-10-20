
#Requires AutoHotkey v2.0

SetTitleMatchMode(2)
SetKeyDelay(150, 100)

waterCADTitle := "WaterCAD"
outputFolder := "C:\Users\Hadi Wehbi\Desktop\WaterCAD_Input"
outputFile := outputFolder "\pipe.csv"

if !WinExist(waterCADTitle)
{
    MsgBox("WaterCAD window not found.")
    ExitApp()
}

waterCAD := WinExist(waterCADTitle)
WinActivate(waterCAD)

if !WinWaitActive(waterCAD, , 5)
{
    MsgBox("Could not activate WaterCAD.")
    ExitApp()
}

Sleep(1000)

; Record the existing output file's modification time
oldTime := FileExist(outputFile)
    ? FileGetTime(outputFile, "M")
    : ""

; Open File > Export > Export to Excel
Send("{Esc}")
Sleep(300)
Send("{Alt}")
Sleep(500)
Send("{Down}")
Sleep(500)
Send("t")
Sleep(500)
Send("x")

; Wait for the Export to Excel dialog
Sleep(1000)

; Move to OK and open Flex Table Export
Send("{Tab 8}")
Sleep(300)
Send("{Enter}")
Sleep(1200)

; Change Format from Excel to CSV
Send("{Down}")
Sleep(500)

; Move to Folder field (contents are already highlighted)
Send("{Tab}")
Sleep(300)
SendText(outputFolder)
Sleep(300)

; Skip Browse and move to OK
Send("{Tab 2}")
Sleep(300)
Send("{Enter}")

; Wait for export to finish
Sleep(2000)

if !FileExist(outputFile)
{
    MsgBox("Export failed: output file not found.`n`n" outputFile)
    ExitApp()
}

newTime := FileGetTime(outputFile, "M")

if (newTime != oldTime)
{
    MsgBox("Export completed successfully!`n`n"
        "File: " outputFile
        "`nPrevious modified time: " (oldTime = "" ? "File did not exist" : oldTime)
        "`nNew modified time: " newTime)
}
else
{
    MsgBox("The output file exists, but its modified time did not change.`n`n"
        "Check whether the export completed.`n`n"
        "File: " outputFile
        "`nModified time: " newTime)
}
