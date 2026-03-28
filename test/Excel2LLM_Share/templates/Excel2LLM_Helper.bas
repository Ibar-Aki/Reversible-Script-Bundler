Attribute VB_Name = "Excel2LLM_Helper"
Option Explicit

Public Sub Excel2LLM_RecalculateWorkbook()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.CalculateFullRebuild
    DoEvents
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Public Sub Excel2LLM_ExportDisplaySnapshot(ByVal outputPath As String)
    Dim ws As Worksheet
    Dim usedRange As Range
    Dim cell As Range
    Dim fileNo As Integer
    Dim lineText As String

    fileNo = FreeFile
    Open outputPath For Output As #fileNo
    Print #fileNo, "Sheet" & vbTab & "Address" & vbTab & "Text" & vbTab & "Formula"

    For Each ws In ThisWorkbook.Worksheets
        Set usedRange = ws.UsedRange
        For Each cell In usedRange.Cells
            lineText = ws.Name & vbTab & cell.Address(False, False) & vbTab & Replace(cell.Text, vbTab, " ") & vbTab & Replace(cell.Formula, vbTab, " ")
            Print #fileNo, lineText
        Next cell
    Next ws

    Close #fileNo
End Sub
