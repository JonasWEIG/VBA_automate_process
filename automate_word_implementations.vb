Option Explicit

Sub CreateWordDocuments()
Dim CustRow, CustCol, LastRow, TemplRow, DaysSince, FrDays, ToDays As Long
Dim DocLoc, TagName, TagValue, TemplName, FileName, olOldBody As String
Dim CurDt, LastAppDt As Date
Dim WordDoc, WordApp, OutApp, OutMail As Object
Dim WordContent As Word.Range

Call Einfügen
Call UmlauteKorrigieren


With Tabelle1

    If .Range("B3").Value = Empty Then
        MsgBox "Please select a current template from the drop down list"
        .Range("G3").Select
        Exit Sub
    End If
        TemplRow = .Range("B3").Value 'Set Template Row
        TemplName = .Range("G3").Value 'Set Template Name
        DocLoc = Tabelle2.Range("F" & TemplRow).Value 'Word Document Filename
        
        'Open Word Template
        On Error Resume Next 'If Word is already running
        Set WordApp = GetObject("Word.Application")
        If Err.Number <> 0 Then
        'Launch a new instance of Word
        Err.Clear
        'On Error GoTo Error_Handler
        Set WordApp = CreateObject("Word.application")
        WordApp.Visible = True 'Make the application visible to the user
        End If
        
        
        LastRow = .Range("E9999").End(xlUp).Row 'Determine Last Row in Table
            For CustRow = 8 To LastRow
                If .Range("AI" & CustRow).Value = "" Then
                
                        Set WordDoc = WordApp.Documents.Open(FileName:=DocLoc, ReadOnly:=False) 'Open Template
                        For CustCol = 4 To 39 'Move Through Columns
                            TagName = .Cells(7, CustCol).Value 'Tag Name
                            TagValue = .Cells(CustRow, CustCol).Value 'Tag Value
                             With WordDoc.Content.Find
                                .Text = TagName
                                .Replacement.Text = TagValue
                                .Wrap = wdFindContinue
                                .Execute Replace:=wdReplaceAll ', Forward:=True, Wrap:=wdFindContinue
                            End With
                        Next CustCol
                
                    If .Range("I3").Value = "PDF" Then
                        FileName = " path " & "\" & .Range("E" & CustRow).Value & "_" & .Range("F" & CustRow).Value & ".pdf" 'Create full filename & Path 
                        WordDoc.ExportAsFixedFormat OutputFileName:=FileName, ExportFormat:=wdExportFormatPDf
                        WordDoc.Close False
                    Else: 'If Word
                        FileName = ThisWorkbook.Path & "\" & .Range("E" & CustRow).Value & "_" & .Range("F" & CustRow).Value & ".docx"
                        WordDoc.SaveAs FileName
                    
                    End If
                        'Template Name
                        .Range("AI" & CustRow).Value = Now
            If .Range("P3").Value = "Email" Then
                Set OutApp = CreateObject("Outlook.Application") 'Create Outlook appl
                Set OutMail = OutApp.CreateItem(0) 'Create Email
                With OutMail
                    .GetInspector.Display
                    olOldBody = .htmlBody
                    .To = Tabelle1.Range("AF" & CustRow).Value
                    .Subject = 'Add subject
                    .htmlBody = 'Add body
                    .Attachments.Add FileName
                    .Display 'To send without Displaying change .Display to .Send
                End With
                
            Else:
                WordDoc.PrintOut
                WordDoc.Close
            End If
               
      End If
      
    Next CustRow
    WordApp.Quit
End With

End Sub


Sub Call_data()
'
' Call raw data from workbook 1
'
    On Error GoTo PROBLEM
'
    Workbooks(1).Activate
    Range("D2:AD2").Select
    Selection.Copy
    Workbooks(2).Activate
    Range("D6").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    
    Exit Sub
    
PROBLEM:
    MsgBox ("Please close all uninvolved Excel files! Then close SQ - ModTool - file and open it again.")
    End

    
End Sub

Sub correct_umlaut()
'
' corrects german letters ä,ö,ü
'



    Cells.Replace What:="Ã¶", Replacement:="ö", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="Ã¼", Replacement:="ü", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="Ã¤", Replacement:="ä", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="ÃŸ", Replacement:="ß", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="Ã„", Replacement:="Ä", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        
End Sub
