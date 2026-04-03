Attribute VB_Name = "ModuleExport"
'MIT License

'Copyright (c) 2021 - 2026 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Sub SaveSelectedSlides()
    
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
        
        Dim TemporaryPresentation, ThisPresentation As Presentation
        Dim PresentationFilename As String
        Dim SlideLoop As Long
        Dim PresentationSlides As Slide
        Dim DotPosition As Integer
        
        Set ThisPresentation = ActivePresentation
        
        'Delete any previous export tags
        On Error Resume Next
        For Each PresentationSlides In ThisPresentation.Slides
            PresentationSlides.Tags.Delete ("INSTRUMENTA EXPORT")
        Next PresentationSlides
        On Error GoTo 0
        
        'Strip extension from filename
        DotPosition = InStrRev(ThisPresentation.name, ".")
        If DotPosition > 0 Then
            PresentationFilename = left(ThisPresentation.name, DotPosition - 1)
        Else
            PresentationFilename = ThisPresentation.name
        End If
        
        'Set filename and e-mailsubject
        PresentationFilename = PresentationFilename & " (slide "
        
        ProgressForm.Show
        
        For SlideLoop = 1 To ActiveWindow.Selection.SlideRange.count
        
        SetProgress (SlideLoop / ActiveWindow.Selection.SlideRange.count * 100)
        
            'ActiveWindow.Selection.SlideRange(SlideLoop).Tags.Add "INSTRUMENTA EXPORT", "YES" --> does not always work on Mac
            ThisPresentation.Slides(ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex).Tags.Add "INSTRUMENTA EXPORT", "YES"
                        
            If SlideLoop <> ActiveWindow.Selection.SlideRange.count Then
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex & ","
            Else
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex
            End If
        Next SlideLoop
        
        ProgressForm.Hide
        Unload ProgressForm
        
        PresentationFilename = PresentationFilename & ")"
        Dim exportFilePath As String
        
                
        #If Mac Then
        
        exportFilePath = MacSaveAsDialog(PresentationFilename & ".pptx")
        
        #Else
        
        Dim exportFileDialog As FileDialog
        Set exportFileDialog = Application.FileDialog(msoFileDialogSaveAs)
      
        If exportFileDialog.Show = -1 Then
        exportFileDialog.InitialFileName = PresentationFilename & ".pptx"
        exportFilePath = exportFileDialog.SelectedItems(1)
        End If
        
        #End If
        
        'Force pptx
        
        DotPosition = InStrRev(exportFilePath, ".")
        If DotPosition > 0 Then
            exportFilePath = left(exportFilePath, DotPosition - 1) & ".pptx"
        Else
            exportFilePath = exportFilePath & ".pptx"
        End If
        
        
        ThisPresentation.SaveCopyAs exportFilePath, ppSaveAsOpenXMLPresentation
        Set TemporaryPresentation = Presentations.Open(exportFilePath)
        
        ProgressForm.Show
        NumberOfSlides = TemporaryPresentation.Slides.count
        For SlideLoop = TemporaryPresentation.Slides.count To 1 Step -1
            SetProgress ((NumberOfSlides - SlideLoop) / NumberOfSlides * 100)
            If TemporaryPresentation.Slides(SlideLoop).Tags("INSTRUMENTA EXPORT") <> "YES" Then TemporaryPresentation.Slides(SlideLoop).Delete
        Next SlideLoop
        ProgressForm.Hide
        Unload ProgressForm
        
        TemporaryPresentation.Save
        TemporaryPresentation.Close

        
        Else
        MsgBox "No slides selected."
        End If
    
End Sub

Sub EmailSelectedSlides()
    
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
        
        
        #If Not Mac Then
        Dim OutlookApplication, OutlookMessage As Object
        #End If
        
        #If Mac Then
        If Not CheckIfAppleScriptPluginIsInstalled > 0 Then
            MsgBox "Optional Instrumenta AppleScript not found. This function is not supported without it."
            Exit Sub
        End If
        #End If
        
        Dim TemporaryPresentation, ThisPresentation As Presentation
        Dim PresentationFilename, EmailSubject As String
        Dim tempDir As String
        Dim SlideLoop As Long
        Dim PresentationSlides As Slide
        Dim DotPosition As Integer
        
        Set ThisPresentation = ActivePresentation
        
        'Delete any previous export tags
        On Error Resume Next
        For Each PresentationSlides In ThisPresentation.Slides
            PresentationSlides.Tags.Delete ("INSTRUMENTA EXPORT")
        Next PresentationSlides
        On Error GoTo 0
        
        'Strip extension from filename
        DotPosition = InStrRev(ThisPresentation.name, ".")
        If DotPosition > 0 Then
            PresentationFilename = left(ThisPresentation.name, DotPosition - 1)
        Else
            PresentationFilename = ThisPresentation.name
        End If
        
        'Set filename and e-mailsubject
        EmailSubject = PresentationFilename
        PresentationFilename = PresentationFilename & " (slide "
        
        ProgressForm.Show
        
        For SlideLoop = 1 To ActiveWindow.Selection.SlideRange.count
        
        SetProgress (SlideLoop / ActiveWindow.Selection.SlideRange.count * 100)
        
            'ActiveWindow.Selection.SlideRange(SlideLoop).Tags.Add "INSTRUMENTA EXPORT", "YES" --> Does not always work on Mac
            ThisPresentation.Slides(ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex).Tags.Add "INSTRUMENTA EXPORT", "YES"
            
            If SlideLoop <> ActiveWindow.Selection.SlideRange.count Then
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex & ","
            Else
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex
            End If
        Next SlideLoop
        
        ProgressForm.Hide
        Unload ProgressForm
        
        PresentationFilename = PresentationFilename & ")"
        
        PresentationFilename = SanitizeFilename(InputBox("Attachment file name:", "Send as e-mail", PresentationFilename))
        If Len(PresentationFilename) = 0 Then Exit Sub
        
        #If Mac Then
        
        tempDir = MacScript("return posix path of (path to temporary items) as string")
        
        If PresentationFilename & ".pptx" = ThisPresentation.name Then
        PresentationFilename = PresentationFilename & "_1"
        End If
        
        ThisPresentation.SaveCopyAs tempDir & PresentationFilename & ".pptx"
        Set TemporaryPresentation = Presentations.Open(tempDir & PresentationFilename & ".pptx")
        #Else
        tempDir = Environ("TEMP") & "\"
        ThisPresentation.SaveCopyAs tempDir & PresentationFilename & ".pptx"
        Set TemporaryPresentation = Presentations.Open(tempDir & PresentationFilename & ".pptx")
        #End If
        
        ProgressForm.Show
        NumberOfSlides = TemporaryPresentation.Slides.count
        For SlideLoop = TemporaryPresentation.Slides.count To 1 Step -1
            SetProgress ((NumberOfSlides - SlideLoop) / NumberOfSlides * 100)
            If TemporaryPresentation.Slides(SlideLoop).Tags("INSTRUMENTA EXPORT") <> "YES" Then TemporaryPresentation.Slides(SlideLoop).Delete
        Next SlideLoop
        ProgressForm.Hide
        Unload ProgressForm
        
        TemporaryPresentation.Save
        TemporaryPresentation.Close
               
        #If Mac Then
        'This does not work anymore
        'OutlookMessageMac = MacSendMailViaOutlook(EmailSubject, ActivePresentation.Path & "/" & PresentationFilename & ".pptx")

        Dim ParamString As String
        Dim OutlookMessageMac As String

        ParamString = BuildAppleScriptParam(EmailSubject, tempDir & PresentationFilename & ".pptx")
        If ParamString = "" Then
            Kill (tempDir & PresentationFilename & ".pptx")
            MsgBox "Temporary file path contains unsupported characters."
            Exit Sub
        End If
        OutlookMessageMac = AppleScriptTask("InstrumentaAppleScriptPlugin.applescript", "SendFileWithOutlook", CStr(ParamString))

        Kill (tempDir & PresentationFilename & ".pptx")
        #Else
               
        On Error Resume Next
        Set OutlookApplication = GetObject(Class:="Outlook.Application")
        Err.Clear
        If OutlookApplication Is Nothing Then Set OutlookApplication = CreateObject(Class:="Outlook.Application")
        On Error GoTo 0
        Set OutlookMessage = OutlookApplication.CreateItem(0)
        
        On Error Resume Next
        With OutlookMessage
            .To = ""
            .cc = ""
            .subject = EmailSubject
            .Body = ""
            .Attachments.Add tempDir & PresentationFilename & ".pptx"
            .Display
        End With
         On Error GoTo 0
        
        'Delete temporary slides
        Kill (tempDir & PresentationFilename & ".pptx")
        
        #End If
        
        Else
        MsgBox "No slides selected."
        End If
    
End Sub

Sub EmailSelectedSlidesAsPDF()
        
        If ActiveWindow.Selection.Type = ppSelectionSlides Then
        
        Set ThisPresentation = ActivePresentation
        
        #If Not Mac Then
        Dim OutlookApplication, OutlookMessage As Object
        #End If
        
        #If Mac Then
        If Not CheckIfAppleScriptPluginIsInstalled > 0 Then
            MsgBox "Optional Instrumenta AppleScript not found. This function is not supported without it."
            Exit Sub
        End If
        #End If
             
        Dim PresentationFilename, EmailSubject As String
        Dim tempDir As String
        Dim SlideLoop As Long
        Dim DotPosition As Integer
        
        DotPosition = InStrRev(ActivePresentation.name, ".")
        
        If DotPosition > 0 Then
            PresentationFilename = left(ActivePresentation.name, DotPosition - 1)
        Else
            PresentationFilename = ActivePresentation.name
        End If
        
        On Error Resume Next
        For Each PresentationSlides In ThisPresentation.Slides
            PresentationSlides.Tags.Delete ("INSTRUMENTA EXPORT")
        Next PresentationSlides
        On Error GoTo 0
        
        'Set filename and e-mailsubject
        EmailSubject = PresentationFilename
        PresentationFilename = PresentationFilename & " (slide "
        
        ProgressForm.Show
        
        For SlideLoop = 1 To ActiveWindow.Selection.SlideRange.count
        
        SetProgress (SlideLoop / ActiveWindow.Selection.SlideRange.count * 100)
        
            'ActiveWindow.Selection.SlideRange(SlideLoop).Tags.Add "INSTRUMENTA EXPORT", "YES" --> Does not always work on Mac
            ThisPresentation.Slides(ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex).Tags.Add "INSTRUMENTA EXPORT", "YES"
            
            If SlideLoop <> ActiveWindow.Selection.SlideRange.count Then
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex & ","
            Else
                PresentationFilename = PresentationFilename & ActiveWindow.Selection.SlideRange(SlideLoop).slideIndex
            End If
        Next SlideLoop
        
        ProgressForm.Hide
        Unload ProgressForm
        
        PresentationFilename = PresentationFilename & ")"
        
        PresentationFilename = SanitizeFilename(InputBox("Attachment file name:", "Send as e-mail", PresentationFilename))
        If Len(PresentationFilename) = 0 Then Exit Sub
      
      
        #If Mac Then
        'This does not work anymore
        'OutlookMessageMac = MacSendMailViaOutlook(EmailSubject, ActivePresentation.Path & "/" & PresentationFilename & ".pptx")
    
        tempDir = MacScript("return posix path of (path to temporary items) as string")
        
        
        ThisPresentation.SaveCopyAs tempDir & PresentationFilename & "_temp.pptx"
        Set TemporaryPresentation = Presentations.Open(tempDir & PresentationFilename & "_temp.pptx")
        
        ProgressForm.Show
        NumberOfSlides = TemporaryPresentation.Slides.count
        For SlideLoop = TemporaryPresentation.Slides.count To 1 Step -1
            SetProgress ((NumberOfSlides - SlideLoop) / NumberOfSlides * 100)
            If TemporaryPresentation.Slides(SlideLoop).Tags("INSTRUMENTA EXPORT") <> "YES" Then TemporaryPresentation.Slides(SlideLoop).Delete
        Next SlideLoop
        ProgressForm.Hide
        Unload ProgressForm
        
        TemporaryPresentation.Save
        TemporaryPresentation.SaveCopyAs tempDir & PresentationFilename & ".pdf", ppSaveAsPDF
        TemporaryPresentation.Close
        Kill (tempDir & PresentationFilename & "_temp.pptx")

        Dim ParamString As String
        Dim OutlookMessageMac As String

        ParamString = BuildAppleScriptParam(EmailSubject, tempDir & PresentationFilename & ".pdf")
        If ParamString = "" Then
            Kill (tempDir & PresentationFilename & ".pdf")
            MsgBox "Temporary file path contains unsupported characters."
            Exit Sub
        End If
        OutlookMessageMac = AppleScriptTask("InstrumentaAppleScriptPlugin.applescript", "SendFileWithOutlook", CStr(ParamString))

        Kill (tempDir & PresentationFilename & ".pdf")
        
        #Else
        
        tempDir = Environ("TEMP") & "\"
        
        ActivePresentation.ExportAsFixedFormat tempDir & PresentationFilename & ".pdf", ppFixedFormatTypePDF, ppFixedFormatIntentPrint, msoFalse, , , , , ppPrintSelection

        On Error Resume Next
        Set OutlookApplication = GetObject(Class:="Outlook.Application")
        Err.Clear
        If OutlookApplication Is Nothing Then Set OutlookApplication = CreateObject(Class:="Outlook.Application")
        On Error GoTo 0
        Set OutlookMessage = OutlookApplication.CreateItem(0)
        
        On Error Resume Next
        With OutlookMessage
            .To = ""
            .cc = ""
            .subject = EmailSubject
            .Body = ""
            .Attachments.Add tempDir & PresentationFilename & ".pdf"
            .Display
        End With
        
        On Error GoTo 0
        
        'Clean temporary PDF
        Kill (tempDir & PresentationFilename & ".pdf")
        
        #End If

        Else
        MsgBox "No slides selected."
        End If
        

    
End Sub

Private Function BuildAppleScriptParam(emailSubject As String, emailAttachment As String) As String
    Dim safeSubject As String
    safeSubject = Replace(emailSubject, ";", ",")
    safeSubject = Replace(safeSubject, vbCr, " ")
    safeSubject = Replace(safeSubject, vbLf, " ")
    safeSubject = Trim(safeSubject)
    If Len(safeSubject) = 0 Then
        safeSubject = "(no subject)"
    End If
    If InStr(emailAttachment, ";") > 0 Or InStr(emailAttachment, vbCr) > 0 Or InStr(emailAttachment, vbLf) > 0 Then
        BuildAppleScriptParam = ""
        Exit Function
    End If
    BuildAppleScriptParam = safeSubject & ";" & SanitizeAppleScriptPath(emailAttachment)
End Function
