'V1.3 修正被判斷惡意程式的問題(匯出完成開啟資料夾)
Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long    ''減少等待時電腦卡住(大量產出檔案時會被誤判為惡意程式)

Sub sutExportAllPic()
On Error GoTo ErrorHandler
Dim ExportFolder As String
Dim high As Integer
Dim width As Integer
Dim w As String
Dim h As String
Dim iso As Double
Dim NoteisFileName As Integer
        
    With Application.ActivePresentation
    .Save
    End With 
   
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            ExportFolder = .SelectedItems(1)
        End If
    End With

If ExportFolder = "" Then
Exit Sub
End If

NoteisFileName = MsgBox("匯出的圖檔是否使用備忘錄字串。(否則使用流水號)", vbQuestion + vbYesNo + vbDefaultButton2, "匯出檔名設置")

w = InputBox("輸入匯出圖片寬")

If w <> "" Then
h = InputBox("輸入匯出圖片高")
End If

If w = "" Or h = "" Then

Call GetWH(width, high)

Else
    high = h
    width = w
End If
    
If NoteisFileName = vbYes Then
    Dim pageNumber As Integer
    pageNumber = 1
    For Each NodeStr In GetSolidNoteString()
        If NodeStr <> "" Then
        ActivePresentation.Slides(pageNumber).Export ExportFolder & "\" & NodeStr & ".png", "png", width, high
        Else
        ActivePresentation.Slides(pageNumber).Export ExportFolder & "\" & Format(pageNumber, "000") & ".png", "png", width, high
        End If
        pageNumber = pageNumber + 1
       ' Wait (1)    '大量產出檔案時會被誤判為惡意程式
    Next

Else

    Dim i As Integer, iCount As Integer
    iCount = ActivePresentation.Slides.Count

    For i = 1 To iCount

    ActivePresentation.Slides(i).Export ExportFolder & "\" & Format(i, "000") & ".png", "png", width, high

    Next

End If
 
  
 Dim data As New DataObject
  data.SetText ExportFolder
  data.PutInClipboard
  
  MsgBox "已匯出至資料夾:" & ExportFolder & ",資料夾路徑已傳入剪貼簿中"

'Call Shell("explorer.exe" & " " & ExportFolder, vbNormalFocus)
Exit Sub
ErrorHandler:
    
    MsgBox "尚未輸入正確的寬高值" & vbCrLf & "密技:檔名後加上_寬值_高值，可自動帶入寬高", vbOKOnly + vbCritical, "錯誤1"

End Sub


Sub GetWH(ByRef w As Integer, ByRef h As Integer)

On Error GoTo ErrorHandler
   Dim sPath As String
   Dim str() As String
   Dim sName As String

    sPath = ActivePresentation.fullname
    sPath = FSOGetFileName(sPath)
    If Len(sPath) > 0 Then
        sName = ActivePresentation.Name
        str = Split(sName, "_")
      w = CInt(str(UBound(str) - 1))
         h = CInt(Split(str(UBound(str)), ".")(0))
         
    
    Else
        MsgBox "File not saved"
    End If
   Exit Sub
    
ErrorHandler:

    MsgBox "尚未輸入正確的寬高值" & vbCrLf & "密技:檔名後加上_寬值_高值，可自動帶入寬高", vbOKOnly + vbCritical, "錯誤2"
    w = 0
    h = 0
End Sub


Function GetSolidNoteString()

On Error GoTo ErrorHandler

  Dim oSld As Slide
  Dim sFileName As String
  Dim sDivider As String
  Dim sSlideNote As String
  Dim iReturnPos As Integer
  Dim sLine As String

    Dim tempArr() As String
    ReDim tempArr(ActivePresentation.Slides.Count - 1)
    Dim i As Integer
    i = 0
  ' Treat each slide in the presentation
  For Each oSld In ActivePresentation.Slides
  
  If oSld.NotesPage.Shapes.Placeholders.Count = 1 Then
  tempArr(i) = oSld.NotesPage.Shapes.Placeholders(1).TextFrame.TextRange.Text
  Else
      tempArr(i) = oSld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
  End If
    ' Replace carriage return characters with a CR + LF so that multiple lines get output correctly
      i = i + 1
  Next
  
  GetSolidNoteString = tempArr
   Exit Function
ErrorHandler:
    MsgBox "尚未輸入正確的寬高值" & vbCrLf & "密技:檔名後加上_寬值_高值，可自動帶入寬高", vbOKOnly + vbCritical, "錯誤3"
    GetSolidNoteString = tempArr
  
    
End Function

Function FSOGetFileName(fullname)
    Dim FileName As String
   
    Set FSO = CreateObject("Scripting.FileSystemObject")

    'Get File Name
    FileName = FSO.GetFileName(fullname)
    
    'Get File Name no Extension
    FileNameWOExt = Left(FileName, InStr(FileName, ".") - 1)
    FSOGetFileName = FileNameWOExt
End Function



'減少等待時電腦卡住(大量產出檔案時會被誤判為惡意程式)
Public Sub Wait(Seconds As Double)
    '搭配Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long
    Dim endtime As Double
    endtime = DateTime.Timer + Seconds
    Do
        WaitMessage
        DoEvents
    Loop While DateTime.Timer < endtime
End Sub

