Attribute VB_Name = "m_Export_For_Git"
Option Explicit

Public Sub ExportForGit()

    Dim strExportFolder$
    Dim Fs As Object
    Dim vC As VBComponent
    Dim v, t
    Dim c%

    strExportFolder = "C:\Users\e008922\Dropbox\_Git\##REPO-FOLDER-HERE##"

    For Each vC In ActiveWorkbook.VBProject.VBComponents
        v = ExportVBComponent(vC, strExportFolder, , True)
        If v <> True Then Call MsgBox("Problem with " & vC.Name & " export :(")
        't = Timer
        'While Timer < t + 0.05
        '    DoEvents
        'Wend
        c = c + 1
    Next
    
    Application.ActiveWorkbook.SaveCopyAs "C:\Users\e008922\Dropbox\_Git\##REPO-FOLDER-HERE##\FILE NAME HERE.xlsm"

End Sub




'	IF ERROR ON UNKNOWN TYPE:
'		- Tools --> References...
'		- Add "Microsoft Visual Basic for Applications Extensibility 5.3"



'   From:   http://www.cpearson.com/excel/vbe.aspx
'           "Exporting A VBComponent Code Module To A Text File"



Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, _
                                  FolderName As String, _
                                  Optional FileName As String, _
                                  Optional OverwriteExisting As Boolean = True) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This function exports the code module of a VBComponent to a text
    ' file. If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim FName As String
    'Extension = ".txt"
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(FileName) = vbNullString Then
        If StrComp(Left(VBComp.Name, 5), "Sheet", vbTextCompare) = 0 And Extension = ".cls" Then
            FName = Replace(VBComp.Properties.Item("Name"), " ", "_", 1, -1, vbTextCompare) & Extension
        Else
            FName = VBComp.Name & Extension
        End If
    Else
        FName = FileName
        If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
            FName = FName & Extension
        End If
    End If

    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        FName = FolderName & FName
    Else
        FName = FolderName & "\" & FName
    End If

    If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill FName
        Else
            ExportVBComponent = False
            Exit Function
        End If
    End If

    On Error GoTo ErrCantExport
    VBComp.Export FileName:=FName
    ExportVBComponent = True
    On Error GoTo 0

ErrCantExport:

    If Not ExportVBComponent Then
        Dim a$, b$, c$, i As Integer

        a$ = Mid(FName, Len(FolderName))
        If InStr(1, a$, "\", vbTextCompare) > 0 Then a$ = Mid(a$, InStr(1, a$, "\", vbTextCompare) + 1)
        For i = 1 To Len(a$)
            b$ = Mid(a$, i, 1)
            If b$ Like "[A-Z,a-z,0-9 _-]" Then
                c$ = c$ & b$
            End If
        Next i

        If StrComp(Right(FolderName, 1), "\", vbTextCompare) = 0 Then
            FName = FolderName & c$
        Else
            FName = FolderName & "\" & c$
        End If

        Resume
    End If

End Function




Public Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns the appropriate file extension based on the Type of
    ' the VBComponent.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select

End Function




