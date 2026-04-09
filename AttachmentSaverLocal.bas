Attribute VB_Name = "AttachmentSaverLocal"
Option Explicit

Public Const DEFAULT_SAVE_PATH As String = "C:\EmailAttachments"

Public Sub SaveAttachments_DefaultFolder()
    SaveSelectedMailAttachments False
End Sub

Public Sub SaveAttachments_SaveAs()
    SaveSelectedMailAttachments True
End Sub

Private Sub SaveSelectedMailAttachments(ByVal askForFolder As Boolean)
    On Error GoTo EH

    Dim exp As Outlook.Explorer
    Dim sel As Outlook.Selection
    Dim itm As Object
    Dim mail As Outlook.MailItem
    Dim att As Outlook.Attachment
    Dim basePath As String
    Dim savedCount As Long
    Dim skippedCount As Long
    Dim errorCount As Long

    Set exp = Application.ActiveExplorer
    If exp Is Nothing Then
        MsgBox "Kein aktives Outlook-Fenster gefunden.", vbExclamation, "Attachment Saver"
        Exit Sub
    End If

    Set sel = exp.Selection
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "Bitte zuerst eine oder mehrere E-Mails auswählen.", vbInformation, "Attachment Saver"
        Exit Sub
    End If

    basePath = DEFAULT_SAVE_PATH

    If askForFolder Then
        basePath = PickFolder(basePath)
        If Len(basePath) = 0 Then Exit Sub
    End If

    EnsureFolderTreeExists basePath

    For Each itm In sel
        If TypeName(itm) = "MailItem" Then
            Set mail = itm

            If mail.Attachments.Count = 0 Then
                skippedCount = skippedCount + 1
            Else
                For Each att In mail.Attachments
                    On Error Resume Next
                    att.SaveAsFile GetUniqueFilePath(basePath, CleanFileName(att.fileName))
                    If Err.Number <> 0 Then
                        errorCount = errorCount + 1
                        Err.Clear
                    Else
                        savedCount = savedCount + 1
                    End If
                    On Error GoTo EH
                Next att
            End If
        End If
    Next itm

    MsgBox savedCount & " Anhang/Anhänge gespeichert in:" & vbCrLf & basePath & _
           IIf(skippedCount > 0, vbCrLf & skippedCount & " E-Mail(s) ohne Anhänge übersprungen", "") & _
           IIf(errorCount > 0, vbCrLf & errorCount & " Fehler aufgetreten", ""), _
           IIf(errorCount > 0, vbExclamation, vbInformation), "Attachment Saver"
    Exit Sub

EH:
    MsgBox "Fehler: " & Err.Description, vbCritical, "Attachment Saver"
End Sub

Private Function GetUniqueFilePath(ByVal folderPath As String, ByVal fileName As String) As String
    Dim fullPath As String
    Dim baseName As String
    Dim ext As String
    Dim i As Long

    fullPath = AddTrailingSlash(folderPath) & fileName
    If Dir(fullPath, vbNormal) = "" Then
        GetUniqueFilePath = fullPath
        Exit Function
    End If

    baseName = GetFileBaseName(fileName)
    ext = GetFileExtension(fileName)
    i = 1

    Do
        fullPath = AddTrailingSlash(folderPath) & baseName & "_" & i & ext
        i = i + 1
    Loop While Dir(fullPath, vbNormal) <> ""

    GetUniqueFilePath = fullPath
End Function

Private Function GetFileBaseName(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then
        GetFileBaseName = Left$(fileName, p - 1)
    Else
        GetFileBaseName = fileName
    End If
End Function

Private Function GetFileExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then
        GetFileExtension = Mid$(fileName, p)
    Else
        GetFileExtension = ""
    End If
End Function

Private Function CleanFileName(ByVal value As String) As String
    Dim badChars As Variant
    Dim i As Long
    badChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")

    value = Trim$(value)
    If Len(value) = 0 Then value = "_"

    For i = LBound(badChars) To UBound(badChars)
        value = Replace$(value, badChars(i), "_")
    Next i

    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    value = NormalizeSpaces(value)

    If Len(value) > 120 Then value = Left$(value, 120)
    CleanFileName = value
End Function

Private Function NormalizeSpaces(ByVal value As String) As String
    Do While InStr(value, "  ") > 0
        value = Replace$(value, "  ", " ")
    Loop
    NormalizeSpaces = Trim$(value)
End Function

Private Function AddTrailingSlash(ByVal path As String) As String
    If Right$(path, 1) = "\" Then
        AddTrailingSlash = path
    Else
        AddTrailingSlash = path & "\"
    End If
End Function

Private Sub EnsureFolderTreeExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Function PickFolder(Optional ByVal initialPath As String = "") As String
    Dim sh As Object
    Dim folder As Object

    Set sh = CreateObject("Shell.Application")
    Set folder = sh.BrowseForFolder(0, "Zielordner für Anhänge auswählen", 0, initialPath)

    If folder Is Nothing Then
        PickFolder = ""
    Else
        PickFolder = folder.Self.path
    End If
End Function

