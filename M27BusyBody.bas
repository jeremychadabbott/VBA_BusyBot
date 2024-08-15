Option Explicit

' Constants for task weights
Const WEIGHT_PROCESS_INVOICE As Integer = 75
Const WEIGHT_LOOK_AT_EMAIL As Integer = 15
Const WEIGHT_LOAD_DRAWINGS As Integer = 5
Const WEIGHT_LOAD_COMMITTED_COST_REPORT As Integer = 5

' Enum for task types
Enum TaskType
    ProcessInvoice = 1
    LookAtEmail = 2
    LoadDrawings = 3
    LoadCommittedCostReport = 4
End Enum

Sub Busy_Body(ByRef emailmessage As String)
    emailmessage = "busy"
    
    Do
        Dim taskType As TaskType
        taskType = SelectRandomTask()
        
        Select Case taskType
            Case TaskType.ProcessInvoice
                ProcessInvoice emailmessage
            Case TaskType.LookAtEmail
                LookAtEmail
            Case TaskType.LoadDrawings
                LoadDrawings
            Case TaskType.LoadCommittedCostReport
                LoadCommittedCostReport
        End Select
    Loop
End Sub

Function SelectRandomTask() As TaskType
    Dim totalWeight As Integer
    totalWeight = WEIGHT_PROCESS_INVOICE + WEIGHT_LOOK_AT_EMAIL + WEIGHT_LOAD_DRAWINGS + WEIGHT_LOAD_COMMITTED_COST_REPORT
    
    Randomize
    Dim randomValue As Integer
    randomValue = Int(totalWeight * Rnd + 1)
    
    Dim cumulativeWeight As Integer
    cumulativeWeight = 0
    
    If randomValue <= (cumulativeWeight + WEIGHT_PROCESS_INVOICE) Then
        SelectRandomTask = TaskType.ProcessInvoice
    ElseIf randomValue <= (cumulativeWeight + WEIGHT_PROCESS_INVOICE + WEIGHT_LOOK_AT_EMAIL) Then
        SelectRandomTask = TaskType.LookAtEmail
    ElseIf randomValue <= (cumulativeWeight + WEIGHT_PROCESS_INVOICE + WEIGHT_LOOK_AT_EMAIL + WEIGHT_LOAD_DRAWINGS) Then
        SelectRandomTask = TaskType.LoadDrawings
    Else
        SelectRandomTask = TaskType.LoadCommittedCostReport
    End If
End Function

Sub ProcessInvoice(ByRef emailmessage As String)
    Dim folderPaths() As String
    folderPaths = Array("\\server2\Faxes\PLATT - 234\Backup", _
                        "\\server2\Faxes\NORTH COAST - 218\Backup", _
                        "\\server2\Faxes\WESCO - 430\Backup")
    
    Dim selectedFolderPath As String
    selectedFolderPath = folderPaths(Int((UBound(folderPaths) - LBound(folderPaths) + 1) * Rnd + LBound(folderPaths)))
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not objFSO.FolderExists(selectedFolderPath) Then
        MsgBox "Selected folder does not exist: " & selectedFolderPath
        Exit Sub
    End If
    
    Dim objFolder As Object
    Set objFolder = objFSO.GetFolder(selectedFolderPath)
    
    Dim colInvoices As Collection
    Set colInvoices = New Collection
    
    Dim dtOneWeekAgo As Date
    dtOneWeekAgo = DateAdd("d", -7, Now)
    
    Dim objFile As Object
    For Each objFile In objFolder.Files
        If InStr(1, objFile.Name, "INV", vbTextCompare) > 0 And objFile.DateCreated > dtOneWeekAgo Then
            colInvoices.Add objFile.Path
        End If
    Next objFile
    
    If colInvoices.Count = 0 Then
        MsgBox "No invoices found within the past week in the selected folder: " & selectedFolderPath
        Exit Sub
    End If
    
    Dim selectedInvoicePath As String
    selectedInvoicePath = colInvoices(Int((colInvoices.Count - 1 + 1) * Rnd + 1))
    
    If Not objFSO.FileExists(selectedInvoicePath) Then
        MsgBox "Selected invoice not found."
        Exit Sub
    End If
    
    Dim newFileName As String
    newFileName = GetNewFileName(selectedInvoicePath)
    
    objFSO.CopyFile selectedInvoicePath, "\\server2\Dropbox\Attachments\" & newFileName
    
    Dim path As String
    path = "\\server2\Dropbox\Attachments"
    Call modle20(xoffset, path, emailmessage)
    
    Sleep 10000 ' Sleep for 10 seconds
    
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set colInvoices = Nothing
End Sub

Function GetNewFileName(ByVal filePath As String) As String
    If InStr(1, filePath, "PLATT", vbTextCompare) > 0 Then
        GetNewFileName = "1234_INVOICE_1234.pdf"
    ElseIf InStr(1, filePath, "NORTH COAST", vbTextCompare) > 0 Then
        GetNewFileName = "Northcoast.pdf"
    ElseIf InStr(1, filePath, "WESCO", vbTextCompare) > 0 Then
        GetNewFileName = "Wesco.pdf"
    Else
        GetNewFileName = "unknown.pdf"
    End If
End Function

Sub LookAtEmail()
    Const TargetURL As String = "https://outlook.office.com/mail/"
    Call OpenChrome(TargetURL)
    Application.Wait Now + TimeValue("00:00:10")
    
    Dim i As Integer
    For i = 1 To 5
        Application.SendKeys "{Down}"
        
        Dim Delay As Integer
        Delay = Int((20 - 5 + 1) * Rnd + 5)
        Application.Wait Now + TimeValue("00:00:" & Format(Delay, "00"))
    Next i
    
    For i = 1 To 3
        Application.SendKeys "^w"
        Application.Wait Now + TimeValue("00:00:02")
    Next i
End Sub

Sub LoadDrawings()
    Const PMmasterLocation As String = "\\server2\Dropbox\Jeremy Abbott\PM assistant (Master).xlsm"
    ' Implement drawing loading logic here
End Sub

Sub LoadCommittedCostReport()
    ' Implement committed cost report loading logic here
End Sub
