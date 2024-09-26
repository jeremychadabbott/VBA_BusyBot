
Sub Busy_Body(emailmessage)
    emailmessage = "busy"
    
    Dim randomTask As Integer
    Dim Case1 As Integer, Case2 As Integer, Case3 As Integer, Case4 As Integer
    Dim totalWeight As Integer, randomValue As Integer, cumulativeWeight As Integer

select_Task:
    ' This label is used for looping back to task selection

    Case1 = 75
    Case2 = 15
    Case3 = 5
    Case4 = 5
    
    totalWeight = Case1 + Case2 + Case3 + Case4
    
    Randomize
    randomValue = Int(totalWeight * Rnd + 1)
    
    cumulativeWeight = 0
    
    If randomValue <= (cumulativeWeight + Case1) Then
        GoTo Task_Process_Invoice
    ElseIf randomValue <= (cumulativeWeight + Case1 + Case2) Then
        GoTo Task_Look_at_Email
    ElseIf randomValue <= (cumulativeWeight + Case1 + Case2 + Case3) Then
        GoTo Task_Load_Drawings
    Else
        GoTo Task_Load_Committed_Cost_Report
    End If
    
    
Task_Process_Invoice:
    Dim folderPaths(1 To 3) As String
    folderPaths(1) = "\\server2\Faxes\PLATT - 234\Backup"
    folderPaths(2) = "\\server2\Faxes\NORTH COAST - 218\Backup"
    folderPaths(3) = "\\server2\Faxes\WESCO - 430\Backup"
    
    Dim selectedFolderPath As String
    Randomize
    selectedFolderPath = folderPaths(Int((3 - 1 + 1) * Rnd + 1))
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FolderExists(selectedFolderPath) Then
        Dim objFolder As Object
        Set objFolder = objFSO.GetFolder(selectedFolderPath)
        
        Dim colInvoices As Collection
        Set colInvoices = New Collection
        
        Dim dtOneWeekAgo As Date
        dtOneWeekAgo = DateAdd("d", -7, Now)
        
        Dim objFile As Object
        For Each objFile In objFolder.Files
            If InStr(1, objFile.Name, "INV", vbTextCompare) > 0 And objFile.DateCreated > dtOneWeekAgo Then
                colInvoices.Add objFile.path
            End If
        Next objFile
        
        If colInvoices.Count > 0 Then
            Dim randomIndex As Integer
            Randomize
            randomIndex = Int((colInvoices.Count - 1 + 1) * Rnd + 1)
            
            Dim selectedInvoicePath As String
            selectedInvoicePath = colInvoices(randomIndex)
            
            If objFSO.FileExists(selectedInvoicePath) Then
                Dim newFileName As String
                Dim folderName As String
                folderName = Split(Mid(selectedFolderPath, InStrRev(selectedFolderPath, "\") + 1), " - ")(0)
                
                Select Case True
                    Case selectedInvoicePath Like "*PLATT*"
                        newFileName = "1234_INVOICE_1234.pdf"
                    Case selectedInvoicePath Like "*NORTH*COAST*"
                        newFileName = "Northcoast.pdf"
                    Case selectedInvoicePath Like "*WESCO*"
                        newFileName = "Wesco.pdf"
                    Case Else
                        newFileName = "unknown.pdf"
                End Select
                
                objFSO.CopyFile selectedInvoicePath, "\\server2\Dropbox\Attachments\" & newFileName
                
                Dim path As String
                path = "\\server2\Dropbox\Attachments"
                Call modle20(xoffset, path, emailmessage)
                
                For Repeat = 1 To 10
                    Sleep 1000
                Next Repeat
                
                GoTo select_Task
            Else
                MsgBox "Selected invoice not found."
            End If
        Else
            'MsgBox "No invoices found within the past week in the selected folder: " & selectedFolderPath
        End If
    Else
        MsgBox "Selected folder does not exist: " & selectedFolderPath
    End If
    
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set colInvoices = Nothing
    
    GoTo select_Task
    
    
Task_Look_at_Email:
    Dim TargetURL As String
    TargetURL = "https://outlook.office.com/mail/"
    Call OpenChrome(TargetURL)
    Application.Wait (Now + TimeValue("00:00:10"))
    
    For Repeat = 1 To 5
        Application.SendKeys "{Down}"
        
        Randomize
        Dim Delay As Integer
        Delay = 5 + Int((20 - 5 + 1) * Rnd)
        
        Application.Wait (Now + TimeValue("00:00:" & Format(Delay, "00")))
    Next Repeat
    
    For Repeat = 1 To 3
        Application.SendKeys "^w"
        Application.Wait (Now + TimeValue("00:00:02"))
    Next Repeat
    
    GoTo select_Task
    
Task_Load_Drawings:
    Dim PMmasterLocation As String
    PMmasterLocation = "\\server2\Dropbox\Jeremy Abbott\PM assistant (Master).xlsm"
    
    ' Implement drawing loading logic
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim fileLocations() As String
    Dim i As Long
    Dim Vpath As String
    Dim Xpath As String
    Dim randomSubfolder As String
    Dim pdfFiles As Collection
    Dim subfolders As Collection
    Dim depth As Integer
    Dim maxDepth As Integer
    Dim fileName As String
    Dim pdfFound As Boolean

    ' Open the workbook as read-only
    Set wb = Workbooks.Open(fileName:=PMmasterLocation, ReadOnly:=True)

    ' Assuming the relevant data is in the first worksheet
    Set ws = wb.Sheets(1)

    ' Find the last row in column B starting from B21
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Resize the array to hold the file locations
    ReDim fileLocations(1 To lastRow - 20)

    ' Loop through cells B21 to lastRow and store file locations in the array
    For i = 21 To lastRow
        fileLocations(i - 20) = ws.Cells(i, 2).Value
    Next i

    ' Close the workbook
    wb.Close SaveChanges:=False

    ' Randomly select one of the locations from the array
    Randomize ' Seed the random number generator
    Vpath = fileLocations(Int((UBound(fileLocations) - LBound(fileLocations) + 1) * Rnd + LBound(fileLocations)))

    ' Output the selected file path
    'MsgBox "In drawing load branch, selected pathway is: " & Vpath

    Dim subfolderArray As Variant
    subfolderArray = Array("Drawings", "RFI's")
    depth = 0 ' Start depth tracking
    maxDepth = 2 ' Set maximum depth for folder navigation

    ' Randomly select either the Drawings or RFI folder to start
    randomIndex = Int((UBound(subfolderArray) - LBound(subfolderArray) + 1) * Rnd + LBound(subfolderArray))
    currentFolder = subfolderArray(randomIndex)

    Do
        ' Initialize collections
        Set pdfFiles = New Collection
        Set subfolders = New Collection

        ' Check for PDF files in the current folder
        fileName = Dir(Vpath & "\" & currentFolder & "\*.pdf")
        
        ' Collect all PDF files in the current folder
        Do While fileName <> ""
            pdfFiles.Add Vpath & "\" & currentFolder & "\" & fileName
            fileName = Dir
        Loop

        ' Check for subfolders in the current folder
        Dim subfolderName As String
        subfolderName = Dir(Vpath & "\" & currentFolder & "\*", vbDirectory)
        
        Do While subfolderName <> ""
            If subfolderName <> "." And subfolderName <> ".." Then
                If (GetAttr(Vpath & "\" & currentFolder & "\" & subfolderName) And vbDirectory) <> 0 Then
                    subfolders.Add Vpath & "\" & currentFolder & "\" & subfolderName
                End If
            End If
            subfolderName = Dir
        Loop

        ' Randomly decide whether to open a PDF or dive into a subfolder
        If pdfFiles.Count > 0 Then
            If Rnd < 0.5 Then
                ' Open a random PDF file from the collected PDFs
                randomIndex = Int((pdfFiles.Count) * Rnd) + 1
                Xpath = pdfFiles(randomIndex)
                'MsgBox "Selected PDF file: " & Xpath
                ' Open the PDF
                Shell "explorer.exe """ & Xpath & """", vbNormalFocus
                ' Wait for 2 minutes
                Application.Wait (Now + TimeValue("00:02:00"))
                ' Close the PDF (use Alt+F4)
                SendKeys "%{F4}" ' Alt + F4 to close
                Exit Do ' Exit the loop after opening the PDF
            Else
                ' Dive deeper into subfolders if within depth limit
                If depth < maxDepth Then
                    Randomize ' Re-seed random number generator
                    randomIndex = Int((subfolders.Count) * Rnd) + 1
                    currentFolder = subfolders(randomIndex) ' Update to the selected subfolder
                    depth = depth + 1 ' Increase depth
                Else
                    'MsgBox "Reached maximum depth. Terminating process."
                    tries = tries + 1
                    If tries > 2 Then
                        tries = 0
                        Exit Do
                    End If
                    GoTo Task_Load_Drawings:
                End If
            End If
        Else
            'MsgBox "No PDF files found in " & currentFolder & ". Terminating process."
            tries = tries + 1
            If tries > 2 Then
                tries = 0
                Exit Do
            End If
            GoTo Task_Load_Drawings:
        End If
    Loop

    
    GoTo select_Task
    
    
Task_Load_Committed_Cost_Report:
    ' Implement committed cost report loading logic
    
    GoTo select_Task
    

End Sub
