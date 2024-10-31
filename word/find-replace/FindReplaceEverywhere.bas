Sub FindAndReplaceEverywhere()
    Dim folderPath As String
    Dim file As String
    Dim doc As Document
    Dim currentDocPath As String
    Dim totalReplacements As Long
    Dim docReplacements As Long
    Dim docsWithReplacements As Long
    Dim totalDocs As Long
    Dim oldText As String
    Dim newText As String
    Dim response As VbMsgBoxResult
    Dim pathSeparator As String
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Determine path separator based on OS
    #If Mac Then
        pathSeparator = "/"
    #Else
        pathSeparator = "\"
    #End If
    
    Application.ScreenUpdating = False
    
    ' Prompt user for the text to be replaced and the replacement text
    oldText = InputBox("Enter the text you want to replace:")
    If oldText = "" Then
        MsgBox "You must enter text to replace.", vbExclamation
        Exit Sub
    End If
    newText = InputBox("Enter the replacement text:")
    
    ' Prompt user to enter folder path or leave blank to use the current document's folder
    folderPath = InputBox("Enter the folder path or leave blank to use the current document's folder:")
    
    ' If no path entered, use the folder of the currently open document
    If folderPath = "" Then
        If Not ActiveDocument Is Nothing Then
            currentDocPath = ActiveDocument.Path
            If currentDocPath = "" Then
                MsgBox "The current document has not been saved. Please save it or enter a path.", vbExclamation
                Exit Sub
            Else
                folderPath = currentDocPath
            End If
        Else
            MsgBox "No document is currently open.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Clean up the path separators based on OS
    folderPath = Replace(folderPath, "\", pathSeparator)
    folderPath = Replace(folderPath, "/", pathSeparator)
    
    ' Basic folder existence check using Dir
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "The specified folder does not exist or is not accessible: " & folderPath, vbExclamation
        Exit Sub
    End If
    
    ' Ensure folder path ends with the correct separator
    If Right(folderPath, 1) <> pathSeparator Then
        folderPath = folderPath & pathSeparator
    End If
    
    ' Debug message to show the folder path being searched
    Debug.Print "Searching in folder: " & folderPath
    
    ' Check for both .doc and .docx files
    Dim foundFiles As String
    Dim fileList() As String
    Dim fileCount As Long
    fileCount = 0
    foundFiles = ""
    
    ' First check .docx files
    file = Dir(folderPath & "*.docx")
    Do While file <> ""
        If Left(file, 2) <> "~$" Then  ' Skip temporary files
            foundFiles = foundFiles & file & vbCrLf
            fileCount = fileCount + 1
            ReDim Preserve fileList(1 To fileCount)
            fileList(fileCount) = file
        End If
        file = Dir
    Loop
    
    ' Then check .doc files
    file = Dir(folderPath & "*.doc")
    Do While file <> ""
        If Left(file, 2) <> "~$" Then  ' Skip temporary files
            foundFiles = foundFiles & file & vbCrLf
            fileCount = fileCount + 1
            ReDim Preserve fileList(1 To fileCount)
            fileList(fileCount) = file
        End If
        file = Dir
    Loop
    
    ' If no files found, show debug information
    If fileCount = 0 Then
        MsgBox "No Word documents found in: " & folderPath & vbCrLf & _
               "Please verify that:" & vbCrLf & _
               "1. The path is correct" & vbCrLf & _
               "2. The documents have .doc or .docx extensions" & vbCrLf & _
               "3. You have permission to access this folder", vbExclamation
        Exit Sub
    End If
    
    ' Ask for confirmation before proceeding
    response = MsgBox("This will search and replace text in all Word documents in:" & vbCrLf & _
                     folderPath & vbCrLf & vbCrLf & _
                     "Replace '" & oldText & "' with '" & newText & "'" & vbCrLf & vbCrLf & _
                     "Documents found (" & fileCount & "):" & vbCrLf & foundFiles & vbCrLf & _
                     "Do you want to continue?", _
                     vbQuestion + vbYesNo, "Confirm")
    
    If response = vbNo Then Exit Sub
    
    ' Initialize counters
    totalReplacements = 0
    docsWithReplacements = 0
    totalDocs = 0
    
    ' Process all found files
    Dim i As Long
    For i = 1 To fileCount
        Debug.Print "Processing file: " & folderPath & fileList(i)
        
        ' Try to open each document with error handling
        On Error Resume Next
        Set doc = Documents.Open(folderPath & fileList(i), ReadOnly:=False, Visible:=False)
        
        If Err.Number <> 0 Then
            Debug.Print "Error opening file: " & fileList(i) & " - " & Err.Description
            MsgBox "Could not open file: " & fileList(i) & vbCrLf & _
                   "Error: " & Err.Description, vbExclamation
            Err.Clear
        Else
            On Error GoTo ErrorHandler
            totalDocs = totalDocs + 1
            docReplacements = 0
            
            ' Run the updates: Replace specified oldText with newText
            With doc.Content.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = oldText
                .Replacement.Text = newText
                .Forward = True
                .Wrap = wdFindContinue
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .Format = False
                
                ' Execute replacement all at once and count occurrences
                docReplacements = .Execute(Replace:=wdReplaceAll)
            End With
            
            ' Track replacements made in this document
            If docReplacements > 0 Then
                docsWithReplacements = docsWithReplacements + 1
                totalReplacements = totalReplacements + docReplacements
                
                ' Try to save with error handling
                On Error Resume Next
                doc.Save
                If Err.Number <> 0 Then
                    MsgBox "Could not save changes to: " & fileList(i) & vbCrLf & _
                           "Error: " & Err.Description, vbExclamation
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End If
            
            doc.Close
        End If
        
        ' Reset any error state before continuing loop
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' Summary message
    MsgBox "Update complete." & vbCrLf & _
           "Total documents processed: " & totalDocs & vbCrLf & _
           "Documents with replacements: " & docsWithReplacements & vbCrLf & _
           "Total replacements made: " & totalReplacements, vbInformation
           
ExitSub:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: " & Err.Source, vbCritical
    Resume ExitSub
End Sub
