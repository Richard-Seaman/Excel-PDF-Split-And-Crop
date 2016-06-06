Attribute VB_Name = "BasicSheet"
Private Const titleRow As Integer = 11
Private Const firstEntryRow As Integer = 12
Private Const rowsBetweenFiles As Integer = 5
Private Const includeColumn As Integer = 1
Private Const nameColumn As Integer = 2
Private Const pathColumn As Integer = 3
Private Const linkColumn As Integer = 4
Private Const firstQuestionColumn As Integer = 5
Private Const canOpenString = "Openable"
Private Const canNotOpenString = "Not Openable"
Private Const notPDFString = "Not a pdf..."
Private Const pageExtractionType = "P"
Private Const cropExtractionType = "C"
Private Const comboExtractionType = "PC"
Private Const cropLeft As Integer = 0           ' for A4
Private Const cropRight As Integer = 600       ' more than for A4



Sub getFilesForThisSheet()

    Dim sourceFolder As String
    Dim sheetname As String
    Dim fileNameMustContainString As String

    sourceFolder = Cells(1, 3)
    sheetname = ActiveSheet.Name
    fileNameMustContainString = Cells(2, 3)
    
    Call getNewFilesFromFolder(sourceFolder, sheetname, firstEntryRow - 1, rowsBetweenFiles, includeColumn, fileNameMustContainString)
    
End Sub

Sub checkIfPDFsCanOpen()
   
    Dim sheetname As String
    Dim sourcePdfFullPath As String
    
    Dim numberOfPagesAtStart As Integer
    Dim numberOfPagesAtEnd As Integer
    
    sheetname = ActiveSheet.Name
           
    ' Cycle through each file and remove the pages
    For i = firstEntryRow To 10000 Step rowsBetweenFiles
    
        include = Sheets(sheetname).Cells(i, 1)
        fileName = Sheets(sheetname).Cells(i, 2)
        sourcePdfFullPath = Sheets(sheetname).Cells(i, 3)
        
        result = ""
        
        ' Don't continue if no more files
        If fileName = "" Then Exit For
        
        ' Only continue if file is a pdf
        If GetExtension(sourcePdfFullPath) = "pdf" Then
            ' Check if Adobe can open the document
            If canPDFOpen(sourcePdfFullPath) Then
                result = canOpenString
            Else
                result = canNotOpenString
            End If
        Else
            result = notPDFString
        End If
        
        ' Output the result under the name
        Cells(i + 1, nameColumn) = result
        
    Next i
    

End Sub

Sub extractQuestions()

    Dim destFolder As String
    Dim sheetname As String
    Dim sourcePdfFullPath As String
    Dim destPdfFullPath As String
    
    Dim firstInput As Integer
    Dim secondInput As Integer
    Dim thirdInput As Integer
    Dim fourthInput As Integer
    
    
    ' Confrim call
    If MsgBox("Extract Questions?", vbYesNo, "Confirm") = vbNo Then End
        
    ' Make sure the inputs are okay
    Call checkInputs
        
    ' Grab destination folder
    destFolder = Cells(3, 3)
    
    If destFolder = "" Then
        MsgBox ("Must enter a destination folder.")
        Exit Sub
    End If
    
    ' Make sure the last character is a "/" or "\"
    If Not right(destFolder, 1) = "/" And Not right(destFolder, 1) = "\" Then
        destFolder = destFolder & "\"
    End If
    
    ' Make sure the destination folder exists
    If Not FolderExists(destFolder) Then
        MsgBox ("Destination folder doesn't exist.")
        Exit Sub
    End If
    
    ' TODO: apply the correct method
    For i = firstEntryRow To 10000 Step rowsBetweenFiles
    
        include = Cells(i, includeColumn)
        fileName = Cells(i, nameColumn)
        sourcePdfFullPath = Cells(i, pathColumn)
        
        ' Exit if no more
        If fileName = "" Then Exit For
        
        ' Only continue if included
        If include = 1 Then
        
            ' Cycle through each of the question columns
            For j = firstQuestionColumn To 10000
            
                ' Don't continue if no more questions to check
                questionTitle = Cells(titleRow, j)
                If questionTitle = "" Then Exit For
                
                ' Only extract if question is present in this file
                ' (if the first entry isn't blank)
                If Cells(i + 1, j) <> "" Then
                
                    ' Grab the inputs (and convert to Integers)
                    extractionType = Cells(i, j)
                    firstInput = Int(Cells(i + 1, j))
                    secondInput = Int(Cells(i + 2, j))
                    thirdInput = Int(Cells(i + 3, j))
                    fourthInput = Int(Cells(i + 4, j))
                    
                    ' Construct the destination full path
                    destPdfFullPath = destFolder & RemoveExtension(fileName) & "-" & questionTitle & ".pdf"
                     
                    If extractionType = comboExtractionType Then
                    
                        ' Combo Extraction
                        Call extractCombo(sourcePdfFullPath, destPdfFullPath, firstInput, secondInput, thirdInput, fourthInput, cropLeft, cropRight)
                    
                    
                    ElseIf extractionType = cropExtractionType Then
                    
                        ' Crop Extraction
                        Call extractCrop(sourcePdfFullPath, destPdfFullPath, firstInput, thirdInput, secondInput, cropLeft, cropRight)
                    
                    Else
                    
                        ' Page Extraction
                        Call extractPages(sourcePdfFullPath, destPdfFullPath, firstInput, secondInput)
                                        
                    End If
                
                End If
            
            Next j
        
        End If
        
    Next i
    
    ' Automatically open the destination folder
    ActiveWorkbook.FollowHyperlink destFolder
    

End Sub


Private Sub checkInputs()

    For i = firstEntryRow To 10000 Step rowsBetweenFiles
    
        include = Cells(i, includeColumn)
        fileName = Cells(i, nameColumn)
        openable = Cells(i + 1, nameColumn)
        
        ' Exit if no more
        If fileName = "" Then Exit For
        
        ' Only check if included
        If include = 1 Then
        
            ' Make sure we've checked if it's openable
            If openable <> canOpenString Then
                MsgBox (fileName & " at row " & i & " is not openable or has not been checked." & vbNewLine & "Use the Check If Openable button or fix the problem (or don't include it).")
                End
            End If
            
            ' Cycle through each of the question columns
            For j = firstQuestionColumn To 10000
                
                ' Don't continue if no more questions to check
                questionTitle = Cells(titleRow, j)
                If questionTitle = "" Then Exit For
            
                extractionType = Cells(i, j)
                firstInput = Cells(i + 1, j)
                secondInput = Cells(i + 2, j)
                thirdInput = Cells(i + 3, j)
                fourthInput = Cells(i + 4, j)
                
                ' Only check inputs if this question is applied to this file
                If extractionType <> "" Then
                
                    ' Check that one of the three methods has been entered
                    If extractionType = pageExtractionType Or extractionType = cropExtractionType Or extractionType = comboExtractionType Then
                    
                        ' Run checks on inputs
                        If extractionType = comboExtractionType Then
                        
                            ' 1 - startPage
                            ' 2 = top crop position (on start page)
                            ' 3 = end page
                            ' 4 = bottom crop position (on end page)
                        
                            ' Check combo inputs
                            If firstInput = "" Or secondInput = "" Or thirdInput = "" Or fourthInput = "" Then
                                ' No inputs
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                                End
                            End If
                            
                            If Not IsNumeric(firstInput) Or Not IsNumeric(secondInput) Or Not IsNumeric(thirdInput) Or Not IsNumeric(fourthInput) Then
                                ' Non numbers entered
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                                End
                            End If
                            
                            If firstInput < 0 Or secondInput < 0 Or thirdInput <= firstInput Or fourthInput < 0 Then
                                ' Inputs must be >= 0
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Input limits exceeded (check for 0's)")
                                End
                            End If
                        
                        ElseIf extractionType = cropExtractionType Then
                        
                            ' Check crop inputs
                            If firstInput = "" Or secondInput = "" Or thirdInput = "" Then
                                ' No inputs
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                                End
                            End If
                            
                            If Not IsNumeric(firstInput) Or Not IsNumeric(secondInput) Or Not IsNumeric(thirdInput) Then
                                ' Non numbers enetered
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                                End
                            End If
                            
                            If firstInput < 0 Or secondInput < 0 Or Int(thirdInput) <= 0 Then
                                ' Inputs must be >= 0
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Input limits exceeded (check for 0's)")
                                End
                            End If
                            
                            
                        Else
                        
                            ' Check page inputs
                            If firstInput = "" Or secondInput = "" Then
                                ' No inputs
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                                End
                            End If
                            
                            If Not IsNumeric(firstInput) Or Not IsNumeric(secondInput) Then
                                ' Non numbers enetered
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                                End
                            End If
                            
                            If firstInput < 0 Or Int(secondInput) <= 0 Then
                                ' Inputs must be >= 0
                                MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Input limits exceeded (check for 0's)")
                                End
                            End If
                        
                        End If
                    
                    Else
                        ' Invalid extraction method entered
                        MsgBox ("Invalid extraction method for " & fileName & " at row " & i & vbNewLine & "Must be either " & pageExtractionType & " or " & cropExtractionType)
                        End
                    End If
                
                End If
            
            Next j
        
        End If
    
    Next i

End Sub

Sub TestExtract()

    Dim destFolder As String
    Dim sheetname As String
    Dim sourcePdfFullPath As String
    Dim destPdfFullPath As String
    
    Dim firstInput As Integer
    Dim secondInput As Integer
    Dim thirdInput As Integer
    Dim fourthInput As Integer
        
    ' Make sure the inputs are okay
    Call checkInputs
        
    ' Grab destination folder
    destFolder = "C:\Users\Rich\Desktop\TestPdfs\"
        
    ' Make sure the destination folder exists
    If Not FolderExists(destFolder) Then
        MsgBox ("The temp pdf folder doesn't exist, see hardcoded VBA value" & vbNewLine & destFolder)
        End
    End If

    rowToUse = Cells(5, 5)
    columnToUse = Cells(6, 5)
    
    ' Check the row / column inputs
    If rowToUse = "" Or columnToUse = "" Then
        ' No inputs
        MsgBox ("Must enter a row and column to test.")
        End
    End If
    
    If Not IsNumeric(rowToUse) Or Not IsNumeric(columnToUse) Then
        ' Non numbers enetered
        MsgBox ("Row and Column must be numbers")
        End
    End If
    
    If rowToUse < titleRow Or columnToUse < firstQuestionColumn Then
        ' Must be in the question range
        MsgBox ("Invalid row / column entered. Make sure it's within the question range.")
        End
    End If
    
    include = Cells(rowToUse, includeColumn)
    fileName = Cells(rowToUse, nameColumn)
    openable = Cells(rowToUse + 1, nameColumn)
    sourcePdfFullPath = Cells(rowToUse, pathColumn)
    
    ' Exit if no more
    If fileName = "" Then
        MsgBox ("No file in row specified")
        End
    End If
    
    extractionType = Cells(rowToUse, columnToUse)
    firstIn = Cells(rowToUse + 1, columnToUse)      ' don't assign to firstInput yet as it's defined as an Intger and it will cause type mismatch in checks below
    secondIn = Cells(rowToUse + 2, columnToUse)
    thirdIn = Cells(rowToUse + 3, columnToUse)
    fourthIn = Cells(rowToUse + 4, columnToUse)
    
    ' Check inputs for test
    If extractionType <> "" Then
    
        ' Check that one of the three methods has been entered
        If extractionType = pageExtractionType Or extractionType = cropExtractionType Or extractionType = comboExtractionType Then
        
            ' Run checks on inputs
            If extractionType = comboExtractionType Then
            
                ' 1 - startPage
                ' 2 = top crop position (on start page)
                ' 3 = end page
                ' 4 = bottom crop position (on end page)
            
                ' Check combo inputs
                If firstIn = "" Or secondIn = "" Or thirdIn = "" Or fourthIn = "" Then
                    ' No inputs
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                    End
                End If
                
                If Not IsNumeric(firstIn) Or Not IsNumeric(secondIn) Or Not IsNumeric(thirdIn) Or Not IsNumeric(fourthIn) Then
                    ' Non numbers entered
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                    End
                End If
                
                If firstIn < 0 Or secondIn < 0 Or thirdIn <= firstIn Or fourthIn < 0 Then
                    ' Inputs must be >= 0
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Input limits exceeded (check for 0's)")
                    End
                End If
            
            ElseIf extractionType = cropExtractionType Then
            
                ' Check crop inputs
                If firstIn = "" Or secondIn = "" Or thirdIn = "" Then
                    ' No inputs
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                    End
                End If
                
                If Not IsNumeric(firstIn) Or Not IsNumeric(secondIn) Or Not IsNumeric(thirdIn) Then
                    ' Non numbers enetered
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                    End
                End If
                
                If firstIn < 0 Or secondIn < 0 Or Int(thirdIn) <= 0 Then
                    ' Inputs must be >= 0
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Input limits exceeded (check for 0's)")
                    End
                End If
                
                
            Else
            
                ' Check page inputs
                If firstIn = "" Or secondIn = "" Then
                    ' No inputs
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                    End
                End If
                
                If Not IsNumeric(firstIn) Or Not IsNumeric(secondIn) Then
                    ' Non numbers enetered
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Must enter a number")
                    End
                End If
                
                If firstIn < 0 Or Int(secondIn) <= 0 Then
                    ' Inputs must be >= 0
                    MsgBox ("QuestionError for " & fileName & " at row " & i & vbNewLine & "Input limits exceeded (check for 0's)")
                    End
                End If
            
            End If
        
        Else
            ' Invalid extraction method entered
            MsgBox ("Invalid extraction method for " & fileName & " at row " & i & vbNewLine & "Must be either " & pageExtractionType & " or " & cropExtractionType)
            End
        End If
    
    Else
    
        MsgBox ("No Extraction Type sepcified for test")
        End
        
    End If
           
    ' Set the integers
    firstInput = Int(firstIn)
    secondInput = Int(secondIn)
    thirdInput = Int(thirdIn)
    fourthInput = Int(fourthIn)
    
    ' Construct the destination full path
    destPdfFullPath = getTestMergeFilePath(destFolder, 0)
         
    If extractionType = comboExtractionType Then
                    
        ' Combo Extraction
        Call extractCombo(sourcePdfFullPath, destPdfFullPath, firstInput, secondInput, thirdInput, fourthInput, cropLeft, cropRight)
        
    ElseIf extractionType = cropExtractionType Then
                    
        ' Crop Extraction
        Call extractCrop(sourcePdfFullPath, destPdfFullPath, firstInput, thirdInput, secondInput, cropLeft, cropRight)
    
    Else
    
        ' Page Extraction
        Call extractPages(sourcePdfFullPath, destPdfFullPath, firstInput, secondInput)
                        
    End If
    
    ' Automatically open the new pdf
    ActiveWorkbook.FollowHyperlink destPdfFullPath
     

End Sub

Sub includeAll()

    ' Set all file's Include to 1 on current sheet

    ' Confrim call
    If MsgBox("This will set include to 1 for all files. Continue?", vbYesNo, "Confirm") = vbNo Then End
     
    For Row = firstEntryRow To 10000 Step rowsBetweenFiles
        
        If Cells(Row, 2) = "" Then Exit For
        
        Cells(Row, 1) = 1
        
    Next Row

End Sub

Sub includeNone()

    ' Set all file's Include to 0 on current sheet

    ' Confrim call
    If MsgBox("This will set include to 0 for all files. Continue?", vbYesNo, "Confirm") = vbNo Then End
     
    For Row = firstEntryRow To 10000 Step rowsBetweenFiles
        
        If Cells(Row, 2) = "" Then Exit For
        
        Cells(Row, 1) = 0
        
    Next Row

End Sub

Sub insertExtraRow()

    ' This sub was used to increase the number of rows between files on sheets that were already populated with data
    ' So that they were the same as new sheets (increased from 4 to 5 to accomodate new extraction type)
    
    currentRowsBetweenFiles = 4
    
    ' Make sure that the current number of rows between files isn't the standard
    ' If this is the case, we don't want to insert any
    If rowsBetweenFiles = currentRowsBetweenFiles Then
        MsgBox ("Must have different number of rows between files on current sheet than the standard. See VB Code")
        End
    End If
    
    startRow = 12
    
    ' Confrim call
    If MsgBox("Insert an extra row for each file?" & vbNewLine & "StartRow = " & startRow & vbNewLine & "Only continue if CurrentRowsBetweenFiles = " & currentRowsBetweenFiles, vbYesNo, "Confirm") = vbNo Then End
     
    Offset = 0
    For Row = startRow + currentRowsBetweenFiles To 10000 Step currentRowsBetweenFiles
        
        If Cells(Row + Offset, 2) = "" Then Exit For
        
        Rows(Row + Offset).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Offset = Offset + 1
        
    Next Row

End Sub

Function getTestMergeFilePath(destFolder As String, number As Integer) As String

    ' Used to get the next available "TestMerged#.pdf" file name
    ' Because if we delete it and then save it closes the active pdf
    ' So just save it as a new name and then manually delete them all later
    
    defaultName = "TestExtract"
    path = destFolder & defaultName & number & ".pdf"
    
    ' Check if it exists
    With New FileSystemObject
        If .FileExists(path) Then
            ' If it does exist, try again with a recursive call
            path = getTestMergeFilePath(destFolder, number + 1)
        End If
    End With
    
    ' Return the path
    getTestMergeFilePath = path

End Function
