Attribute VB_Name = "FileInfo"

Public Sub getNewFilesFromFolder(folderPath As String, sheetname As String, startRow As Integer, rowIncrement As Integer, startColumn As Integer, Optional mustContainString As String)

    ' Lists all the new files within a folder
    ' Outputs the file info to the given sheet starting at the given row and column
    ' If the file has already been listed on the sheet, it isn't listed again
    ' Column outputs are as follows
    ' Include? / Name / Path / Link

    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim i As Integer
    
    ' Check that the folder path was entered
    If folderPath = "" Then
        MsgBox ("You must enter a destination file path")
        Exit Sub
    End If
    
    ' Make sure the last character is a "/" or "\"
    If Not right(folderPath, 1) = "/" And Not right(folderPath, 1) = "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Check that the folder path is vaild
    If Not FolderExists(folderPath) Then
        MsgBox ("The specified folder path does not exist")
        Exit Sub
    End If
    
    ' Check that the sheet exists
    If Not SheetExists(sheetname) Then
        MsgBox ("The specified sheet does not exist")
        Exit Sub
    End If
    
    ' Check that the start row and column makes sense
    If startRow <= 0 Then
        MsgBox ("The start row must be greater than 0")
        Exit Sub
    End If
    If startColumn <= 0 Then
        MsgBox ("The start column must be greater than 0")
        Exit Sub
    End If
    
    'Create an instance of the FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Get the folder object
    Set objFolder = objFSO.GetFolder(folderPath)
    
    ' Set the titles
    Sheets(sheetname).Cells(startRow, startColumn) = "Include?"
    Sheets(sheetname).Cells(startRow, startColumn + 1) = "Name"
    Sheets(sheetname).Cells(startRow, startColumn + 2) = "Path"
    Sheets(sheetname).Cells(startRow, startColumn + 3) = "Link"
    
    firstEntryRow = startRow + 1
    j = startColumn
    
    ' Get next available row (so that we don't overwrite what's already there)
    For i = firstEntryRow To 10000 Step rowIncrement
        ' If the name column is empty, assume this is the next available row
        If Sheets(sheetname).Cells(i, startColumn + 1) = "" Then Exit For
    Next i
    
    ' If no string entered that the name must contain, then use a wildcard
    If mustContainString = "" Then mustContainSring = "*"
    
    'loops through each file in the directory and prints their names and path
    For Each objFile In objFolder.Files
    
        ' Only output the file if it includes the required string
        fileName = objFile.Name
        If fileName Like mustContainString Then
            
            ' Check if the file has been recorded yet
            alreadyListed = False
            For checkRow = firstEntryRow To 10000 Step rowIncrement
                checkedName = Sheets(sheetname).Cells(checkRow, startColumn + 1)
                If checkedName = fileName Then
                    alreadyListed = True
                    Exit For
                End If
            Next checkRow
            
            ' If no matches were found, record it
            If Not alreadyListed Then
                ' Put a border over the and under the row
                With Cells(i, startColumn).Rows.EntireRow
                         With .Borders(xlEdgeTop)
                             .LineStyle = xlContinuous
                             .ColorIndex = 1
                         End With
                End With
                With Cells(i + rowIncrement - 1, startColumn).Rows.EntireRow
                         With .Borders(xlEdgeBottom)
                             .LineStyle = xlContinuous
                             .ColorIndex = 1
                         End With
                End With
                'whether to include the file or not (1/0 - adjusted by user, default 1)
                Cells(i, startColumn) = 1
                'print file name
                Cells(i, startColumn + 1) = objFile.Name
                'print file path
                Cells(i, startColumn + 2) = objFile.path
                'add a link
                Cells(i, startColumn + 3).Formula = "=HYPERLINK(""" & objFile.path & """,""" & "Open" & """)"
                ' Instead of just incrementing i, set it to the next available row
                ' This allows for files to be deleted in random places and listed in the correct place when this is next run
                ' Get next available row (so that we don't overwrite what's already there)
                For i = firstEntryRow To 10000 Step rowIncrement
                    ' If the name column is empty, assume this is the next available row
                    If Sheets(sheetname).Cells(i, startColumn + 1) = "" Then Exit For
                Next i
            End If
            
        End If
        
    Next objFile
    
    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing


End Sub

Public Function GetDirectory(path)
   GetDirectory = left(path, InStrRev(path, "\"))
End Function

Public Function IsPDF(path) As Boolean
    If GetExtension(path) = "pdf" Then
        IsPDF = True
    Else
        IsPDF = False
    End If
End Function

Public Function GetExtension(path)
   GetExtension = right(path, Len(path) - InStrRev(path, "."))
End Function

Public Function RemoveExtension(path)
   RemoveExtension = left(path, InStrRev(path, ".") - 1)
End Function


Public Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
    
End Function

Public Function FolderExists(strFolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

Public Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Public Sub DeleteFile(ByVal FileToDelete As String)
    If FileToDelete <> "" Then
        If FileExists(FileToDelete) Then 'See above
           SetAttr FileToDelete, vbNormal
           Kill FileToDelete
        End If
    End If
End Sub
