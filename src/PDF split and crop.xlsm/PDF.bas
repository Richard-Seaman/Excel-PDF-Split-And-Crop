Attribute VB_Name = "PDF"

Public Sub removePages(sourceFullPath As String, destFullPath As String, pagesToRemoveAtStart As Integer, pagesToRemoveAtEnd As Integer)
                       
    ' Check that the file paths were entered
    If sourceFullPath = "" Then
        MsgBox ("You must enter the source file path")
        Exit Sub
    End If
    If destFullPath = "" Then
        MsgBox ("You must enter a destination file path")
        Exit Sub
    End If
        
    ' Check that the folders exist
    If Not FolderExists(GetDirectory(sourceFullPath)) Then
        MsgBox ("The specified source file path does not exist" & vbNewLine & sourceFullPath)
        Exit Sub
    End If
    If Not FolderExists(GetDirectory(destFullPath)) Then
        MsgBox ("The specified destination file path does not exist" & vbNewLine & destFullPath)
        Exit Sub
    End If
        
    ' PDF EDIT BEGINS HERE
    
    Dim pdDoc As Acrobat.CAcroPDDoc, newPDF As Acrobat.CAcroPDDoc
    ' Dim pdPage As Acrobat.CAcroPDPage
    Dim PNum, PToRemove As Long
        
    ' Open the source PDF
    Set pdDoc = CreateObject("AcroExch.pdDoc")
    result = pdDoc.Open(sourceFullPath)
    If Not result Then
       MsgBox "Can't open file: " & sourceFullPath
       Exit Sub
    End If
    
    ' Check that the combined number of pages to remove does not exceed the total number of pages
    PNum = pdDoc.GetNumPages
    PToRemove = pagesToRemoveAtStart + pagesToRemoveAtEnd
    
    If PNum <= PToRemove Then
        MsgBox ("Trying to remove more pages than there are, aborting..." & vbNewLine & vbNewLine & "PDF: " & sourceFullPath & vbNewLine & "Actual # pages:" & PNum & vbNewLine & "# pages to remove: " & PToRemove)
        Exit Sub
    End If
    
    ' Create a new PDF
    Set newPDF = CreateObject("AcroExch.pdDoc")
    newPDF.Create
    newPDF.InsertPages -1, pdDoc, 0 + pagesToRemoveAtStart, PNum - pagesToRemoveAtStart - pagesToRemoveAtEnd, 0
    newPDF.Save 1, destFullPath
    newPDF.Close
    
    Set newPDF = Nothing
    Set pdDoc = Nothing
    
End Sub

Public Sub extractPages(sourceFullPath As String, destFullPath As String, startPage As Integer, numPagesToExtract As Integer)
                       
    ' Check that the file paths were entered
    If sourceFullPath = "" Then
        MsgBox ("You must enter the source file path")
        Exit Sub
    End If
    If destFullPath = "" Then
        MsgBox ("You must enter a destination file path")
        Exit Sub
    End If
        
    ' Check that the folders exist
    If Not FolderExists(GetDirectory(sourceFullPath)) Then
        MsgBox ("The specified source file path does not exist" & vbNewLine & sourceFullPath)
        Exit Sub
    End If
    If Not FolderExists(GetDirectory(destFullPath)) Then
        MsgBox ("The specified destination file path does not exist" & vbNewLine & destFullPath)
        Exit Sub
    End If
        
    ' PDF EDIT BEGINS HERE
    
    Dim pdDoc As Acrobat.CAcroPDDoc, newPDF As Acrobat.CAcroPDDoc
    Dim PNum, PToRemove As Long
        
    ' Open the source PDF
    Set pdDoc = CreateObject("AcroExch.pdDoc")
    result = pdDoc.Open(sourceFullPath)
    If Not result Then
       MsgBox "Can't open file: " & sourceFullPath
       Exit Sub
    End If
    
    ' Check that the combined number of pages to remove does not exceed the total number of pages
    PNum = pdDoc.GetNumPages
    
    If PNum <= numPagesToExtract Then
        MsgBox ("Trying to extract more pages than there are, aborting..." & vbNewLine & vbNewLine & "PDF: " & sourceFullPath & vbNewLine & "Actual # pages:" & PNum & vbNewLine & "# pages to remove: " & numPagesToExtract)
        Exit Sub
    End If
    
    ' Create a new PDF
    Set newPDF = CreateObject("AcroExch.pdDoc")
    newPDF.Create
    newPDF.InsertPages -1, pdDoc, startPage, numPagesToExtract, 0
    newPDF.Save 1, destFullPath
    newPDF.Close
    
    Set newPDF = Nothing
    Set pdDoc = Nothing
    
End Sub

Public Sub extractCrop(sourceFullPath As String, destFullPath As String, page As Integer, top As Integer, bottom As Integer, left As Integer, right As Integer)
    
    Dim pdDoc As Acrobat.CAcroPDDoc
    Dim pdPage As Acrobat.CAcroPDPage
    Dim pageRect As Object
    
    Set pdDoc = CreateObject("AcroExch.PDDoc")
        'Set pdPage = pdDoc(1)
     
    ' Create a copy of the source pdf incase the actual one is being viewed (which it probably will be)
    copyPath = RemoveExtension(sourceFullPath) & "Copy.pdf"
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(sourceFullPath, copyPath)
    
    ' Open the copy so we can crop it
    pdDoc.Open (copyPath)
    
    ' Grab the first page
    Set pdPage = pdDoc.AcquirePage(page)
    ' Grab the page rectangle so we can figure out the width
    Set pageRect = pdPage.GetSize
    
    Set rect = CreateObject("AcroExch.Rect")
        
    rect.top = top
    rect.left = left
    rect.bottom = bottom
    rect.right = pageRect.x
    
    ' First crop all the pages
    pdDoc.CropPages 0, pdDoc.GetNumPages, 0, rect
        
    ' Then delete all the pages except the one we want
    ' Remember, the position of the page we want changes when we start deleting pages
    totalNumberOfPages = pdDoc.GetNumPages
    For pdfPage = 0 To totalNumberOfPages - 1
    
        If pdfPage <> page Then
            ' Delete the page
            pdDoc.DeletePages pdfPage, pdfPage
            ' Correct our page reference
            If pdfPage < page Then
                page = page - 1
            End If
            ' Start at the beginning of the loop again (will add one at next)
            pdfPage = -1
            ' Recount the pages
            totalNumberOfPages = pdDoc.GetNumPages
            ' If we're at the last page, then quit
            If totalNumberOfPages = 1 Then Exit For
        End If
        
    Next pdfPage
    
    pdDoc.Save 1, destFullPath
     
    pdDoc.Close
    
    ' Delete the copy we created to edit
    fso.DeleteFile copyPath
    
    Set fso = Nothing
    
    Set pdDoc = Nothing
    Set rect = Nothing

End Sub

Private Sub testCombo()

    Dim sourceFullPath As String
    Dim destFullPath As String
    Dim startPage As Integer
    Dim top As Integer
    Dim endPage As Integer
    Dim bottom As Integer
    Dim left As Integer
    Dim right As Integer

    sourceFullPath = "C:\Users\Rich\Desktop\Accounting-2004-EV-OL.pdf"
    destFullPath = "C:\Users\Rich\Desktop\ComboExtractTest.pdf"
    startPage = 1
    top = 600
    endPage = 5
    bottom = 300
    left = 0
    right = 600
    
    Call extractCombo(sourceFullPath, destFullPath, startPage, top, endPage, bottom, left, right)
    

End Sub

Public Sub extractCombo(sourceFullPath As String, destFullPath As String, startPage As Integer, top As Integer, endPage As Integer, bottom As Integer, left As Integer, right As Integer)
    
    ' Combo extract involves cropping the first and last pages
    ' Extracting any pages between (if present)
    ' Combining the files into the final pdf
    
    ' Multiple steps involved:
    
    ' 1 - crop the start Page from the 'top' position to 0 (the bottom of the page)
    '   - save the cropped start page in a temp location
    
    ' 2 - crop the last Page from the 'bottom' position to 1000 (well above the top of th page - assuming A4)
    '   - save the cropped last page in a temp location
    
    ' 3 - Check if there's any files inbetween the start and end page
    '   - if there is, do a page extraction and save to a temp location
    
    ' 4 - combine the pages as follows:
    '   - startPage (cropped) -> middlePages (if present) -> lastPage (cropped)
    '   - save the combined pdf to the destination location
    
    ' 5 - Delete the temporary files created
    
    
    ' For some reason calling the PDF crop and ages extract subs overrides the variables entered into this sub
    ' So make a copy and hope it persists
    Dim comboStartPage As Integer
    Dim comboEndPage As Integer
    Dim comboTop As Integer
    Dim comboBottom As Integer
    
    comboStartPage = startPage
    comboEndPage = endPage
    comboTop = top
    comboBottom = bottom
    
    Dim copyPath As String
        
    Dim pdDocStart As Acrobat.CAcroPDDoc
    Dim pdDocEnd As Acrobat.CAcroPDDoc
    Dim pdDocMiddle As Acrobat.CAcroPDDoc
    
    Set pdDocStart = CreateObject("AcroExch.PDDoc")
    Set pdDocEnd = CreateObject("AcroExch.PDDoc")
    Set pdDocMiddle = CreateObject("AcroExch.PDDoc")
    
    Dim startTempPath As String
    Dim middleTempPath As String
    Dim endTempPath As String
    Dim startAndMiddleTempPath As String
    
    ' Delete the current file if present
    DeleteFile (destFullPath)
    
    ' Create a copy of the source pdf incase the actual one is being viewed (which it probably will be)
    copyPath = RemoveExtension(sourceFullPath) & "Copy.pdf"
    ' Delete the current copy if present
    DeleteFile (copyPath)
    ' then copy
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(sourceFullPath, copyPath)
    
    ' Get the folder path from the source path
    ' The temp files will be saved in the same folder
    tempFolder = GetDirectory(sourceFullPath)
    
    ' Create the three temp file paths
    startTempPath = tempFolder + "tempStart.pdf"
    middleTempPath = tempFolder + "tempMiddle.pdf"
    endTempPath = tempFolder + "tempEnd.pdf"
    startAndMiddleTempPath = tempFolder + "tempStartPlusMiddle.pdf"
    
    ' Delete the temps if present (although they shouldn't be)
    DeleteFile (startTempPath)
    DeleteFile (middleTempPath)
    DeleteFile (endTempPath)
    DeleteFile (startAndMiddleTempPath)
    
    ' 1 - crop the start Page from the 'top' position to 0 (the bottom of the page)
    '   - save the cropped start page in a temp location
    Call extractCrop(copyPath, startTempPath, startPage, top, 0, left, right)
    
    ' 2 - crop the last Page from the 'bottom' position to 1000 (well above the top of th page - assuming A4)
    '   - save the cropped last page in a temp location
    Call extractCrop(copyPath, endTempPath, endPage, 1000, bottom, left, right)
    
    ' 3 - Check if there's any files inbetween the start and end page
    '   - if there is, do a page extraction and save to a temp location
    middlePagesPresent = False
    If comboEndPage - comboStartPage > 1 Then
        middlePagesPresent = True
        Call extractPages(copyPath, middleTempPath, comboStartPage + 1, comboEndPage - comboStartPage - 1)
    End If
    
    ' 4 - combine the pages as follows:
    '   - startPage (cropped) -> middlePages (if present) -> lastPage (cropped)
    '   - save the combined pdf to the destination location
    
    If middlePagesPresent Then
        
        ' Combine first page and middle pages
        Call MergePDFs(startTempPath, middleTempPath, startAndMiddleTempPath)
        
        ' Then combine these with the end page - saving to the actual Extract Combo dest
        Call MergePDFs(startAndMiddleTempPath, endTempPath, destFullPath)
    
    Else
    
        ' Combine start and end pages, saving to the actual dest
        Call MergePDFs(startTempPath, endTempPath, destFullPath)
    
    End If
    
    
    ' 5 - Delete the temporary files created (if they exist)
    DeleteFile (startTempPath)
    DeleteFile (middleTempPath)
    DeleteFile (endTempPath)
    DeleteFile (startAndMiddleTempPath)
    DeleteFile (copyPath)
    
    
    
    ' Clean Up
    Set pdDocStart = Nothing
    Set pdDocEnd = Nothing
    Set pdDocMiddle = Nothing
    
    ' Automatically open the new pdf
    ' ActiveWorkbook.FollowHyperlink destFullPath

End Sub

Private Sub MergePDFs(firstFullPath As String, secondFullPath As String, destFullPath As String)

    ' This sub combines two pdfs together and saves it to a destination location
    ' the pdf at firstFullPath will be before the pdf at secondFullPath

    Dim Part1Document As Acrobat.CAcroPDDoc
    Dim Part2Document As Acrobat.CAcroPDDoc

    Dim numPages As Integer

    Set Part1Document = CreateObject("AcroExch.PDDoc")
    Set Part2Document = CreateObject("AcroExch.PDDoc")

    Part1Document.Open (firstFullPath)
    Part2Document.Open (secondFullPath)

    ' Insert the pages of Part2 after the end of Part1
    numPages = Part1Document.GetNumPages()

    If Part1Document.InsertPages(numPages - 1, Part2Document, 0, Part2Document.GetNumPages(), True) = False Then
        MsgBox "Cannot insert pages"
    End If

    If Part1Document.Save(PDSaveFull, destFullPath) = False Then
        MsgBox "Cannot save the modified document"
    End If

    Part1Document.Close
    Part2Document.Close

    Set Part1Document = Nothing
    Set Part2Document = Nothing

End Sub

Sub Test()

    Dim pdDoc As Acrobat.CAcroPDDoc
    
    Set pdDoc = CreateObject("AcroExch.PDDoc")
     
    pdDoc.Open ("C:\Users\Rich\Desktop\SEC Web Scraper\Material\LeavingCert\QuestionsRaw\Accounting-2001-EV-HL-Q7.pdf")
     
    Set rect = CreateObject("AcroExch.Rect")
     
    rect.top = 450
    rect.left = 0
    rect.bottom = 0
    rect.right = 600
    
    page = 6
    
    ' First crop all the pages
    pdDoc.CropPages 0, pdDoc.GetNumPages, 0, rect
    
    totalNumberOfPages = pdDoc.GetNumPages
    
    ' Then delete all the pages except the one we want
    ' Remember, the position of the page we want changes when we start deleting pages
    For pdfPage = 0 To totalNumberOfPages - 1
    
        If pdfPage <> page Then
            ' Delete the page
            pdDoc.DeletePages pdfPage, pdfPage
            ' Correct our page reference
            If pdfPage < page Then
                page = page - 1
            End If
            ' Start at the beginning of the loop again (will add one at next)
            pdfPage = -1
            ' Recount the pages
            totalNumberOfPages = pdDoc.GetNumPages
            ' If we're at the last page, then quit
            If totalNumberOfPages = 1 Then Exit For
        End If
        
    Next pdfPage
    
    pdDoc.Save 1, "C:\Users\Rich\Desktop\SEC Web Scraper\Material\LeavingCert\QuestionsRaw\test.pdf"
     
    pdDoc.Close
    Set pdDoc = Nothing
    Set rect = Nothing

End Sub

Public Function canPDFOpen(pdfFullPath As String) As Boolean
       
    Dim openable As Boolean
                       
    ' Checks that should immediately return false
    
    If pdfFullPath = "" Then
        canPDFOpen = False
        Exit Function
    End If
        
    If Not FolderExists(GetDirectory(pdfFullPath)) Then
        canPDFOpen = False
        Exit Function
    End If
    
    If Not IsPDF(pdfFullPath) Then
        canPDFOpen = False
        Exit Function
    End If
        
        
    ' Actual PDF Openable check starts here
    
    Dim pdDoc As Acrobat.CAcroPDDoc
        
    ' Open the source PDF
    Set pdDoc = CreateObject("AcroExch.pdDoc")
    result = pdDoc.Open(pdfFullPath)
    If Not result Then
        openable = False
    Else
        openable = True
    End If
    
    Set pdDoc = Nothing
    
    canPDFOpen = openable
    

End Function

Private Sub TestMerge()

    Dim AcroApp As Acrobat.CAcroApp

    Dim Part1Document As Acrobat.CAcroPDDoc
    Dim Part2Document As Acrobat.CAcroPDDoc

    Dim numPages As Integer

    Set AcroApp = CreateObject("AcroExch.App")

    Set Part1Document = CreateObject("AcroExch.PDDoc")
    Set Part2Document = CreateObject("AcroExch.PDDoc")

    Part1Document.Open ("C:\Users\Rich\Desktop\SEC Web Scraper\Material\LeavingCert\QuestionsFinal\Accounting-2001-EV-HL-Q8.pdf")
    Part2Document.Open ("C:\Users\Rich\Desktop\SEC Web Scraper\Material\LeavingCert\QuestionsFinal\Accounting-2001-EV-HL-Q7.pdf")

    ' Insert the pages of Part2 after the end of Part1
    numPages = Part1Document.GetNumPages()

    If Part1Document.InsertPages(numPages - 1, Part2Document, 0, Part2Document.GetNumPages(), True) = False Then
        MsgBox "Cannot insert pages"
    End If

    If Part1Document.Save(PDSaveFull, "C:\Users\Rich\Desktop\MergedFile.pdf") = False Then
        MsgBox "Cannot save the modified document"
    End If

    Part1Document.Close
    Part2Document.Close

    AcroApp.Exit
    Set AcroApp = Nothing
    Set Part1Document = Nothing
    Set Part2Document = Nothing

    MsgBox "Done"


End Sub

Public Sub AddText()

    Dim pdApp As Acrobat.AcroApp
    Dim pdDoc As Acrobat.AcroPDDoc
    ' Dim pdPage As Acrobat.AcroPDPage
    Dim jso As Object
    Dim doc As Variant
    
    Set pdApp = CreateObject("AcroExch.App")
    Set pdDoc = CreateObject("AcroExch.PDDoc")
    pdDoc.Open ("C:\Users\Rich\Desktop\Accounting-2001-EV-HL-Q7.pdf")
    Set jso = pdDoc.GetJSObject
    
    Dim textToAdd As String
    textToAdd = "2007"
    
    Call jso.addWaterMarkFromText(textToAdd, jso.app.Constants.Align.top, _
    jso.Font.Helv, 12, _
    jso.Color.black, 0, 0, True, True, True, _
    jso.app.Constants.Align.right, jso.app.Constants.Align.top, _
    -10, -10, False, 1, False, 0, 1)    ' from right, from top
  
    
    pdDoc.Save 1, "C:\Users\Rich\Desktop\Accounting-2001-EV-HL-Q7-edited.pdf"
    pdDoc.Close
    Set pdDoc = Nothing
    
    
End Sub

Private Sub TestSplit()

    Dim pdDoc As Acrobat.CAcroPDDoc, newPDF As Acrobat.CAcroPDDoc
    ' Dim pdPage As Acrobat.CAcroPDPage
    Dim thePDF As String, PNum As Long
    
    thePDF = "C:\Users\Rich\Desktop\Accounting-2004-EV-OL.pdf"
    
    '...
    Set pdDoc = CreateObject("AcroExch.pdDoc")
    result = pdDoc.Open(thePDF)
    If Not result Then
       MsgBox "Can't open file: " & fileName
       Exit Sub
    End If
    
    newPath = "C:\Users\Rich\Desktop\"
    
    '...
    PNum = pdDoc.GetNumPages
    
    For i = 0 To PNum - 1
        Set newPDF = CreateObject("AcroExch.pdDoc")
        newPDF.Create
        NewName = newPath & "Page_" & i & "_of_" & PNum & ".pdf"
        newPDF.InsertPages -1, pdDoc, i, 1, 0
        newPDF.Save 1, NewName
        newPDF.Close
        Set newPDF = Nothing
    Next i

End Sub
