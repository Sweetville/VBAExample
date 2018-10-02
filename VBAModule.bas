Attribute VB_Name = "Module1"
Sub CompareUni()
    Dim sText As String
    Dim sSText As String
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iSRow As Integer
    Dim iSColumn As Integer
    Dim sCPaste1 As String
    Dim sCPaste2 As String
    Dim iCPaste1 As Integer
    Dim iCPaste2 As Integer
    Dim iCColumn1 As Integer
    Dim iCColumn2 As Integer
    
    
    
    iColumn = 2
    iSColumn = 6
    iCPaste1 = 4
    iCPaste2 = 5
    iCColumn1 = 9
    iCColumn2 = 11
    
    
    For iRow = 1 To 2206
         If IsEmpty(Cells(iRow, iColumn)) = False Then
             sText = Cells(iRow, iColumn).Value
             
             For iSRow = 1 To 3040
                sSText = Cells(iSRow, iSColumn).Value
                If sSText = sText Then
                    Cells(iRow, iColumn).Interior.ColorIndex = 3
                    'sCPaste1 = Cells(iSRow, iCColumn1).Value
                    'Cells(iRow, iCPaste1).Value = sCPaste1
                    Cells(iRow, iCPaste1).Value = Cells(iSRow, iCColumn1).Value
                                        'Cells(iSRow, iCPaste1).Value = 3
                    Cells(iRow, iCPaste2).Value = Cells(iSRow, iCColumn2).Value
                    
                
                Else
                End If
                
             Next
        Else
        End If
        
    Next
        
    


End Sub
Sub FindDupes()
    Dim sText As String
    Dim sSText As String
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iSRow As Integer
    Dim iSColumn As Integer
    
    iColumn = 2
    
    For iRow = 1 To 2206
        
            If IsEmpty(Cells(iRow, iColumn)) = False Then
                sText = Cells(iRow, iColumn).Value
                
                    For iSRow = (iRow + 1) To 2206
                    
                        sSText = Cells(iSRow, iColumn).Value
                        If sSText = sText Then
                            'Cells(iSRow, iColumn).ClearContents
                            Cells(iSRow, iColumn).EntireRow.Interior.ColorIndex = 4
                        Else
                        End If
                    
                    Next
                
            Else
            End If
            
        Next
    
End Sub
Sub ColorDupesO()
    Dim sText As String
    Dim sSText As String
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iSRow As Integer
    Dim iSColumn As Integer
    
    iColumn = 4
    
    
    For iRow = 1 To 2206
        
            If IsEmpty(Cells(iRow, iColumn)) = False Then
                sText = Cells(iRow, iColumn).Value
                    
                If (IsEmpty(Cells(iRow, 8)) = False) Then
                    If (IsEmpty(Cells(iRow, 9)) = False) Then
                        For iSRow = (iRow + 1) To 2206
                        
                            sSText = Cells(iSRow, iColumn).Value
                            If sSText = sText Then
                                'Cells(iSRow, iColumn).ClearContents
                                Cells(iSRow, iColumn).EntireRow.Interior.ColorIndex = 4
                            Else
                            End If
                        
                        Next
                    Else
                    End If
                    
                Else
                End If
                    
            Else
            End If
            
        Next
    
End Sub

Sub you()


Dim row As Integer
Dim column As Integer
Dim text As String
Dim testing As String

text = "& r"
    
    column = 8
    For row = 1 To 2206
       testing = Cells(row, column).Value
        
        If InStr(testing, text) > 0 Then
            Cells(row, column).EntireRow.Interior.ColorIndex = 8
        Else
        End If
    Next
End Sub

Sub count()
    

Dim row As Integer
Dim column As Integer
Dim total As Integer

    total = 0
    
    column = 8
    For row = 1 To 2206
       
        
        If Cells(row, column).Interior.ColorIndex = 4 Then
            total = total + 1
            Cells(row, column + 4).Value = 3333
        Else
        End If
    Next
    
    Cells(1, 1).Value = total
End Sub

Sub deleteRow()


Dim row As Integer
Dim column As Integer
Dim text As String

text = "xxx"
column = 11

For row = 1 To 2206
    
        If Cells(row, column).Value = text Then
            Rows(row).Delete
        Else
        End If
        
        Next
        

    

End Sub
Sub orderNumbers()

Dim row As Integer
Dim column As Integer
Dim number As Integer

column = 1
number = 1

For row = 1 To 2206
    
        If (Cells(row, column + 1).Interior.ColorIndex <> 4) And (Cells(row, column + 3).Value <> 0) Then
            Cells(row, column).Value = number
            number = number + 1
        Else
        Cells(row, column).Clear
        End If
Next
End Sub
Sub movePlaces()
Dim row As Integer
Dim rowChange As Integer
Dim column As Integer
Dim moveCol As Integer
Dim movedCol As Integer

column = 2
moveCol = 14
movedCol = 8

For row = 1 To 2206
    
    For rowChange = 1 To 1400
        
        If Cells(row, column).Value = Cells(rowChange, moveCol).Value Then
            Cells(row, movedColumn).Value = Cells(rowChange, moveCol - 1).Value
            Cells(row, movedColumn + 1).Value = Cells(rowChange, moveCol).Value
            Cells(row, movedColumn + 2).Value = Cells(rowChange, moveCol + 1).Value
            
        Else
        End If
    Next

Next



End Sub
Sub Copy_Ten()
Dim X As Long
Dim LastRow As Long
Dim CopyRange As Range

LastRow = Cells(Cells.Rows.count, "B").End(xlUp).row
For X = 2 To LastRow Step 2
    If CopyRange Is Nothing Then
        Set CopyRange = Rows(X).EntireRow
    Else
        Set CopyRange = Union(CopyRange, Rows(X).EntireRow)
    End If
Next
If Not CopyRange Is Nothing Then
CopyRange.Copy Destination:=Sheets("Sheet4").Range("A2")
End If
End Sub
Sub reorderNum()
    Dim sText As String
    Dim sSText As String
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iSRow As Integer
    Dim numColumn As Integer
    Dim iCount As Integer
    Dim iCountPlace As Integer

    
    
    'declare var
    iColumn = 5
    numColumn = 2
    iCount = 1
    iCountPlace = 1

    'begin rotation
    For iRow = 2 To 2471
    
        'count the number in that one spot
        If Cells(iRow, iColumn).Interior.ColorIndex = 9 Then
            iCountPlace = 1
        Else
        End If
        
        
         If IsEmpty(Cells(iRow, iColumn)) = False Then
             sText = Cells(iRow, iColumn).Value
             
             'cycle through everything before it to check if value exists
             For iSRow = (iRow - 1) To 2 Step -1
                sSText = Cells(iSRow, iColumn).Value
                'if the same name is there, get rid of the number becaues it already exists elsewhere
                If sSText = sText Then
                    
                    Cells(iRow, numColumn).ClearContents
                    'break the for loop
                    iSRow = 2
                    iCount = iCount - 1
                    'Cells(iRow, iColumn).Interior.ColorIndex = 3

                Else
                'otherwise paste the number and increase count
                    Cells(iRow, numColumn).Value = iCount

                    
                End If
                

                
             Next
            iCount = iCount + 1
        Else
        End If
        If IsEmpty(Cells(iRow, numColumn)) = False Then
            Cells(iRow, 1).Value = iCountPlace
            iCountPlace = iCountPlace + 1
        Else
        End If
        
        
        
    Next
        
    


End Sub
Sub FindDupesa()
    Dim sText As String
    Dim sSText As String
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iSRow As Integer
    Dim iSColumn As Integer
    
    iColumn = 2
    
    For iRow = 1 To 2206
        
            If IsEmpty(Cells(iRow, iColumn)) = False Then
                sText = Cells(iRow, iColumn).Value
                
                    For iSRow = (iRow + 1) To 2206
                    
                        sSText = Cells(iSRow, iColumn).Value
                        If sSText = sText Then
                            'Cells(iSRow, iColumn).ClearContents
                            Cells(iSRow, iColumn).EntireRow.Interior.ColorIndex = 4
                        Else
                        End If
                    
                    Next
                
            Else
            End If
            
        Next
    
End Sub
Sub countKanji()
    

Dim row As Integer
Dim column As Integer
Dim total As Integer
Dim columnCheck As Integer


    total = 0
    
    column = 4
    columnCheck = 2
    For row = 2 To 3000
       
        
        If IsEmpty(Cells(row, column)) = False And IsEmpty(Cells(row, columnCheck)) = False Then
            total = total + 1
        ElseIf IsEmpty(Cells(row, column)) = False And IsEmpty(Cells(row, columnCheck)) = True Then
        
        Else
            Cells(row, 1).Value = total
            Cells(row, column).EntireRow.Interior.ColorIndex = 9
            total = 0
        End If
    Next
    
    Cells(1, 1).Value = total
End Sub
Sub pastereading()
    
Dim pRow As Integer
Dim pColumn As Integer
Dim row As Integer
Dim ncolumn As Integer
Dim srow As Integer
Dim scolumn As Integer
Dim carry As Integer



ncolumn = 1
scolumn = 3
'srow = row below
'reading is row above and column 3

pRow = 2
pColumn = 15

For row = 2 To 2500
          
        If IsEmpty(Cells(row, ncolumn)) = False And Cells(row, ncolumn) <> 0 Then
            
            Cells(pRow, pColumn).Value = Cells(row, ncolumn).Value
            Cells(pRow, pColumn - 1).Value = Cells(row - 1, scolumn).Value
            pRow = pRow + 1
                        
        Else

        End If
    Next




End Sub
Sub insertRows()
    
Dim pRow As Integer
Dim pColumn As Integer
Dim row As Integer
Dim column As Integer
Dim total As Integer
Dim carry As String



column = 3
pRow = 1
pColumn = 14
row = 2204
carry = "no"

While carry <> "hi"
        
        carry = Cells(row, column).Value
        If (Cells(row, column).Value <> carry) And (IsEmpty(Cells(row, column))) = False Then
            carry = Cells(row, column).Value
            Range(row).EntireRow.Insert
        Else
        End If
        
        row = row + 1
        
Wend

End Sub


Sub getMarker()

Dim cRow As Integer
Dim cColumn As Integer
Dim pRow As Integer
Dim pColumn As Integer
Dim pRowLoop As Integer
Dim Storage As String
Dim Check As Boolean


'Declare
cColumn = 2
pColumn = 6
pRow = 65


'loop the column if word
For cRow = 2 To 2500

    Check = True
    
    'if a word; then copy both that and left
    If IsEmpty(Cells(cRow, cColumn)) = False Then
        Storage = Cells(cRow, cColumn - 1).Value
    
    
        For pRowLoop = 2 To 214
            If Storage = Cells(pRowLoop, pColumn - 1) Then
                
                Check = False
                
            Else
            End If
            

        Next
        
        
                'paste to next paste row and increment
        If Check = True Then
        
            Cells(pRow, pColumn) = Cells(cRow, cColumn)
            Cells(pRow, pColumn - 1) = Cells(cRow, cColumn - 1)
                     
            pRow = pRow + 1
                     
        Else
        
        End If
    Else
    End If
        

    
    
 
    
    
Next
Sub testing()

End Sub
Sub generateRoomText()

Dim cRow As Integer
Dim cColumn As Integer
Dim pRow As Integer
Dim pColumn As Integer
Dim Storage As String


'Declare varialbels for the start of the list you want to copy
Dim CSTART As Integer
Dim CEND As Integer
Dim PSTART As Integer
Dim KanjiColumn As Integer


'give values
CSTART = 1498
CEND = 1509
cColumn = 18

pRow = 2022
pColumn = 8
KanjiColumn = 4
Storage = "hi"


'loop the list to copy from
For cRow = CSTART To CEND
    ' test Cells(1420, cColumn).Value = Cells(1421, cColumn).Value
    
    Storage = Cells(cRow, cColumn).Value
    
    'while the kanji column is not empty
    While IsEmpty(Cells(pRow, KanjiColumn)) = False
        'paste storage variable
        Cells(pRow, pColumn).Value = Storage
        'Paste one below and
        pRow = pRow + 1
    Wend
    'increment
    pRow = pRow + 1
        
Next




End Sub
Sub test()
Cells(1, 1).Value = 500
End Sub

Sub CreateReviewWithIncorrect()

Dim IncStart As Integer
Dim IncEnd As Integer
Dim Increment As Integer
Dim LoopR As Integer
Dim checkX As String
Dim PasteR As Integer

Dim readC As Integer
Dim kanjiC As Integer
Dim numC As Integer
Dim countR As Integer
Dim fixBlank As Integer




Dim mainworkBook As Workbook
Set mainworkBook = ActiveWorkbook


IncStart = 2361
IncEnd = 2460
PasteR = 2
readC = 6
kanjiC = 4
numC = 2
countR = 0
fixBlank = 1



'loop all from start
For LoopR = IncStart To IncEnd

    'if the line was incorrect or ensecure with an x
    If Cells(LoopR, 1).Interior.Color = vbYellow Then
        
        'grab line above and paste it elswhere
        
        'fix it so it wont grab a row that has no number
        While IsEmpty(Cells(LoopR - fixBlank, numC)) Or Cells(LoopR + fixBlank, numC).Interior.Color = vbYellow
            fixBlank = fixBlank + 1
        Wend
        
        mainworkBook.Sheets("Sheet1").Rows(LoopR - fixBlank).EntireRow.Copy
        mainworkBook.Sheets("Sheet2").Select
        
        mainworkBook.Sheets("Sheet2").Range("A1").Select
        mainworkBook.Sheets("Sheet2").Paste
        
        If Cells(PasteR - 1, 4).Value <> Cells(1, kanjiC).Value Then
            Cells(PasteR, 2).Value = Cells(1, numC).Value
            Cells(PasteR, 3).Value = Cells(1, readC).Value
            Cells(PasteR, 4).Value = Cells(1, kanjiC).Value
            Cells(PasteR, 5).Value = Cells(1, kanjiC).Value
            PasteR = PasteR + 1
        
        Else
        End If
        
        
        mainworkBook.Sheets("Sheet1").Select
        
   'MIDDDDDLEEEEE//////////////////////////////////////////////
            
            'while it does
            While Cells(LoopR, 1).Interior.Color = vbYellow
                'grab it, paste it, and move to the next one
        mainworkBook.Sheets("Sheet1").Rows(LoopR).EntireRow.Copy
        mainworkBook.Sheets("Sheet2").Select
        
        mainworkBook.Sheets("Sheet2").Range("A1").Select
        mainworkBook.Sheets("Sheet2").Paste
        
        Cells(PasteR, 2).Value = Cells(1, numC).Value
        Cells(PasteR, 3).Value = Cells(1, readC).Value
        Cells(PasteR, 4).Value = Cells(1, kanjiC).Value
       ' Cells(PasteR, 2).Value = Cells(1, kanjiC).Value
       
       'count it
       countR = countR + 1
        
        
        mainworkBook.Sheets("Sheet1").Select
                'increment so it can check it again
                LoopR = LoopR + 1
                PasteR = PasteR + 1
                
            Wend
       'otherwise do nothing
        
        
                
    'BOTTTOMMMMMMMMMM////////////////////////////////////////
    'grab line below all that
    
     'fix it so it wont grab a row that has no number
        'make fixblank 0
        fixBlank = 0
        
        While IsEmpty(Cells(LoopR + fixBlank, numC)) Or Cells(LoopR + fixBlank, numC).Interior.Color = vbYellow
            fixBlank = fixBlank + 1
        Wend
    
        mainworkBook.Sheets("Sheet1").Rows(LoopR + fixBlank).EntireRow.Copy
        mainworkBook.Sheets("Sheet2").Select
        
        mainworkBook.Sheets("Sheet2").Range("A1").Select
        mainworkBook.Sheets("Sheet2").Paste
        
        Cells(PasteR, 2).Value = Cells(1, numC).Value
        Cells(PasteR, 3).Value = Cells(1, readC).Value
        Cells(PasteR, 4).Value = Cells(1, kanjiC).Value
        Cells(PasteR, 5).Value = Cells(1, kanjiC).Value
        
        
        mainworkBook.Sheets("Sheet1").Select
        PasteR = PasteR + 1
    
            'reset fixblank
        fixBlank = 1
    Else
    End If

Next
mainworkBook.Sheets("Sheet2").Select
Cells(1, 1).Value = countR

End Sub
Sub CreateReviewWithIncorrectPRUNE()

Dim IncStart As Integer
Dim IncEnd As Integer
Dim Increment As Integer
Dim LoopR As Integer
Dim checkX As String
Dim PasteR As Integer

Dim readC As Integer
Dim kanjiC As Integer
Dim numC As Integer
Dim countR As Integer
Dim fixBlank As Integer
Dim LoopU As Integer
Dim LoopD As Integer


'INITIALIZE
Dim mainworkBook As Workbook
Set mainworkBook = ActiveWorkbook

'Pos var
IncStart = 2
IncEnd = 2473


'Count Var
PasteR = 3
countR = 0
fixBlank = 1

'Column Var
readC = 6
kanjiC = 4
numC = 2


'loop all from start
For LoopR = IncStart To IncEnd

    'If yellow
    If Cells(LoopR, 1).Interior.Color = vbYellow Then
    

'P1///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        LoopU = LoopR - 1
        'Before yellow?
        
        'Find A valid number entry from above
        While IsEmpty(Cells(LoopU, numC))
            LoopU = LoopU - 1
        Wend
        
        'If it isn't yellow, we do stuff
        If Cells(LoopU, 1).Interior.Color <> vbYellow Then
        
                'Grab and Paste with Answer
            mainworkBook.Sheets("Sheet1").Rows(LoopU).EntireRow.Copy
            mainworkBook.Sheets("Sheet2").Select
            
            mainworkBook.Sheets("Sheet2").Range("A1").Select
            mainworkBook.Sheets("Sheet2").Paste
    
                'Already grabbed?
            If Cells(PasteR - 1, 4).Value <> Cells(1, kanjiC).Value Then
                Cells(PasteR, 2).Value = Cells(1, numC).Value
                Cells(PasteR, 3).Value = Cells(1, readC).Value
                Cells(PasteR, 4).Value = Cells(1, kanjiC).Value
                Cells(PasteR, 5).Value = Cells(1, kanjiC).Value
                PasteR = PasteR + 1
            
            Else
            End If
            
            'Reselect main page
            mainworkBook.Sheets("Sheet1").Select
            
            
        Else
        End If
'P2///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        'Grab and paste line without Answer
        mainworkBook.Sheets("Sheet1").Rows(LoopR).EntireRow.Copy
        mainworkBook.Sheets("Sheet2").Select
        
        mainworkBook.Sheets("Sheet2").Range("A1").Select
        mainworkBook.Sheets("Sheet2").Paste
        
        Cells(PasteR, 2).Value = Cells(1, numC).Value
        Cells(PasteR, 3).Value = Cells(1, readC).Value
        Cells(PasteR, 4).Value = Cells(1, kanjiC).Value
       
       'count it
        countR = countR + 1
        
        'Reselect main Page
        mainworkBook.Sheets("Sheet1").Select
        
        'increment Pasting line
        PasteR = PasteR + 1
'P3///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        LoopD = LoopR + 1
        'After yellow?
        
        'Find A valid number entry from below
        While IsEmpty(Cells(LoopD, numC))
            LoopD = LoopD + 1
        Wend
        
        'If it isn't yellow, we do stuff
        If Cells(LoopD, 1).Interior.Color <> vbYellow Then
        
            'Grab with answer
            mainworkBook.Sheets("Sheet1").Rows(LoopD).EntireRow.Copy
            mainworkBook.Sheets("Sheet2").Select
            
            mainworkBook.Sheets("Sheet2").Range("A1").Select
            mainworkBook.Sheets("Sheet2").Paste
            
            Cells(PasteR, 2).Value = Cells(1, numC).Value
            Cells(PasteR, 3).Value = Cells(1, readC).Value
            Cells(PasteR, 4).Value = Cells(1, kanjiC).Value
            Cells(PasteR, 5).Value = Cells(1, kanjiC).Value
            PasteR = PasteR + 1
            
        Else
        End If
        
        'Reselect main page
        mainworkBook.Sheets("Sheet1").Select
        
    Else
    End If
 
Next

mainworkBook.Sheets("Sheet2").Select
Cells(1, 1).Value = countR

End Sub

Sub reformatMessed()

Dim cRow As Integer
Dim cCol As Integer
Dim pRow As Integer
Dim pCol As Integer
Dim Start As Integer
Dim eend As Integer
Dim ColSize As Integer
Dim RowSize As Integer
Dim count As Integer


ColSize = 30
RowSize = 2

pCol = 3
pRow = 1

cCol = 1

Start = 31
eend = 929
counter = 3

For cRow = Start To eend
    If pRow = 31 Then
    
        pRow = 1
        counter = counter + RowSize
    Else
    End If
    pCol = counter
    
        For cCol = 1 To RowSize
        
            Cells(pRow, pCol) = Cells(cRow, cCol).Value
    
            pCol = pCol + 1
        Next
    
    pRow = pRow + 1

Next

End Sub

Sub testingoo()
Cells(1, 1).Interior.ColorIndex = 9
End Sub
