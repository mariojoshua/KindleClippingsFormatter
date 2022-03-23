# How to Process your notes from kindle using Excel

1. Go to Kindle in File Explorer > Documents > MyClippings > Clippings.txt and copy the file to your local disk
2. Open Excel > Data Tab > Get External Data > From Text > Select Clippings.txt. > Import data > Clippings.txt
3. You will notice each quote and the included metadata occupying exactly 5 rows with a blank line. 

![Screenshot 2022-03-23 at 5 44 52 PM](https://user-images.githubusercontent.com/75271182/159696894-2787877e-2b3c-43db-8afd-493d5cf97e41.png)

4. Run the below VBA code by going to Developer Tab > Visual Basic > Insert Module > Paste the below code there and press the save button.

```vb
Sub KindleClippingsFormatter()
Static rngA As Range
Static I As Long
Static lookVal As String
    
    'Init starting I to 1
    I = 1
    'Init starting range to ("A2")
    Set rngA = Range("A" & I)
    'WHILE rangeA is not equal to blank
    Do
    'While Range("A" & I).Value <> ""
            lookVal = Range("A" & I).Value
            Set rngA = Range("A" & 1 + I)
            'IF range A value is blank, then exit
           If rngA.Value = "" Then
               Exit Do
           End If
            
        'IF RANGEA contains entry Then
        If InStr(rngA.Value, "Entry") > 0 Then
             I = I + 1
             'Go to next row
             'Set rngA = Range("A" & I)
             rngA.Select
             
        Else
            'If rngA.Value <> "" Then
                'Exit Do
            'IF Range "B"I is empty Then
            If IsEmpty(Range("B" & I)) Then
                'Paste in Range "B"I
                rngA.Copy Range("B" & I)
             'ELSEIF Range "C"I is empty Then
            ElseIf IsEmpty(Range("C" & I)) Then
                'Paste in Range "C"I
                rngA.Copy Range("C" & I)
            'ElseIf IsEmpty(Range("D" & I)) Then
            Else
                'Paste in Range "D"I
                rngA.Copy Range("D" & I)
            End If
            'Delete RangeA
            Range("A" & I + 1).EntireRow.Delete
            
        End If
    
        Set rngA = Range("A" & I)
         rngA.Select
    'Wend
    Loop
    
    'Copy A1 header to B1,C1,D1
    Range("A1").Copy Range("B1")
    Range("A1").Copy Range("C1")
    Range("A1").Copy Range("D1")
    
   Set rngA = Range("E:E")
   For J = 1 To 4
        Set rngA = rngA.Offset(, -1)
        
        'Insert 9 Lines AfterColumn
        For I = 1 To 8
          rngA.Offset(, 1).EntireColumn.Insert
        Next I
          
        'Range D Text to Columns
        rngA.TextToColumns _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=True, _
        Comma:=True
   Next J
  
   'Fit Columns to size
    Cells.EntireColumn.AutoFit
    
    'Delete UnwantedColumns
    Range("AG:AG").EntireColumn.Delete
    Range("AB:AD").EntireColumn.Delete
    Range("X:X").EntireColumn.Delete
    Range("S:U").EntireColumn.Delete
    Range("O:O").EntireColumn.Delete
    Range("J:L").EntireColumn.Delete
    Range("F:F").EntireColumn.Delete
    Range("A:C").EntireColumn.Delete
    
End Sub
```

5. Go to Macros and Press Run, All the columns should get sorted, add appropriate naming and delete empty columns.
6. Voila! You have you data and can filter it based on book. 

![Screenshot 2022-03-23 at 5 54 21 PM](https://user-images.githubusercontent.com/75271182/159698556-a871c2d8-c205-46ea-b26d-5c042b3a56ff.png)

P.S.
If Developer Tab is not enabled, go to preference/settings > ribbon & toolbar > Main Tabs > Check the Developer column and save.

