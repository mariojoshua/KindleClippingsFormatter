# How to Download and process your notes from kindle

1. Go to Kindle in File Explorer > Documents > MyClippings > Clippings.txt and copy the file to your local disk
2. Open Excel > Data Tab > Get External Data > From Text > Select Clippings.txt. > Import data > Clippings.txt
3. You will notice each quote and the included metadata occupying exactly 5 rows with a blank line. 
![Screenshot 2022-03-23 at 5 44 52 PM](https://user-images.githubusercontent.com/75271182/159696894-2787877e-2b3c-43db-8afd-493d5cf97e41.png)
4. Run the below VBA code by going to Developer Tab > Visual Basic > Insert Module > Paste the below code there and press the save button.

```vb
Sub Trans5()
Dim rng As Range
Dim I As Long
    
    Set rng = Range("A1")
    While rng.Value <> ""
        I = I + 1
        rng.Resize(5).Copy
        Range("B" & I).PasteSpecial Transpose:=True
        Set rng = rng.Offset(5)
        
    Wend
    rng.EntireColumn.Delete
End Sub
```

5. Go to Macros and Press Run, All the columns should get sorted, add appropriate naming and delete empty columns.
6.Viola! You have you data and can filter it based on book. 

P.S.
If Developer Tab is not enabled, go to preference/settings > ribbon & toolbar > Main Tabs > Check the Developer column and save.

