# VBA - learn & archive

## Insert a row, relocate old values, get a clean new row
- Insert a row in the active cell`s line
- Move the old value(s) up by one cell

<div align="center">
    <img src="docs/insert_row.png" </img> 
</div>

```
Sub InsertRow_Active_Cell()
    
    ''' INSERT A NEW ROW
    ActiveCell.EntireRow.Insert
    
    ''' STARS - RELOCATE OLD VALUE
    Cell_Star_New = "G" & ActiveCell.Row + 1
    
    Range(Cell_Star_New).Offset(-1, 0) = Range(Cell_Star_New).Value
    Range(Cell_Star_New).Value = None
    
    ''' DATE - RELOCATE OLD VALUES
    New_Row_Value = ActiveCell.Row + 1
    Date_New_Range = "K" & New_Row_Value & ":" & "M" & New_Row_Value
    
    Range(Date_New_Range).Offset(-1, 0) = Range(Date_New_Range).Value
    Range(Date_New_Range).Value = None
    

End Sub
```



