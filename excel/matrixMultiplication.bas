'Still working on this

Sub MatrixMultiplication()
Dim rng1, rng2  As Range
Dim cellValue As Integer
Dim colValues(), rowValues(), cellProduct() As Variant


    Set rng1 = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
    Set rng2 = Application.InputBox("Select a range", "Obtain Range Object", Type:=8)
    
    For Each Row In rng1.Rows
        Row.Select
        For Each RowCell In Row
            RowCell.Select
            For Each Column In rng2.Columns
                For Each ColCell In Column
                    For i = 1 To rng1.Columns.Count
                        rowValues(i) = RowCell.Value
                    Next i
                    For j = 1 To rng2.Rows.Count
                        colValues(j) = ColCell.Value
                    Next j
                    For k = 1 To i
                        For L = 1 To j
                            For m = 1 To j
                                cellProduct(m) = rowValues(k) * colValues(L)
                            Next m
                        Next L
                    Next k
                    For Each Number In cellProduct
                        Total = Total + Val(Number)
                    Next Number
                    RowCell.Offset(0, 10).Value = Total
                Next ColCell
            Next Column
        Next RowCell
    Next Row
            
    
    
    
    
    
    'MsgBox "You have selected an " & rng1.Rows.Count & " by " & rng1.Columns.Count & " matrix."

    

End Sub
