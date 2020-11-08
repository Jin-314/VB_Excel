Attribute VB_Name = "Module1"
Function Kinji(num As Integer, x As Range, y As Range)
    Dim sum As Long
    Dim tmpX As Long
    Dim Array1(num, num) As Long
    Dim Array2(0, num) As Long
    sum = 0
    For i = 0 To num
        For j = 0 To num
            For Each tmpX In x
                tmp = tmp ^ ((4 - i) - j)
                sum = sum + tmp
            Next x
        Array1(i, j) = sum
        Next j
    Next i
    sum = 0
    tmp = 0
    For i = 0 To num
        For j = 0 To x.Columns.Count
            tmp = x.Item(0, j) ^ (num - i) * y.Item(0, j)
            sum = sum + tmp
        Next j
    Next i
    
    
End Function
