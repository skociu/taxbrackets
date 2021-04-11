#  Calculating Federal Tax Owed in Excel
The following is just for illustration purposes in calculating taxes and is not to be taken as real investment advice. 

## Function used to perform the calculations
To perform the calculations, I've created a custom defined function called CalcTax which takes as input two variables, the taxable amount and the tax table

```
Public Function CalcTax(Amount As Range, TaxTable As Range) As Single
Dim n As Integer
Dim m As Integer
Dim cnt As Integer
Dim MyArray As Variant
MyArray = TaxTable.Value

'Find the location row of the Amount in the TaxTable - variable cnt
For n = 1 To UBound(MyArray)
    If Amount > MyArray(n, 1) Then
        cnt = n
    End If
Next n

'Calculate the Tax using the TaxTable
For m = 1 To cnt
    If m < cnt Then
        CalcTax = CalcTax + ((MyArray(m, 2) - MyArray(m, 1)) * MyArray(m, 3))
    ElseIf m = cnt Then
        CalcTax = CalcTax + (Amount - MyArray(m, 1)) * MyArray(m, 3)
    End If
Next m

End Function

```
