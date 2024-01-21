Attribute VB_Name = "variant_dynamic_array"
Sub variant_test()

Dim test() As Variant
ReDim test(3)

test = Array("Germany population in million: ", 83, Date)

Dim i As Integer


For i = 0 To 2
    
    Debug.Print "Element " & i & " of test array is: " & test(i) & " of type " & TypeName(test(i))

Next i


End Sub
