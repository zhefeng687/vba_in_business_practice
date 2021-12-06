Attribute VB_Name = "Bubble_Sort"
Option Explicit

Public ds As Object
Public arr As Variant
Public reports As String

' construct a business process
Sub sorting_demo()

Call data_acquisition
Call Bubble_Sort
Call reporting

End Sub

' data acquisition
' Ref: https://excelmacromastery.com/vba-arraylist/
' Declare and Create ArrayList
Sub data_acquisition()
    Dim item As Variant
    Set ds = CreateObject("System.Collections.ArrayList")
        ' add items
        ds.Add 91
        ds.Add 81
        ds.Add 71
        ds.Add 61
        ds.Add 51
    arr = ds.ToArray    
End Sub
	
' core business: sorting
' a) # of iterations
' b) # of comparision in each iteration
Sub Bubble_Sort()
    Dim size As Integer
    Dim i As Integer
    Dim j As Integer
    size = UBound(arr) - LBound(arr) + 1
    reports = ""
    reports = reports & reporting_msg("Orig:  ", arr)
    
    '# of iterations: size -1 = 4
    For i = 0 To size - 2
    '# of comparision in each iteration: size -1 = 4
        For j = 0 To size - 2
            'abc swap
            If arr(j) > arr(j + 1) Then
                Call swap_method(arr, j, j + 1)
            End If
        Next
    ' reporting the result of each iteration
      reports = reports & reporting_msg("Iter" & i & ": ", arr)
    Next
End Sub

'original: 91, 81, 71, 61, 51
'iter 0: 81, 71, 61, 51, 91
'iter 1: 71, 61, 51, 81, 91
'iter 2: 61, 51, 71, 81, 91
'iter 3: 51, 61, 71, 81, 91
'Fin: 51, 61, 71, 81, 91
Sub reporting()
'Only needs to report the final result
    reports = reports & reporting_msg("Fin:   ", arr)
    Debug.Print reports
End Sub


' swapping elements
' subroutines is used when a desired task is needed but without a returning value
Sub swap_method(data As Variant, pos_1 As Integer, pos_2 As Integer)
        Dim temp As Integer
        temp = data(pos_1)
        data(pos_1) = data(pos_2)
        data(pos_2) = temp
End Sub


' reporting function
Function reporting_msg(pre_fix As String, data As Variant)
    Dim msg As String
    Dim n As Integer   
    msg = pre_fix
    For n = 0 To UBound(arr)
        msg = msg & data(n) & ","
    Next
    reporting_msg = Left(msg, Len(msg) - 1) & vbNewLine
End Function


