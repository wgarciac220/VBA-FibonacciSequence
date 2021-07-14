Sub prueba()
    Debug.Print fib(50) ' prints 12586269025
    Debug.Print fib(70) ' prints 190392490709135
End Sub

Function fib(num, Optional memo As Scripting.Dictionary) As LongLong

    If memo Is Nothing Then Set memo = New Scripting.Dictionary

    If memo.Exists(num) Then
        fib = memo(num)
        GoTo EXIT_HERE
    End If

    If num <= 2 Then
        fib = 1
        Exit Function
    End If
    
    memo.Add num, fib(num - 1, memo) + fib(num - 2, memo)
    
EXIT_HERE:
    fib = memo(num)
    
End Function
