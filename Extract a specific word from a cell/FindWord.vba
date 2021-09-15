Function FindWord(Source As String, Position As Integer) As String
     On Error Resume Next
     FindWord = Split(WorksheetFunction.Trim(Source), " ")(Position - 1)
     On Error GoTo 
End Function

Function FindWordRev(Source As String, Position As Integer) As String
     Dim Arr() As String
     Arr = VBA.Split(WorksheetFunction.Trim(Source), " ")
     On Error Resume Next
     FindWordRev = Arr(UBound(Arr) - Position + 1)
     On Error GoTo 
End Function