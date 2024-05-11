Attribute VB_Name = "Module_Distances"
Option Explicit

Public Function HammingDistance(str1 As String, str2 As String) As Long

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim i As Integer, Distance As Integer
    
    If Len(str1) <> Len(str2) Then
        HammingDistance = -1
        Exit Function
    End If
    
    Distance = 0
    For i = 1 To Len(str1)
        If Mid(str1, i, 1) <> Mid(str2, i, 1) Then
            Distance = Distance + 1
        End If
    Next i
    
    HammingDistance = Distance
    
End Function



Public Function EditDistance(Seq1 As String, Seq2 As String) As Integer
    
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel

    Dim i As Integer, j As Integer, Penalty As Integer
    Dim Distance() As Long
    Dim Min1 As Long, Min2 As Long, Min3 As Long
    
    If Len(Seq1) = 0 Then
        EditDistance = Len(Seq2)
        Exit Function
    End If
    
    If Len(Seq2) = 0 Then
        EditDistance = Len(Seq1)
        Exit Function
    End If
    
    ReDim Distance(Len(Seq1), Len(Seq2))
    
    For i = 0 To Len(Seq1)
        Distance(i, 0) = i
    Next
    
    For j = 0 To Len(Seq2)
        Distance(0, j) = j
    Next
    
    For i = 1 To Len(Seq1)
        For j = 1 To Len(Seq2)
            If Mid(Seq1, i, 1) = Mid(Seq2, j, 1) Then
                Penalty = 0
            Else
                Penalty = 1
            End If
            
            Min1 = (Distance(i - 1, j) + 1)
            Min2 = (Distance(i, j - 1) + 1)
            Min3 = (Distance(i - 1, j - 1) + Penalty)
            
            Distance(i, j) = Application.WorksheetFunction.Min(Min1, Min2, Min3)
            
        Next
    Next
    
    EditDistance = Distance(Len(Seq1), Len(Seq2))

End Function
