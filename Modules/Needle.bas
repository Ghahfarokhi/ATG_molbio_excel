Attribute VB_Name = "Module_Needleman"
Option Explicit

Public Function NeedleAlignmnet(Text1 As String, Text2 As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    'Explicit definition of variables:
    Dim NeedleTable() As Integer, TraceBack() As String
    Dim i As Integer, j As Integer, Beside As Integer, Diag As Integer, Up As Integer
    Dim Additional As Integer, x As Integer, y As Integer, MatchScore As Long
    Dim Seq1 As String, Seq2 As String, CaseSelect As String
    Dim SeqAlign1 As String, SeqAlign2 As String
    
    'Remove the UCase to make it case sensitive.
    Seq1 = UCase(Text1)
    Seq2 = UCase(Text2)
    
    'Defining constants:
    Dim GAP As Integer, ExtGAP As Integer, Mismatch As Integer, Match As Integer
    GAP = Sheets("info").Range("Needle_Gap_Open").Value
    ExtGAP = Sheets("info").Range("Needle_Gap_Extend").Value
    Mismatch = Sheets("info").Range("Needle_Mismatch").Value
    Match = Sheets("info").Range("Needle_Match").Value
    

    'Defining the dimension of arrays:
    ReDim NeedleTable(Len(Seq2), Len(Seq1)) As Integer
    ReDim TraceBack(Len(Seq2), Len(Seq1))

    'Start the Needle table with 0 for the first element:
    NeedleTable(0, 0) = 0
    TraceBack(0, 0) = "Diag"

    For i = 1 To Len(Seq1)
        NeedleTable(0, i) = NeedleTable(0, i - 1) + ExtGAP
        TraceBack(0, i) = "Beside"
    Next i

    For j = 1 To Len(Seq2)
        NeedleTable(j, 0) = NeedleTable(j - 1, 0) + ExtGAP
        TraceBack(j, 0) = "Up"
    Next j
    
    For i = 1 To Len(Seq2)
        For j = 1 To Len(Seq1)
            If TraceBack(i, j - 1) = "Beside" Or TraceBack(i, j - 1) = "Up" Then
                Beside = NeedleTable(i, j - 1) + ExtGAP
                Up = NeedleTable(i - 1, j) + ExtGAP
            Else
                Beside = NeedleTable(i, j - 1) + GAP
                Up = NeedleTable(i - 1, j) + GAP
            End If
            
            'Diag = NeedleTable(i - 1, j - 1)
            If Mid(Seq2, i, 1) = Mid(Seq1, j, 1) Then
                Diag = NeedleTable(i - 1, j - 1) + Match
            Else
                Diag = NeedleTable(i - 1, j - 1) + Mismatch
            End If
    
            
            CaseSelect = Max1(Beside, Diag, Up)
            Select Case CaseSelect
                Case False
                  MsgBox "The Max function returned an error!"
                  Exit Function
                Case Else
                    If InStr(1, CaseSelect, "c") > 0 Then 'use "a" for right alignment
                        NeedleTable(i, j) = Up 'Beside 'for right alignment
                        TraceBack(i, j) = "Up" '"Beside" 'for right alignment
                    Else
                        If InStr(1, CaseSelect, "b") > 0 Then
                            NeedleTable(i, j) = Diag
                            TraceBack(i, j) = "Diag"
                        Else
                            NeedleTable(i, j) = Beside 'Up 'for right alignment
                            TraceBack(i, j) = "Beside" '"Up" 'for right alignment
                        End If
                    End If
                End Select
        Next j
    Next i
    SeqAlign1 = ""
    SeqAlign2 = ""

    On Error Resume Next

    Do While i > 0 And j > 0
        x = i - 1
        y = j - 1
        If TraceBack(i - 1, j - 1) = "Diag" Then
            SeqAlign1 = Mid(Seq1, y, 1) + SeqAlign1
            SeqAlign2 = Mid(Seq2, x, 1) + SeqAlign2
            i = i - 1
            j = j - 1
        Else
            If TraceBack(i - 1, j - 1) = "Beside" Then
                SeqAlign1 = Mid(Seq1, y, 1) + SeqAlign1
                SeqAlign2 = "-" + SeqAlign2
                j = j - 1
            Else
                SeqAlign1 = "-" + SeqAlign1
                SeqAlign2 = Mid(Seq2, x, 1) + SeqAlign2
                i = i - 1
            End If
        End If
    Loop

    Dim Pairing As String
    For i = 1 To Len(SeqAlign1)
        
        If Not Mid(SeqAlign1, i, 1) = Mid(SeqAlign2, i, 1) Then

            Pairing = Pairing & " "
            
        Else
        
            Pairing = Pairing & "|"
            
        End If
        
    Next i
    
    NeedleAlignmnet = UCase(SeqAlign1) & vbNewLine & Pairing & vbNewLine & UCase(SeqAlign2)

End Function



Private Function Max1(a As Integer, b As Integer, c As Integer) As String
    If a = b And b = c And a = c Then
        Max1 = "abc"
        Exit Function
    End If
    
    If a = b And a > c Then
        Max1 = "ab"
        Exit Function
    End If
    
    If b = c And b > a Then
        Max1 = "bc"
        Exit Function
    End If
        
    If a = c And a > b Then
        Max1 = "ac"
        Exit Function
    End If
    
    If a > b Then
        If a > c Then
            Max1 = "a"
        Else
            Max1 = "c"
        End If
    Else
        If b > c Then
            Max1 = "b"
        Else
            Max1 = "c"
        End If
    End If

End Function






