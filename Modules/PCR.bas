Attribute VB_Name = "Module_PCR"

Public Function PCR(Fwd_Primer As String, Rev_Primer As String, Template As String, Optional Mode As String = "Sequence") As Variant
    
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    ' Convert all sequences uo upper case
    
    Fwd_Primer = UCase(Fwd_Primer)
    Rev_Primer = UCase(Rev_Primer)
    Template = UCase(Template)
    
    ' Check all sequences are valid
    
    If Len(Replace(Replace(Replace(Replace(Fwd_Primer, "A", ""), "T", ""), "C", ""), "G", "")) > 0 Then
        PCR = "ERROR : Fwd sequence contains invalid charchters! Only ATCGatcg is allowed in the input!"
        Exit Function
    ElseIf Len(Replace(Replace(Replace(Replace(Rev_Primer, "A", ""), "T", ""), "C", ""), "G", "")) > 0 Then
        PCR = "ERROR : Rev sequence contains invalid charchters! Only ATCGatcg is allowed in the input!"
        Exit Function
    ElseIf Len(Replace(Replace(Replace(Replace(Template, "A", ""), "T", ""), "C", ""), "G", "")) > 0 Then
        PCR = "ERROR : Template sequence contains invalid charchters! Only ATCGatcg is allowed in the input!"
        Exit Function
    End If
    
    'Check if length of primers is above 15
    
    If Len(Fwd_Primer) < 15 Then
        PCR = "ERROR : Length of Fwd Primer is below 15 nucleotides!"
        Exit Function
    ElseIf Len(Rev_Primer) < 15 Then
        PCR = "ERROR : Length of Rev Primer is below 15 nucleotides!"
        Exit Function
    ElseIf Len(Template) < 30 Then
        PCR = "ERROR : Length of Template is below 30 nucleotides!"
        Exit Function
    End If
    
    ' Extract the 15 right nucleotides of Fwd and Rev primers and prepare revcomps
    
    Dim Fwd_15 As String, Rev_15 As String
    Dim Fwd_15_RC As String, Rev_15_RC As String, Template_RC As String
    
    Fwd_15 = Right(Fwd_Primer, 15)
    Rev_15 = Right(Rev_Primer, 15)
    Fwd_15_RC = ReverseComplement(Fwd_15)
    Rev_15_RC = ReverseComplement(Rev_15)
    Template_RC = ReverseComplement(Template)
    
    ' Check if Fwd and Rev primers exist in Template
    
    Dim Fwd_in_Sense As Integer, Fwd_in_AntiSense As Integer
    Dim Rev_in_Sense As Integer, Rev_in_AntiSense As Integer
    
    Fwd_in_Sense = (Len(Template) - Len(Replace(Template, Fwd_15, ""))) / 15
    Fwd_in_AntiSense = (Len(Template) - Len(Replace(Template_RC, Fwd_15, ""))) / 15
    
    Rev_in_Sense = (Len(Template) - Len(Replace(Template, Rev_15, ""))) / 15
    Rev_in_AntiSense = (Len(Template) - Len(Replace(Template_RC, Rev_15, ""))) / 15
    
    If (Fwd_in_Sense = 1 And Fwd_in_AntiSense = 0 And Rev_in_Sense = 0 And Rev_in_AntiSense = 1) Or _
       (Fwd_in_Sense = 0 And Fwd_in_AntiSense = 1 And Rev_in_Sense = 1 And Rev_in_AntiSense = 0) Then
        
        ' Continue
        
    Else
        
        If Fwd_in_Sense > 1 Or Fwd_in_AntiSense > 1 Then
            PCR = "ERROR : Fwd primers exists in the template more than once!"
            Exit Function
        ElseIf Rev_in_Sense > 1 Or Rev_in_AntiSense > 1 Then
            PCR = "ERROR : Rev primers exists in the template more than once!"
            Exit Function
        ElseIf Fwd_in_Sense = 0 And Fwd_in_AntiSense = 0 Then
            PCR = "ERROR : Fwd primer doesn't exist in the template !"
            Exit Function
        ElseIf Rev_in_Sense = 0 And revd_in_AntiSense = 0 Then
            PCR = "ERROR : Fwd primer doesn't exist in the template !"
            Exit Function
        End If
        
        PCR = "ERROR in finding primers in template!"
        Exit Function
        
    End If
    
    Dim Start_Pos As Integer, End_Pos As Integer
    
    If Fwd_in_Sense = 1 And Fwd_in_AntiSense = 0 And Rev_in_Sense = 0 And Rev_in_AntiSense = 1 Then
        
        Start_Pos = InStr(1, Template, Fwd_15)
        End_Pos = InStr(1, Template, Rev_15_RC) + 15
        
        If End_Pos - Start_Pos < 15 Then
        
            PCR = "ERROR : Fwd and Rev Primers exist in the Template, but not facing towards each other!"
            Exit Function
            
        End If
        
        PCR = Mid(Template, Start_Pos, End_Pos - Start_Pos)
        PCR = Left(Fwd_Primer, Len(Fwd_Primer) - 15) + PCR + ReverseComplement(Left(Rev_Primer, Len(Rev_Primer) - 15))
        
        If Mode = "Sequence" Then
            Exit Function
        Else
            PCR = Len(PCR)
            Exit Function
        End If
        
    End If
    
    
    If Fwd_in_Sense = 0 And Fwd_in_AntiSense = 1 And Rev_in_Sense = 1 And Rev_in_AntiSense = 0 Then
        
        Start_Pos = InStr(1, Template_RC, Fwd_15)
        End_Pos = InStr(1, Template_RC, Rev_15_RC) + 15
        
        If End_Pos - Start_Pos < 15 Then
        
            PCR = "ERROR : Fwd and Rev Primers exist in the Template, but not facing towards each other!"
            
        End If
        
        PCR = Mid(Template_RC, Start_Pos, End_Pos - Start_Pos)
        PCR = Left(Fwd_Primer, Len(Fwd_Primer) - 15) + PCR + ReverseComplement(Left(Rev_Primer, Len(Rev_Primer) - 15))
        
        If Mode = "Sequence" Then
            Exit Function
        Else
            PCR = Len(PCR)
            Exit Function
        End If
        
    End If
    
End Function

