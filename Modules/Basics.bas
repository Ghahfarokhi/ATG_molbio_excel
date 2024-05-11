Attribute VB_Name = "Module_RevComp_Translate"
Option Explicit

Public Function Reverse(Sequence As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel

    Reverse = StrReverse(Sequence)
    

End Function

Public Function ReverseComplement_Relaxed(Sequence As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel

    ReverseComplement_Relaxed = StrReverse(Complement_Relaxed(Sequence))
    
End Function

Public Function ReverseComplement(Sequence As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim dnaComp As String
    
    dnaComp = Complement(Sequence)
    
    If InStr(1, dnaComp, "Error") > 0 Then
        
        ReverseComplement = dnaComp
    
    Else
        
        ReverseComplement = StrReverse(dnaComp)
    
    End If
    
End Function


Public Function Complement(Sequence As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    If Len(Sequence) = 0 Then
        Complement = ""
        Exit Function
    End If

    Dim compDNA As String
    
    compDNA = UCase(Sequence)
    
    If Not Len(Replace(Replace(Replace(Replace(Replace(compDNA, "A", ""), "T", ""), "C", ""), "G", ""), "U", "")) > 0 Then
        
        compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "A", "1"), "T", "2"), "C", "3"), "G", "4"), "U", "5")
        compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "1", "T"), "2", "A"), "3", "G"), "4", "C"), "5", "A")
        Complement = compDNA
        
    Else
        
        Complement = "Error : non-DNA letters in input!"
        
    End If
    
    

End Function


Public Function Complement_Relaxed(Sequence As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel

    Dim compDNA As String
    
    compDNA = Sequence
    
    compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "1", " "), "2", " "), "3", " "), "4", " "), "5", " ")
    compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "6", " "), "7", " "), "8", " "), "9", " "), "0", " ")
    
    'ATCG UpperCase
    compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "A", "1"), "T", "2"), "C", "3"), "G", "4"), "U", "5")
    compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "1", "T"), "2", "A"), "3", "G"), "4", "C"), "5", "A")
    
    'ATCG LowerCase
    compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "a", "1"), "t", "2"), "c", "3"), "g", "4"), "u", "5")
    compDNA = Replace(Replace(Replace(Replace(Replace(compDNA, "1", "t"), "2", "a"), "3", "g"), "4", "c"), "5", "a")
    
    'Y     pYrimidine              C T          R
    'R     puRine                  A G          Y
    'K     Keto                    T/U G        M
    'M     aMino                   A C          K
    
    'IUPAC 2-letters Uppercase
    compDNA = Replace(Replace(Replace(Replace(compDNA, "Y", "1"), "R", "2"), "K", "3"), "M", "4")
    compDNA = Replace(Replace(Replace(Replace(compDNA, "1", "R"), "2", "Y"), "3", "M"), "4", "K")
    
    'IUPAC 2-letters Lowercase
    compDNA = Replace(Replace(Replace(Replace(compDNA, "y", "1"), "r", "2"), "k", "3"), "m", "4")
    compDNA = Replace(Replace(Replace(Replace(compDNA, "1", "r"), "2", "y"), "3", "m"), "4", "k")
    

    'B     not A                   C G T        V
    'D     not C                   A G T        H
    'H     not G                   A C T        D
    'V     not T/U                 A C G        B

    'IUPAC 3-letters Uppercase
    compDNA = Replace(Replace(Replace(Replace(compDNA, "B", "1"), "D", "2"), "H", "3"), "V", "4")
    compDNA = Replace(Replace(Replace(Replace(compDNA, "1", "V"), "2", "H"), "3", "D"), "4", "B")
    
    'IUPAC 3-letters Lowercase
    compDNA = Replace(Replace(Replace(Replace(compDNA, "b", "1"), "d", "2"), "h", "3"), "v", "4")
    compDNA = Replace(Replace(Replace(Replace(compDNA, "1", "v"), "2", "h"), "3", "d"), "4", "b")
    
    'S     Strong(3Hbonds)         G C          S
    'W     Weak(2Hbonds)           A T          W
    'N     Unknown                 A C G T      N
    
    'Note that S, W, and N are already RevComp of themselves!
    
    Complement_Relaxed = compDNA

End Function

Public Function gcContent(Seq As String) As Double
    
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim DNA As String
    DNA = DNA_Sanitizer(Seq)
    
    If Len(DNA) > 0 Then
    
        gcContent = Round(Len(Replace(Replace(DNA, "A", ""), "T", "")) / Len(DNA) * 100, 2)
        
    Else
        
        gcContent = 0
        
    End If

End Function



Private Function DNA_Sanitizer(ByVal Seq As String) As String
    
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim DNA As String, SanitizedDNA As String, i As Long, Base As String
    
    DNA = UCase(Seq)
    DNA = Replace(DNA, "U", "T")
    
    For i = 1 To Len(DNA)
        
        Base = Mid(DNA, i, 1)
        
        If InStr(1, "ATCG", Base) > 0 Then
            
            SanitizedDNA = SanitizedDNA & Base
            
        End If
        
    Next i
    
    DNA_Sanitizer = SanitizedDNA

End Function


Public Function Translate3LettersAA(ByVal Sequence As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim Codon As String, Codons As String, Translation As String, Length As Integer
    Dim i As Long
    
    Sequence = UCase(Sequence)
    Sequence = Replace(Sequence, "-", "")
    Sequence = Replace(Sequence, " ", "")
    Sequence = Replace(Sequence, "U", "T")
    
    Translation = ""
    
    Length = Len(Sequence) \ 3
    
    Codons = "_0:---_TTT:Phe_TCT:Ser_TAT:Tyr_TGT:Cys_TTC:Phe_TCC:Ser_TAC:Tyr_TGC:Cys_TTA:Leu_TCA:Ser_TAA:Stp_TGA:Stp_TTG:Leu_TCG:Ser_TAG:Stp_TGG:Trp_CTT:Leu_CCT:Pro_CAT:His_CGT:Arg_CTC:Leu_CCC:Pro_CAC:His_CGC:Arg_CTA:Leu_CCA:Pro_CAA:Gln_CGA:Arg_CTG:Leu_CCG:Pro_CAG:Gln_CGG:Arg_ATT:Ile_ACT:Thr_AAT:Asn_AGT:Ser_ATC:Ile_ACC:Thr_AAC:Asn_AGC:Ser_ATA:Ile_ACA:Thr_AAA:Lys_AGA:Arg_ATG:Met_ACG:Thr_AAG:Lys_AGG:Arg_GTT:Val_GCT:Ala_GAT:Asp_GGT:Gly_GTC:Val_GCC:Ala_GAC:Asp_GGC:Gly_GTA:Val_GCA:Ala_GAA:Glu_GGA:Gly_GTG:Val_GCG:Ala_GAG:Glu_GGG:Gly_"
    
    For i = 1 To Length
    
        Codon = Left(Sequence, 3)
        Sequence = Right(Sequence, Len(Sequence) - 3)
        
        Translation = Translation + Right(Mid(Codons, InStr(1, Codons, Codon) + 1, 6), 3) + " "
        
    Next i
    
    If InStr(1, Translation, "Stp") > 0 Then
        
        Translation = Left(Translation, InStr(1, Translation, "Stp") + 3)
        
    End If
    
    Translate3LettersAA = Translation + " "
    
End Function



Public Function Translate(ByVal Sequence As String) As String

    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim Codon As String, Codons As String, Translation As String, Length As Integer
    Dim i As Long
    
    Sequence = UCase(Sequence)
    Sequence = Replace(Sequence, "-", "")
    Sequence = Replace(Sequence, " ", "")
    Sequence = Replace(Sequence, "U", "T")
    
    On Error Resume Next
    
    Translation = ""
    
    Length = Len(Sequence) \ 3
    
    Codons = "_0:?_TTT:F_TCT:S_TAT:Y_TGT:C_TTC:F_TCC:S_TAC:Y_TGC:C_TTA:L_TCA:S_TAA:*_TGA:*_TTG:L_TCG:S_TAG:*_TGG:W_CTT:L_CCT:P_CAT:H_CGT:R_CTC:L_CCC:P_CAC:H_CGC:R_CTA:L_CCA:P_CAA:Q_CGA:R_CTG:L_CCG:P_CAG:Q_CGG:R_ATT:I_ACT:T_AAT:N_AGT:S_ATC:I_ACC:T_AAC:N_AGC:S_ATA:I_ACA:T_AAA:K_AGA:R_ATG:M_ACG:T_AAG:K_AGG:R_GTT:V_GCT:A_GAT:D_GGT:G_GTC:V_GCC:A_GAC:D_GGC:G_GTA:V_GCA:A_GAA:E_GGA:G_GTG:V_GCG:A_GAG:E_GGG:G_"
    
    
    For i = 1 To Length
    
        Codon = Left(Sequence, 3)
        Sequence = Right(Sequence, Len(Sequence) - 3)
        
        Translation = Translation + Right(Mid(Codons, InStr(1, Codons, Codon) + 1, 4), 1)
        
    Next i

    
    Translate = Translation
    
End Function
