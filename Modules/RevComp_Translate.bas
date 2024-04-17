Attribute VB_Name = "RevComp_Translate"
Option Explicit

Function Reverse(Sequence As String) As String

    'Date:20240415
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel

    Reverse = StrReverse(Sequence)
    

End Function

Function ReverseComplement(Sequence As String) As String

    'Date:20240415
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel

    ReverseComplement = StrReverse(Complement(Sequence))
    

End Function


Function Complement(Sequence As String) As String

    'Date:20240415
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel

    Dim compDNA As String
    
    compDNA = Sequence
    
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
    
    Complement = compDNA

End Function


Function Translate3LettersAA(Sequence As String) As String

    'Date:20240415
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim Codon As String, Translation As String, Length As Integer
    Dim i As Long
    
    Sequence = UCase(Sequence)
    Sequence = Replace(Sequence, "-", "")
    Sequence = Replace(Sequence, " ", "")
    Sequence = Replace(Sequence, "U", "T")
    
    Translation = ""
    
    Length = Len(Sequence) \ 3
    
    For i = 1 To Length
    
        Codon = Left(Sequence, 3)
        Sequence = Right(Sequence, Len(Sequence) - 3)
        
        Select Case Codon
            Case "TTT"
             Translation = Translation + "Phe "
            
            Case "TTC"
             Translation = Translation + "Phe "
            
            Case "TTA"
             Translation = Translation + "Leu "
            
            Case "TTG"
             Translation = Translation + "Leu "
            
            Case "CTT"
             Translation = Translation + "Leu "
            
            Case "CTC"
             Translation = Translation + "Leu "
            
            Case "CTA"
             Translation = Translation + "Leu "
            
            Case "CTG"
             Translation = Translation + "Leu "
            
            Case "ATT"
             Translation = Translation + "Ile "
            
            Case "ATC"
             Translation = Translation + "Ile "
            
            Case "ATA"
             Translation = Translation + "Ile "
            
            Case "ATG"
             Translation = Translation + "Met "
            
            Case "GTT"
             Translation = Translation + "Val "
            
            Case "GTC"
             Translation = Translation + "Val "
            
            Case "GTA"
             Translation = Translation + "Val "
            
            Case "GTG"
             Translation = Translation + "Val "
            
            Case "TCT"
             Translation = Translation + "Ser "
            
            Case "TCC"
             Translation = Translation + "Ser "
            
            Case "TCA"
             Translation = Translation + "Ser "
            
            Case "TCG"
             Translation = Translation + "Ser "
            
            Case "CCT"
             Translation = Translation + "Pro "
            
            Case "CCC"
             Translation = Translation + "Pro "
            
            Case "CCA"
             Translation = Translation + "Pro "
            
            Case "CCG"
             Translation = Translation + "Pro "
            
            Case "ACT"
             Translation = Translation + "Thr "
            
            Case "ACC"
             Translation = Translation + "Thr "
            
            Case "ACA"
             Translation = Translation + "Thr "
            
            Case "ACG"
             Translation = Translation + "Thr "
            
            Case "GCT"
             Translation = Translation + "Ala "
            
            Case "GCC"
             Translation = Translation + "Ala "
            
            Case "GCA"
             Translation = Translation + "Ala "
            
            Case "GCG"
             Translation = Translation + "Ala "
            
            Case "TAT"
             Translation = Translation + "Tyr "
            
            Case "TAC"
             Translation = Translation + "Tyr "
            
            Case "TAA"
             Translation = Translation + "Stp "
             Exit For
            
            Case "TAG"
             Translation = Translation + "Stp "
             Exit For
             
            Case "CAT"
             Translation = Translation + "His "
            
            Case "CAC"
             Translation = Translation + "His "
            
            Case "CAA"
             Translation = Translation + "Gln "
            
            Case "CAG"
             Translation = Translation + "Gln "
            
            Case "AAT"
             Translation = Translation + "Asn "
            
            Case "AAC"
             Translation = Translation + "Asn "
            
            Case "AAA"
             Translation = Translation + "Lys "
            
            Case "AAG"
             Translation = Translation + "Lys "
            
            Case "GAT"
             Translation = Translation + "Asp "
            
            Case "GAC"
             Translation = Translation + "Asp "
            
            Case "GAA"
             Translation = Translation + "Glu "
            
            Case "GAG"
             Translation = Translation + "Glu "
            
            Case "TGT"
             Translation = Translation + "Cys "
            
            Case "TGC"
             Translation = Translation + "Cys "
            
            Case "TGA"
             Translation = Translation + "Stp "
             Exit For
             
            Case "TGG"
             Translation = Translation + "Trp "
            
            Case "CGT"
             Translation = Translation + "Arg "
            
            Case "CGC"
             Translation = Translation + "Arg "
            
            Case "CGA"
             Translation = Translation + "Arg "
            
            Case "CGG"
             Translation = Translation + "Arg "
            
            Case "AGT"
             Translation = Translation + "Ser "
            
            Case "AGC"
             Translation = Translation + "Ser "
            
            Case "AGA"
             Translation = Translation + "Arg "
            
            Case "AGG"
             Translation = Translation + "Arg "
            
            Case "GGT"
             Translation = Translation + "Gly "
            
            Case "GGC"
             Translation = Translation + "Gly "
            
            Case "GGA"
             Translation = Translation + "Gly "
            
            Case "GGG"
             Translation = Translation + "Gly "
            
            Case Else
            Translation = Translation + "--- "
        End Select
        
    Next i
    
    Translate3LettersAA = Translation + " "
    
End Function



Function Translate(Sequence As String) As String

    'Date:20240415
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim Codon As String, Translation As String, Length As Integer
    Dim i As Long
    
    Sequence = UCase(Sequence)
    Sequence = Replace(Sequence, "-", "")
    Sequence = Replace(Sequence, " ", "")
    Sequence = Replace(Sequence, "U", "T")
    
    Translation = ""
    
    Length = Len(Sequence) \ 3
    
    For i = 1 To Length
    
        Codon = Left(Sequence, 3)
        Sequence = Right(Sequence, Len(Sequence) - 3)
        
        Select Case Codon
            Case "TTT"
             Translation = Translation + "F"
            
            Case "TTC"
             Translation = Translation + "F"
            
            Case "TTA"
             Translation = Translation + "L"
            
            Case "TTG"
             Translation = Translation + "L"
            
            Case "CTT"
             Translation = Translation + "L"
            
            Case "CTC"
             Translation = Translation + "L"
            
            Case "CTA"
             Translation = Translation + "L"
            
            Case "CTG"
             Translation = Translation + "L"
            
            Case "ATT"
             Translation = Translation + "I"
            
            Case "ATC"
             Translation = Translation + "I"
            
            Case "ATA"
             Translation = Translation + "I"
            
            Case "ATG"
             Translation = Translation + "M"
            
            Case "GTT"
             Translation = Translation + "V"
            
            Case "GTC"
             Translation = Translation + "V"
            
            Case "GTA"
             Translation = Translation + "V"
            
            Case "GTG"
             Translation = Translation + "V"
            
            Case "TCT"
             Translation = Translation + "S"
            
            Case "TCC"
             Translation = Translation + "S"
            
            Case "TCA"
             Translation = Translation + "S"
            
            Case "TCG"
             Translation = Translation + "S"
            
            Case "CCT"
             Translation = Translation + "P"
            
            Case "CCC"
             Translation = Translation + "P"
            
            Case "CCA"
             Translation = Translation + "P"
            
            Case "CCG"
             Translation = Translation + "P"
            
            Case "ACT"
             Translation = Translation + "T"
            
            Case "ACC"
             Translation = Translation + "T"
            
            Case "ACA"
             Translation = Translation + "T"
            
            Case "ACG"
             Translation = Translation + "T"
            
            Case "GCT"
             Translation = Translation + "A"
            
            Case "GCC"
             Translation = Translation + "A"
            
            Case "GCA"
             Translation = Translation + "A"
            
            Case "GCG"
             Translation = Translation + "A"
            
            Case "TAT"
             Translation = Translation + "Y"
            
            Case "TAC"
             Translation = Translation + "Y"
            
            Case "TAA"
             Translation = Translation + "*"
            
            Case "TAG"
             Translation = Translation + "*"
             
            Case "CAT"
             Translation = Translation + "H"
            
            Case "CAC"
             Translation = Translation + "H"
            
            Case "CAA"
             Translation = Translation + "Q"
            
            Case "CAG"
             Translation = Translation + "Q"
            
            Case "AAT"
             Translation = Translation + "N"
            
            Case "AAC"
             Translation = Translation + "N"
            
            Case "AAA"
             Translation = Translation + "K"
            
            Case "AAG"
             Translation = Translation + "K"
            
            Case "GAT"
             Translation = Translation + "D"
            
            Case "GAC"
             Translation = Translation + "D"
            
            Case "GAA"
             Translation = Translation + "E"
            
            Case "GAG"
             Translation = Translation + "E"
            
            Case "TGT"
             Translation = Translation + "C"
            
            Case "TGC"
             Translation = Translation + "C"
            
            Case "TGA"
             Translation = Translation + "*"
             
            Case "TGG"
             Translation = Translation + "W"
            
            Case "CGT"
             Translation = Translation + "R"
            
            Case "CGC"
             Translation = Translation + "R"
            
            Case "CGA"
             Translation = Translation + "R"
            
            Case "CGG"
             Translation = Translation + "R"
            
            Case "AGT"
             Translation = Translation + "S"
            
            Case "AGC"
             Translation = Translation + "S"
            
            Case "AGA"
             Translation = Translation + "R"
            
            Case "AGG"
             Translation = Translation + "R"
            
            Case "GGT"
             Translation = Translation + "G"
            
            Case "GGC"
             Translation = Translation + "G"
            
            Case "GGA"
             Translation = Translation + "G"
            
            Case "GGG"
             Translation = Translation + "G"
            
            Case Else
            Translation = Translation + "?"
        End Select
        
    Next i
    
    Translate = Translation + " "
    
End Function
