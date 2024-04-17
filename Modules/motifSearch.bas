Attribute VB_Name = "MotifSearch"

Option Explicit

Function Motifs(Sequence As String, Motif As String, Mode As String) As String
 

    'Date:20240415
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim MotifCandid As String
    Dim i As Long, j As Long, MotifLength As Long, SequenceLength As Long
    Dim NumberFormat As String, ListText As String, IUPACBases As String
    
    Mode = UCase(Mode)
    
    IUPACBases = "_GG_AA_TT_CC_BC_BG_BT_DA_DG_DT_HA_HC_HT_KG_KT_MA_MC_NA_NC_NG_NT_RA_RG_SC_SG_VA_VC_VG_WA_WT_YC_YT_CB_GB_TB_AD_GD_TD_AH_CH_TH_GK_TK_AM_CM_AN_CN_GN_TN_AR_GR_CS_GS_AV_CV_GV_AW_TW_CY_TY_"
    
    NumberFormat = "000"
    
    MotifLength = Len(Motif)
    SequenceLength = Len(Sequence)
    If MotifLength > SequenceLength Then
        MsgBox "Sequence is shorter than Motif's length!", vbCritical, "Error!"
        Exit Function
    End If
    
    j = 1
    For i = 1 To SequenceLength - MotifLength + 1
        
        MotifCandid = Mid(Sequence, i, MotifLength)
        
        If IsMotifPresent(MotifCandid, Motif, IUPACBases) = True Then
            
            ListText = ListText + "F_" + Format(j, NumberFormat) + " " + MotifCandid + vbNewLine
            j = j + 1
            
        End If
        
    Next i
    

    Sequence = ReverseComplement(Sequence)
    
    For i = 1 To SequenceLength - MotifLength + 1
        
        MotifCandid = Mid(Sequence, i, MotifLength)
        
        If IsMotifPresent(MotifCandid, Motif, IUPACBases) = True Then
            
            ListText = ListText + "R_" + Format(j, NumberFormat) + " " + MotifCandid + vbNewLine
            j = j + 1
            
        End If
        
    Next i
    
    If Mode = "LIST" Then
    
        Motifs = Left(ListText, Len(ListText) - 1)
        
    ElseIf Mode = "COUNT" Then
    
        Motifs = "Motif count = " + Str(j - 1)
        
    Else
        
        Motifs = "Incorrect mode argument is provided: " + Mode + ". Please provide either 'List' or 'Count'."
        
    End If

End Function



Private Function IsMotifPresent(DNASequence As String, Motif As String, IUPACBases As String) As Boolean


    'Date:20240415
    'Author:Amir.Taheri.Ghahfarokhi@Gmail.com
    'Github: https://github.com/Ghahfarokhi/ATG_molbio_excel
    
    Dim i As Integer
    Dim DNA_Motif_Chars As String
    
    If Not Len(DNASequence) = Len(Motif) Then
        MsgBox "DNA Sequence and the Motif must have equal length", vbCritical, "IsMotifPresent Function Error!"
        Exit Function
    End If
    
    DNASequence = UCase(DNASequence)
    Motif = UCase(Motif)
    
    For i = 1 To Len(Motif)
        DNA_Motif_Chars = Mid(DNASequence, i, 1) & Mid(Motif, i, 1)
        
        If Not InStr(1, IUPACBases, DNA_Motif_Chars) > 0 Then
            
            IsMotifPresent = False
            Exit Function
            
        End If
        
    Next i
    
    IsMotifPresent = True
    
End Function
