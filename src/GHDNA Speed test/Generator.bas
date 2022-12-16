Attribute VB_Name = "Module1"
'  ________________________________                          ___________
' /  GHDNA function                \________________________/   v1.00   |
' |                                                                     |
' |                                                                     |
' |     Description:  GHDNA: A new hash function for DNA sequences      |
' |                   used in database engine design                    |
' |                                                                     |
' |          Author:  Dr. Paul A. Gagniuc                               |
' |                                                                     |
' |    Date Created:  July 2010                                         |
' |          Update:  December 2022                                     |
' |       Tested On:  Win Vista, Win XP, Win 7, Win 10, Win 11          |
' |           Email:  paul_gagniuc@acad.ro                              |
' |                                                                     |
' |           Notes:  GHDNA Speed test                                  |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|
'

Function GENEREAZA_NUCLEOTIDE(ByVal nr As Variant, ByVal tip As String) As String

    Dim nucleo(1 To 5) As String
    
    nucleo(1) = "A"
    nucleo(2) = "T"
    nucleo(3) = "G"
    nucleo(4) = "C"
    nucleo(5) = "U"
    
    For N = 1 To nr
    
        DoEvents
        
        If (tip = "ADN") Then
            C = Int(3 * Rnd(3))
            p = p & nucleo(C + 1)
        End If
        
        If (tip = "ARN") Then
            C = Int(4 * Rnd(4))
            If (C + 1 = 2) Then C = 4
            p = p & nucleo(C + 1)
        End If
    
    Next N

    GENEREAZA_NUCLEOTIDE = p

End Function
