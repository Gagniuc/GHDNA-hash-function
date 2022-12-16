Attribute VB_Name = "Module1"
'  ________________________________                          ___________
' /  GHDNA function                \________________________/   v1.00   |
' |                                                                     |
' |         Journal:  BMC Bioinformatics                                |
' |     Description:  GHDNA: A new hash function for DNA sequences      |
' |                   used in database engine design                    |
' |                                                                     |
' |        Category:  Software                                          |
' |          Author:  Paul Gagniuc                                      |
' |                                                                     |
' |    Date Created:  July 2010                                         |
' |       Tested On:  Windows Vista, Windows XP, Windows 7              |
' |           Email:  paulgagniuc@yahoo.com                             |
' |                                                                     |
' |           Notes:  GHDNA Domain test                                 |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|
'
Function GENEREAZA_NUCLEOTIDE(ByVal nr As Variant, ByVal tip As String) As String
'***
Dim nucleo(1 To 5) As String
nucleo(1) = "A"
nucleo(2) = "T"
nucleo(3) = "G"
nucleo(4) = "C"
nucleo(5) = "U"

For N = 1 To nr

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
'***
GENEREAZA_NUCLEOTIDE = p
End Function





Function DNA_generator(ByVal no As Variant) As String
Dim nucleotide(1 To 4) As String
Dim N As Integer
Dim C As Integer

nucleotide(1) = "A"
nucleotide(2) = "T"
nucleotide(3) = "G"
nucleotide(4) = "C"

For N = 1 To no

    C = Int(3 * Rnd(3))
    sequence = sequence & nucleotide(C + 1)

Next N

DNA_generator = sequence
End Function


