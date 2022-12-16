VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GHDNA"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Start_GHDNA 
      Caption         =   "Get HASH"
      Height          =   735
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox INPUT_SEQUENCE 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   $"GHDNA.frx":0000
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Hash value:"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Sequence_lenght 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 b"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label measure 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Processing time: 0 ms"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label OUTPUT_SEQUENCE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00000000000000"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' |           Notes:  Article exemple                                   |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long


Function GHDNA(ByVal sequence As String) As String

    Dim correction As Variant
    Dim N(1 To 3) As String
    Dim Prehash As Variant
    Dim hash As Variant
    Dim alfa As Variant
    Dim beta As Variant
    Dim C As Variant
    Dim x As Variant
    Dim i As Variant
    Dim u As Variant


    t = (Len(sequence) - (Len(sequence) Mod 2))
    beta = ((Len(sequence) - (Len(sequence) Mod 2)) / 2) - 1

    For i = 1 To beta

        N(1) = Mid(sequence, 2 * i - 1, 1)
        N(2) = Mid(sequence, 2 * i, 1)
        N(3) = Mid(sequence, 2 * i + 1, 1)
        
        C1 = (f(N(1)) - Sqr((i Mod 2) + 1))
        C2 = ((f(N(2)))) - Sqr((i Mod 3) + 1)
        C3 = f(N(3))
        
        C = C + ((C1 * C2) / C3)
            
    Next i

    For u = t To Len(sequence)

        N(1) = Mid(sequence, u, 1)
        C = C + (f(N(1)) - (Len(sequence) - Sqr(u))) / f(N(1))

    Next u

    ID = Len(sequence) / C
    Prehash = Round(ID * 10 ^ 14) - Len(sequence)
    DS = Mid(Prehash, 8, 7) & Mid(Prehash, 1, 7)
    
    x = Len(sequence) Mod 10
    DU = Mid(DS, 1, 7) & x & Mid(DS, 9, 6)
    
    GHDNA = DU

End Function

Function f(ByVal nucleotide As String) As Integer

    If nucleotide = "A" Then f = 3
    If nucleotide = "T" Then f = 5
    If nucleotide = "C" Then f = 7
    If nucleotide = "G" Then f = 11

End Function

Private Sub Form_Load()
    INPUT_SEQUENCE_Change
End Sub

Private Sub INPUT_SEQUENCE_Change()
    Sequence_lenght.Caption = Len(INPUT_SEQUENCE) & " b"
End Sub

Private Sub Start_GHDNA_Click()

    Dim tm As Long
    Dim tma As Long
    
    tm = timeGetTime
    OUTPUT_SEQUENCE.Caption = GHDNA(INPUT_SEQUENCE)
    tma = timeGetTime
    
    measure.Caption = "Processing time: " & Val(tma - tm) & " ms"

End Sub


