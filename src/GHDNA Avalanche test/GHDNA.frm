VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GHDNA"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   14625
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox GDBA 
      Caption         =   "Use GHDNA_DATA_BLOCK"
      Height          =   255
      Left            =   6960
      TabIndex        =   23
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Test type"
      Height          =   1575
      Left            =   9960
      TabIndex        =   19
      Top             =   720
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "Insert nucleotide"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Replace nucleotide"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select test"
      Height          =   2055
      Left            =   12240
      TabIndex        =   14
      Top             =   240
      Width           =   2175
      Begin VB.CheckBox GC 
         Caption         =   "test for G"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox CC 
         Caption         =   "test for C"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox TC 
         Caption         =   "test for T"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox AC 
         Caption         =   "test for A"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Avalanche_test 
      Height          =   7575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   1080
      Width           =   6615
   End
   Begin VB.CommandButton Avalanche 
      Caption         =   "Avalanche"
      Height          =   975
      Left            =   6960
      TabIndex        =   12
      Top             =   840
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "X = GHDNA domain range from 0 to 10^14 and Y=DNA lenght"
      Height          =   6255
      Left            =   6840
      TabIndex        =   3
      Top             =   2400
      Width           =   7575
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4935
         Left            =   600
         ScaleHeight     =   325
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   437
         TabIndex        =   5
         Top             =   720
         Width           =   6615
      End
      Begin VB.CommandButton EO 
         Caption         =   "Erase output"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   6495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   480
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   600
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   7200
         X2              =   7200
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Label Label1 
         Caption         =   "3b"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   5760
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "1000b"
         Height          =   255
         Left            =   6960
         TabIndex        =   10
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "10^14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   5520
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "10^7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   600
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label6 
         Caption         =   "<- DNA lenght ->"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   5760
         Width           =   2175
      End
   End
   Begin VB.CommandButton Start_GHDNA 
      Caption         =   "Get GHDNA"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox INPUT_SEQUENCE 
      Height          =   405
      Left            =   3240
      TabIndex        =   0
      Text            =   "GCACACACCAACCGTACATATTATATTCGCGCGATTACTCGCAACCGTAACACCAATTCGCGCGATTACGCGATTACTCG"
      Top             =   240
      Width           =   7695
   End
   Begin VB.Label DNA_LEN 
      Caption         =   "0b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label OUTPUT_SEQUENCE 
      Caption         =   "00000000000000"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   1815
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
' |           Notes:  GHDNA Avalanche test                              |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|

Dim tmpx As Variant
Dim tmpy As Variant


Function f(ByVal nucleotide As String) As Integer

        If nucleotide = "A" Then f = 3
        If nucleotide = "T" Then f = 5
        If nucleotide = "C" Then f = 7
        If nucleotide = "G" Then f = 11
        
End Function


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


Function GHDNA_DATA_BLOCK(ByVal sequence As String) As Variant

    Dim a, b, C As String
    Dim i, BlockSize As Long
    Dim EA, EB, u As Integer
    
    BlockSize = Block_Alocation(Len(sequence))
    b = "12345678912345"
    
    For i = 1 To Len(sequence) Step BlockSize
    
        a = GHDNA(Mid(sequence, i, BlockSize))
    
        For u = 1 To 14
            EA = Val(Mid(a, u, 1))
            EB = Val(Mid(b, u, 1))
            C = C & (Val(EA + EB) Mod 10)
        Next u
    
        b = C
        C = ""
    Next i
    
    GHDNA_DATA_BLOCK = b
    
End Function


Function Block_Alocation(ByVal L As Variant) As Integer

    Dim a, t, b, m As Integer
    
     a = 1
     t = 1
     b = 1
     m = 10
    
    Do Until t > 3 And v = 0
        a = a + 1
        t = (L Mod a)
        r = (L - t)
        v = r Mod 2
    Loop
    
    
    Do Until b = 0 Or m > 1000
        m = m + 1
        b = r Mod m
    Loop
    
    Block_Alocation = m
    
End Function


Private Sub Avalanche_Click()

    Dim N(1 To 4) As String
    
    Picture1.Print " IC=" & IC(INPUT_SEQUENCE) & " %"
    
    If IC(INPUT_SEQUENCE) > 99 Then MsgBox "IC > 99"
    
    yy = Picture1.ScaleHeight / 10 ^ 14
    xx = Picture1.ScaleWidth / Len(INPUT_SEQUENCE)
    
    N(1) = "A"
    N(2) = "T"
    N(3) = "C"
    N(4) = "G"
    
    For u = 1 To 4
    
        tmpx = 0
        tmpy = 0
        
        If AC.Value = 0 And N(u) = "A" Then GoTo next_U
        If TC.Value = 0 And N(u) = "T" Then GoTo next_U
        If CC.Value = 0 And N(u) = "C" Then GoTo next_U
        If GC.Value = 0 And N(u) = "G" Then GoTo next_U
    
        For i = 1 To Len(INPUT_SEQUENCE)
        
            If Option1(0).Value = True Then
                tmp_avalache = Mid(INPUT_SEQUENCE, 1, i) & N(u) & Mid(INPUT_SEQUENCE, i + 1, Len(INPUT_SEQUENCE))
            End If
            
            If Option1(1).Value = True Then
                tmp_avalache = Mid(INPUT_SEQUENCE, 1, i) & N(u) & Mid(INPUT_SEQUENCE, i, Len(INPUT_SEQUENCE))
        
            End If

            If GDBA.Value = 1 Then
                tmp_hash = GHDNA_DATA_BLOCK(tmp_avalache)
            Else
                tmp_hash = GHDNA(tmp_avalache)
            End If
            
            
            If N(u) = "A" Then color_graph = vbRed
            If N(u) = "T" Then color_graph = vbBlue
            If N(u) = "C" Then color_graph = vbYallow
            If N(u) = "G" Then color_graph = vbGreen
            
            Picture1.Line (tmpx, tmpy)-(xx * i, Val(yy * tmp_hash)), color_graph
            
            tmpx = (xx * i)
            tmpy = (Val(yy * tmp_hash))
            
            OUTPUT_SEQUENCE.Caption = tmp_hash
            Avalanche_test = Avalanche_test & vbCrLf & tmp_hash & " - " & tmp_avalache
    
        Next i
next_U:
    Next u
    
End Sub


Function IC(ByVal sequence As String) As Variant

    Dim count As Long
    Dim total As Double
    Dim max As Variant
    Dim s1 As String
    Dim s2 As String
    Dim i As Long
    Dim u As Long

    s1 = sequence
    max = Len(s1) - 1

    For u = 1 To max
    DoEvents
        s2 = Mid(s1, u + 1)

        For i = 1 To Len(s2)
        DoEvents
            If Mid(s1, i, 1) = Mid(s2, i, 1) Then
                count = count + 1
            End If
        Next i
        
        total = total + (count / Len(s2) * 100)
        count = 0
        
    Next u
    
    IC = Round((total / max), 2)
    
End Function


Private Sub EO_Click()
    Picture1.Cls
    Avalanche_test.Text = ""
End Sub


Private Sub Form_Load()
    INPUT_SEQUENCE_Change
End Sub


Private Sub INPUT_SEQUENCE_Change()
    Label2.Caption = Len(INPUT_SEQUENCE) & " b"
    DNA_LEN.Caption = Len(INPUT_SEQUENCE) & " b"
End Sub


Private Sub Start_GHDNA_Click()
    If GDBA.Value = 1 Then
        OUTPUT_SEQUENCE.Caption = GHDNA_DATA_BLOCK(INPUT_SEQUENCE)
    Else
        OUTPUT_SEQUENCE.Caption = GHDNA(INPUT_SEQUENCE)
    End If
End Sub
