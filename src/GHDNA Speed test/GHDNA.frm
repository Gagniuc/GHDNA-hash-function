VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GHDNA speed test"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar DNA_BLOCK 
      Height          =   255
      Left            =   1320
      Max             =   10000
      Min             =   100
      TabIndex        =   26
      Top             =   4320
      Value           =   1000
      Width           =   1935
   End
   Begin VB.CheckBox Use 
      Caption         =   "Use GHDNA_DATA_BLOCK"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   25
      Top             =   5160
      Width           =   2655
   End
   Begin VB.CheckBox Use 
      Caption         =   "Use GHDNA"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   24
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox IC_OPTION 
      Caption         =   "Process Index of Coincidence"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   3600
      Width           =   2655
   End
   Begin VB.HScrollBar sensitivity 
      Height          =   255
      Left            =   1320
      Max             =   10000
      Min             =   100
      TabIndex        =   19
      Top             =   3960
      Value           =   1000
      Width           =   1935
   End
   Begin VB.OptionButton Plot 
      Caption         =   "Plot POINTS"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   15
      Top             =   2880
      Width           =   1815
   End
   Begin VB.OptionButton Plot 
      Caption         =   "Plot LINES"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   14
      Top             =   2520
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      Height          =   735
      Left            =   2400
      TabIndex        =   12
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton KL 
      Caption         =   "Speed test"
      Height          =   735
      Left            =   480
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Speed test"
      Height          =   6255
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   7695
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4935
         Left            =   600
         ScaleHeight     =   325
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   445
         TabIndex        =   5
         Top             =   720
         Width           =   6735
      End
      Begin VB.CommandButton EO 
         Caption         =   "Erase output"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   6495
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         X1              =   2160
         X2              =   2160
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         X1              =   3960
         X2              =   3960
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         X1              =   5640
         X2              =   5640
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Label Label8 
         Caption         =   "25 Kb"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "75 Kb"
         Height          =   255
         Left            =   5520
         TabIndex        =   17
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "50 Kb"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   5880
         Width           =   495
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
         X1              =   7320
         X2              =   7320
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Label Label1 
         Caption         =   "0b"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   5880
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "100 Kb"
         Height          =   255
         Left            =   6960
         TabIndex        =   9
         Top             =   5880
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "500 ms"
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
         Left            =   80
         TabIndex        =   8
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0 ms"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "250ms"
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
         TabIndex        =   6
         Top             =   2760
         Width           =   615
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   600
         Y1              =   3000
         Y2              =   3000
      End
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   4200
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label11 
      Caption         =   "DNA block:"
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label block 
      Caption         =   "1000 b"
      Height          =   255
      Left            =   3360
      TabIndex        =   27
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label IC_SHOW 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Index of Coincidence - 0%"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00404040&
      Height          =   3015
      Left            =   240
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   1095
      Left            =   240
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   1935
      Left            =   240
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Step_lenght 
      Caption         =   "1000 b"
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Sensitivity:"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Sequence_generator 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start test ..."
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Sequence_lenght 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 b"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label measure 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Processing time: 0 ms"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label OUTPUT_SEQUENCE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00000000000000"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
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
' |           Notes:  GHDNA Speed test                                  |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim stop_test As Boolean

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
    m = DNA_BLOCK.Value

    Do Until t > 3 And v = 0
        a = a + 1
        t = (L Mod a)
        r = (L - t)
        v = r Mod 2
    Loop

    Do Until b = 0 Or m >= 999
        m = m + 1
        b = r Mod m
    Loop

    Block_Alocation = m

End Function


Private Sub DNA_BLOCK_Change()
    block.Caption = DNA_BLOCK.Value & " b"
End Sub


Private Sub EO_Click()
    Picture1.Cls
End Sub


Private Sub Form_Load()
    stop_test = True
End Sub


Private Sub KL_Click()

    Dim tm As Long
    Dim tma As Long
    
    stop_test = False
    
    yy = Picture1.ScaleHeight / 500
    xx = Picture1.ScaleWidth / 100000
    
    For u = 8 To 110000 Step sensitivity.Value
        
        DoEvents
            
        If stop_test = True Then GoTo 1
                 
            Sequence_generator.Caption = "Generating " & u & " b DNA sequence ..."
            sequence = GENEREAZA_NUCLEOTIDE(u, "ADN")
            Sequence_generator.Caption = "Waiting for GHDNA function ..."
                    
            DoEvents
                
        If IC_OPTION.Value = 1 Then
            Sequence_generator.Caption = "Procesing Index of Coincidence ..."
            IC_SHOW.Caption = "Index of Coincidence - " & IC(sequence) & " %"
            Sequence_generator.Caption = "Waiting for GHDNA function ..."
        End If
                
        tm = timeGetTime
                
        If Use(1).Value = 1 Then tmp = GHDNA_DATA_BLOCK(sequence)
        If Use(0).Value = 1 Then tmp = GHDNA(sequence)
                    
        tma = timeGetTime
                
        Time_Spent = Val(tma - tm)
            
        DoEvents
            
        If Plot(0).Value = True Then
            Picture1.Line (tmpx, tmpy)-(xx * u, Val(yy * Time_Spent))
        Else
            Picture1.PSet (xx * u, Val(yy * Time_Spent)), vbBlue
        End If
            
        tmpx = xx * u
        tmpy = Val(yy * Time_Spent)
          
        measure.Caption = "GHDNA Processing time: " & Time_Spent & " ms"
        OUTPUT_SEQUENCE.Caption = "GHDNA Hash: " & tmp
        Sequence_lenght.Caption = "DNA sequence length " & u & " b"
    
    Next u
1:

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


Private Sub sensitivity_Change()
    Step_lenght.Caption = sensitivity.Value & " b"
End Sub


Private Sub Stop_Click()
    stop_test = True
End Sub


Private Sub Use_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Use(0).Value = 1 Then Use(1).Value = 0 Else Use(1).Value = 1
End Sub
