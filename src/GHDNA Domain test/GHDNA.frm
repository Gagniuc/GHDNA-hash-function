VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GHDNA hash function"
   ClientHeight    =   11415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "No. collision / DNA sequence lenght"
      Height          =   4815
      Left            =   9360
      TabIndex        =   27
      Top             =   6480
      Width           =   5055
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Left            =   360
         ScaleHeight     =   261
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   293
         TabIndex        =   28
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "<- DNA lenght ->"
         Height          =   255
         Left            =   1680
         TabIndex        =   34
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00000000&
         X1              =   4800
         X2              =   4800
         Y1              =   4320
         Y2              =   4440
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00000000&
         X1              =   360
         X2              =   360
         Y1              =   4440
         Y2              =   4320
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00000000&
         X1              =   360
         X2              =   240
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         X1              =   360
         X2              =   240
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         X1              =   360
         X2              =   240
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label11 
         Caption         =   "1000b"
         Height          =   255
         Left            =   4440
         TabIndex        =   33
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "3b"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
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
         TabIndex        =   31
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
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
         TabIndex        =   30
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         TabIndex        =   29
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Collision"
      Height          =   6255
      Left            =   11880
      TabIndex        =   23
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox CL 
         Caption         =   "process collisions"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   4935
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   2400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label count_collision 
         Alignment       =   2  'Center
         Caption         =   "0 collisions"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "OutPut"
      Height          =   4695
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Width           =   9135
      Begin VB.CheckBox CO 
         Caption         =   "Count to output"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox out_seq 
         Height          =   3885
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   600
         Width           =   8775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Y = GHDNA domain range from 0 to 10^14 and X=DNA lenght"
      Height          =   6255
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton EO 
         Caption         =   "Erase output"
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Top             =   240
         Width           =   6495
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4935
         Left            =   600
         ScaleHeight     =   325
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   445
         TabIndex        =   6
         Top             =   720
         Width           =   6735
      End
      Begin VB.Label S_lenght 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<- 0 b ->"
         Height          =   195
         Left            =   1080
         TabIndex        =   36
         Top             =   5880
         Width           =   585
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   600
         Y1              =   3000
         Y2              =   3000
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
         TabIndex        =   19
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   600
         Width           =   255
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
         Left            =   0
         TabIndex        =   9
         Top             =   5520
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "1000b"
         Height          =   255
         Left            =   7080
         TabIndex        =   8
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "3b"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   5880
         Width           =   375
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   7320
         X2              =   7320
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   600
         Y1              =   5760
         Y2              =   5640
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   480
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   600
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Methods"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   3975
      Begin VB.CheckBox BS 
         Caption         =   "Use Block Alocation"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton stop 
         Caption         =   "Stop"
         Height          =   615
         Left            =   2160
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton KL 
         Caption         =   "Test"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox CV 
         Caption         =   "correction variable"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox DS 
         Caption         =   "DIGIT SHIFT"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox DU 
         Caption         =   "DIGIT UNCERTITUDE"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   2400
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Block_Size 
         Caption         =   "Block size"
         Height          =   255
         Left            =   2520
         TabIndex        =   37
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "GHDNA manual sequence test"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox copy_hash 
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         Top             =   1080
         Width           =   1695
      End
      Begin VB.HScrollBar DNA_lenght 
         Height          =   255
         Left            =   240
         Min             =   3
         TabIndex        =   22
         Top             =   2640
         Value           =   20
         Width           =   3375
      End
      Begin VB.CommandButton Generate_ADN 
         Caption         =   "Generate DNA - 20 b"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox INPUT_SEQUENCE 
         Height          =   405
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton Start_GHDNA 
         Caption         =   "GHDNA hash"
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label OUTPUT_SEQUENCE 
         Alignment       =   2  'Center
         Caption         =   "00000000000000"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
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
' |           Notes:  GHDNA Domain test                                 |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|

Dim stop_test As Boolean
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
    Dim x As Integer
    Dim i As Integer
    Dim u As Integer

    t = (Len(sequence) - (Len(sequence) Mod 2)) + 1
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


    If C = Empty Then C = 1
    ID = Len(sequence) / C
    
    If CV.Value = 1 Then
        correction = Len(sequence)
    Else
        correction = 0
    End If
    
    Prehash = Round(ID * 10 ^ 14) - correction
    
    If DS.Value = 1 Then
        DSH = Mid(Prehash, 8, 7) & Mid(Prehash, 1, 7)
    Else
        DSH = Prehash
    End If
    
    If DU.Value = 1 Then
        x = Len(sequence) Mod 10
        DUC = Mid(DSH, 1, 7) & x & Mid(DSH, 9, 6)
    Else
        DUC = DSH
    End If
    
    GHDNA = DUC

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


Private Sub DNA_lenght_Change()
    Generate_ADN.Caption = "Generate DNA " & DNA_lenght.Value & " b"
End Sub


Private Sub EO_Click()
    Picture1.Cls
End Sub


Private Sub Form_Load()

    stop_test = True
    tmpx = 0
    tmpy = 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    stop_test = True
End Sub


Private Sub Generate_ADN_Click()
    INPUT_SEQUENCE.Text = GENEREAZA_NUCLEOTIDE(DNA_lenght.Value, "ADN")
End Sub


Private Sub INPUT_SEQUENCE_Change()
    INPUT_SEQUENCE.Text = UCase(INPUT_SEQUENCE.Text)
End Sub


Private Sub Start_GHDNA_Click()

    If Len(INPUT_SEQUENCE) < 3 Then
        MsgBox "Sequence to small for GHDNA !"
        Exit Sub
    End If


    If BS.Value = 1 Then
        OUTPUT_SEQUENCE.Caption = GHDNA_DATA_BLOCK(INPUT_SEQUENCE)
    Else
        OUTPUT_SEQUENCE.Caption = GHDNA(INPUT_SEQUENCE)
    End If

    copy_hash.Text = OUTPUT_SEQUENCE.Caption
    
End Sub


Private Sub KL_Click()

    stop_test = False
    yy = Picture1.ScaleHeight / 10 ^ 14
    xx = Picture1.ScaleWidth / 1000
    
    For u = 8 To 1000
        
        For m = 1 To 10
        
            DoEvents
            
            S_lenght.Caption = "<- Generating and hashing 10 random DNA variants of " & u & " b ->"
            
            If stop_test = True Then GoTo 1
    
back_in:
            INPUT_SEQUENCE.Text = GENEREAZA_NUCLEOTIDE(u, "ADN")
            
            If InStr(1, tmp_collection, INPUT_SEQUENCE.Text) > 0 Then
                GoTo back_in
            End If
                
            tmp_collection = tmp_collection & " - " & INPUT_SEQUENCE.Text
                
            Start_GHDNA_Click
    
            If CO.Value = 1 Then
                out_seq = out_seq & OUTPUT_SEQUENCE.Caption & " - " & INPUT_SEQUENCE.Text & vbCrLf
                out_seq.SelStart = Len(out_seq.Text) - 1
                out_seq.SelLength = 1
            End If
    
    
            If CL.Value = 1 Then
                List1.AddItem OUTPUT_SEQUENCE.Caption
            End If
    
            Picture1.PSet (xx * u, Val(yy * Val(OUTPUT_SEQUENCE.Caption)))
            
        Next m
    
        If CL.Value = 1 Then
            Call Collision_Checker
        End If
    
        If CO.Value = 1 Then
            out_seq = out_seq & " lenght  " & u & " b " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        End If
    
    Next u
    
1:
End Sub


Private Sub stop_Click()
    stop_test = True
End Sub


Function Collision_Checker()

    Dim j As Integer
    
    yy = Picture2.ScaleHeight / 10
    xx = Picture2.ScaleWidth / 1000
    
    j = 0
    
    Do While j < List1.ListCount
        
        List1.Text = List1.List(j)
    
        If List1.ListIndex <> j Then
            List1.RemoveItem j
            count_collision.Caption = Val(count_collision.Caption) + 1
            tmpc = tmpc + 1
        Else
            j = j + 1
        End If
    Loop
    
    Picture2.Line (tmpx, tmpy)-(xx * Len(INPUT_SEQUENCE.Text), Val(yy * tmpc)), vbRed
    tmpx = xx * Len(INPUT_SEQUENCE.Text)
    tmpy = Val(yy * tmpc)
    
End Function




