VERSION 5.00
Begin VB.Form GHDNA_DATABASE 
   Caption         =   "GHDNA DATABASE ENGINE"
   ClientHeight    =   12330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   ScaleHeight     =   12330
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Deleting sequences in db.txt"
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   10080
      Width           =   11895
      Begin VB.CommandButton Delete_by_DNA 
         Caption         =   "Delete sequence"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Text            =   "ttccga"
         Top             =   1440
         Width           =   9495
      End
      Begin VB.CommandButton Delete_by_hash 
         Caption         =   "Delete sequence"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Text            =   "ttccga"
         Top             =   720
         Width           =   9495
      End
      Begin VB.Label Label2 
         Caption         =   "Delete by sequence:"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Delete by hash:"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Introducing new sequences to db.txt"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   8400
      Width           =   11895
      Begin VB.TextBox Input_seq 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Text            =   "ttccga"
         Top             =   720
         Width           =   9495
      End
      Begin VB.CommandButton Input_DB 
         Caption         =   "Store sequence"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Motif sequence:"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search db.txt"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   11895
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2160
         TabIndex        =   6
         Text            =   "gggcagtgctgatcgtagccattccggactgtagctatgc"
         Top             =   720
         Width           =   9495
      End
      Begin VB.CommandButton Search2 
         Caption         =   "Motif discoverer"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Search1 
         Caption         =   "Motif direct match"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2160
         TabIndex        =   3
         Text            =   "ttccgg"
         Top             =   1440
         Width           =   9495
      End
      Begin VB.Label Label5 
         Caption         =   "Enter Motif sequence:"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Enter DNA sequence:"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Result"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   360
         Width           =   11415
      End
   End
End
Attribute VB_Name = "GHDNA_DATABASE"
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
' |           Notes:  GHDNA DATABASE ENGINE                             |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|

Dim r_filter As String

Private Sub Delete_by_DNA_Click()

    Dim FN As Long
    On Error Resume Next
    Kill App.Path & "\tmp_db.txt"
    
    
    For i = 0 To UBound(m_PAK)
    
        If Text5.Text = m_PAK(i).DNA_sequence Then
        
        Else
        
            If m_PAK(i).DNA_sequence <> "" Then
                FN = FreeFile
                Open App.Path & "\tmp_db.txt" For Append As FN
                Print #FN, m_PAK(i).GHDNA_hash & "::" & m_PAK(i).DNA_sequence
                Close #FN
            End If
        
        End If
    
    Next i
    
    Kill App.Path & "\db.txt"
    Call FileCopy(App.Path & "\tmp_db.txt", App.Path & "\db.txt")
    Read_DB_File m_PAK, App.Path & "\db.txt"
    
End Sub

Private Sub Delete_by_hash_Click()

    Dim FN As Long
    On Error Resume Next
    Kill App.Path & "\tmp_db.txt"
    
    For i = 0 To UBound(m_PAK)
    
        If Text4.Text = m_PAK(i).GHDNA_hash Then
        
        Else
        
            If m_PAK(i).GHDNA_hash <> "" Then
                FN = FreeFile
                Open App.Path & "\tmp_db.txt" For Append As FN
                Print #FN, m_PAK(i).GHDNA_hash & "::" & m_PAK(i).DNA_sequence
                Close #FN
            End If
        
        End If
    
    Next i
    
    
    Kill App.Path & "\db.txt"
    Call FileCopy(App.Path & "\tmp_db.txt", App.Path & "\db.txt")
    Read_DB_File m_PAK, App.Path & "\db.txt"
    
End Sub

Private Sub Form_Load()
    Read_DB_File m_PAK, App.Path & "\db.txt"
End Sub

Private Sub Input_DB_Click()

    Dim FN As Long
    
    tmp_hash = GHDNA_DATA_BLOCK(UCase(Input_seq.Text))
    
    For i = 0 To UBound(m_PAK)
    
        If tmp_hash = m_PAK(i).GHDNA_hash Then
            MsgBox "Sequence [" & Input_seq.Text & "] already exists in GHDNA database !"
            Exit Sub
        End If
    
    Next i
        
    FN = FreeFile
    Open App.Path & "\db.txt" For Append As FN
    Print #FN, tmp_hash & "::" & Input_seq.Text
    Close #FN
        
    Read_DB_File m_PAK, App.Path & "\db.txt"
    
End Sub

Private Sub Search1_Click()

    tmp_hash = GHDNA_DATA_BLOCK(UCase(Text1.Text))
    
    For i = 0 To UBound(m_PAK)

        If tmp_hash = m_PAK(i).GHDNA_hash Then Text3.Text = Text3.Text & "[" & m_PAK(i).DNA_sequence & "] exists in our database, with the hash number: " & m_PAK(i).GHDNA_hash
    
    Next i
    
End Sub


Private Sub Search2_Click()

    Motif_len = 6
    Text3.Text = Text3.Text & vbCrLf & "Total length of input sequence = " & Len(Text2.Text) & vbCrLf
    
    For w = 1 To Len(Text2.Text) - Motif_len
        
        tmp_window = Mid(Text2.Text, w, Motif_len)
        tmp_hash = GHDNA_DATA_BLOCK(UCase(tmp_window))
            
        If Restriction_filter(tmp_window) = True Then
        
            Text3.Text = Text3.Text & "Restriction filter for -[" & tmp_window & "]" & vbCrLf & vbCrLf
            GoTo 1
            
        End If
            
        For i = 0 To UBound(m_PAK)
            
            If tmp_hash = m_PAK(i).GHDNA_hash Then
                
                Text3.Text = Text3.Text & "Motif found at position: " & w & " - " & m_PAK(i).DNA_sequence & vbCrLf
                tmp1 = Mid(Text2.Text, 1, w - 1)
                tmp2 = Mid(Text2.Text, w + Len(tmp_window), Len(Text2.Text) - w + Len(tmp_window))
                Text3.Text = Text3.Text & tmp1 & "-[" & tmp_window & "]-" & tmp2 & vbCrLf & vbCrLf
            
            End If
        
        Next i
    
1:
    Next w

End Sub


Function Restriction_filter(ByVal h As String) As Boolean

    If InStr(r_filter, h) <> 0 Then Restriction_filter = True Else Restriction_filter = False
    r_filter = r_filter & "-" & h
    
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
    m = 10

    Do Until t > 3 And v = 0
        a = a + 1
        t = (L Mod a)
        R = (L - t)
        v = R Mod 2
    Loop


    Do Until b = 0 Or m > 1000
        m = m + 1
        b = R Mod m
    Loop

    Block_Alocation = m
    
End Function
