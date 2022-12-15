Attribute VB_Name = "DB_LODAD"
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
' |           Notes:  GHDNA DATABASE ENGINE                             |
' |                                                                     |
' |                  _____________________________                      |
' |_________________/                             \_____________________|
'

Public Type DB_ITEM
  GHDNA_hash As String
  DNA_sequence As String
End Type

Public m_PAK() As DB_ITEM

Public Function Read_DB_File(R() As DB_ITEM, ByVal Filepath As String)
On Error Resume Next
Dim f%, S$, i&
Dim x0&, x1&, X2&
Dim tmp_sigZ() As String
  On Error GoTo ErrHandler
  
  f = FreeFile()
  Open Filepath For Input As #f
  
  i = 0
  ReDim R(i) As DB_ITEM
  
  Do Until EOF(f)
  
    Line Input #f, S
      If S <> "" Then
        i = i + 1
        ReDim Preserve R(i) As DB_ITEM
          With R(i)
           .GHDNA_hash = Split(S, "::")(0)
           .DNA_sequence = Split(S, "::")(1)
          End With
      End If
  Loop
  
  Close #f

Exit Function
ErrHandler:
  Err.Clear
  Debug.Print Err.Description & "Database information - NOT loaded !!"
End Function
