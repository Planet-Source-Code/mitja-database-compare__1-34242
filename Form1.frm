VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Create"
      Height          =   345
      Left            =   3585
      TabIndex        =   8
      Top             =   1050
      Width           =   1305
   End
   Begin VB.TextBox Text3 
      Height          =   4140
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form1.frx":0000
      Top             =   1545
      Width           =   6120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compare"
      Height          =   345
      Left            =   1260
      TabIndex        =   6
      Top             =   1065
      Width           =   1305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5160
      TabIndex        =   5
      Top             =   465
      Width           =   405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   105
      Width           =   405
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1605
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   480
      Width           =   3510
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1605
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   105
      Width           =   3510
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   105
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   45
      X2              =   6375
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label2 
      Caption         =   "Database 2:"
      Height          =   255
      Left            =   675
      TabIndex        =   1
      Top             =   525
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Database 1:"
      Height          =   270
      Left            =   675
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   CommonDialog1.Filter = "Microsoft database|*.mdb"
   CommonDialog1.ShowOpen
   Text1 = CommonDialog1.filename
   
End Sub

Private Sub Command2_Click()
   CommonDialog1.Filter = "Microsoft database|*.mdb"
   CommonDialog1.ShowOpen
   Text2 = CommonDialog1.filename

End Sub

Private Sub Command3_Click()
   Dim WrkSpace As Workspace
   Dim Baza1 As Database
   Dim Baza2 As Database
   Dim rec1 As Recordset
   Dim rec2 As Recordset
   Dim TableDfs As TableDef
   
   
   Set WrkSpace = DBEngine.CreateWorkspace("Compare", "Admin", "")
   Set Baza1 = WrkSpace.OpenDatabase(Text1)
   Set Baza2 = WrkSpace.OpenDatabase(Text2)
   On Error Resume Next
   Text3 = ""
   For N = 0 To Baza1.TableDefs.Count - 1
      If Baza1.TableDefs(N).Properties(5) = 0 Then
         DoEvents
         Debug.Print Baza1.TableDefs(N).Name;
         T_Ime = Baza1.TableDefs(N).Name
         Set rec1 = Baza1.OpenRecordset(T_Ime)
         Set rec2 = Baza2.OpenRecordset(T_Ime)
         If Err = 3078 Then
            Debug.Print "   ERR"
            Text3 = Text3 & "Missing table: " & T_Ime & vbNewLine
            Err = 0
         Else
            Debug.Print "   OK"
            For m = 0 To rec1.Fields.Count - 1
               R_Ime = rec1.Fields(m).Name
               Debug.Print "    " & R_Ime;
               nekej = rec2.Fields(R_Ime).Name
               If Err = 0 Then
                  Debug.Print "   OK"
               Else
                  Debug.Print "   ERR"
                  Text3 = Text3 & "Missing field: " & T_Ime & "." & R_Ime & vbNewLine
                  Err = 0
               End If
            Next m
         End If
      End If
   Next N
   If Text3 = "" Then
      Text3 = "No differences found."
   End If
End Sub

Private Sub Command4_Click()
Dim WrkSpace As Workspace
   Dim Baza1 As Database
   Dim Baza2 As Database
   Dim rec1 As Recordset
   Dim rec2 As Recordset
   Dim NewTable As TableDef
   Dim NewFld As Field
   
   Set WrkSpace = DBEngine.CreateWorkspace("Compare", "Admin", "")
   Set Baza1 = WrkSpace.OpenDatabase(Text1)
   Set Baza2 = WrkSpace.OpenDatabase(Text2)
   On Error Resume Next
   Text3 = ""
   For N = 0 To Baza1.TableDefs.Count - 1
      If Baza1.TableDefs(N).Properties(5) = 0 Then
         DoEvents
         Debug.Print Baza1.TableDefs(N).Name;
         T_Ime = Baza1.TableDefs(N).Name
         Set rec1 = Baza1.OpenRecordset(T_Ime)
         Set rec2 = Baza2.OpenRecordset(T_Ime)
         If Err = 3078 Then
            Debug.Print "   ERR"
            Text3 = Text3 & "Missing table: " & T_Ime & vbNewLine
            Set NewTable = Baza2.CreateTableDef(T_Ime)
            With NewTable
               For m = 0 To rec1.Fields.Count - 1
                  R_Ime = rec1.Fields(m).Name
                  R_Type = rec1.Fields(m).Type
                  R_Size = rec1.Fields(m).Size
                  .Fields.Append .CreateField(R_Ime, R_Type, R_Size)
               Next m
            End With
            Baza2.TableDefs.Append NewTable
            Err = 0
         Else
            Debug.Print "   OK"
            For m = 0 To rec1.Fields.Count - 1
               R_Ime = rec1.Fields(m).Name
               Debug.Print "    " & R_Ime;
               nekej = rec2.Fields(R_Ime).Name
               If Err = 0 Then
                  Debug.Print "   OK"
               Else
                  Debug.Print "   ERR"
                  Text3 = Text3 & "Missing field: " & T_Ime & "." & R_Ime & vbNewLine
                  Err = 0
                  On Error GoTo 0
                  Set NewTable = Baza2.TableDefs(T_Ime)
                  rec2.Close
                  With NewTable
                     R_Ime = rec1.Fields(m).Name
                     R_Type = rec1.Fields(m).Type
                     R_Size = rec1.Fields(m).Size
                     .Fields.Append .CreateField(R_Ime, R_Type, R_Size)
                  End With
                  Set rec2 = Baza2.OpenRecordset(T_Ime)
                  'Baza2.TableDefs.Append NewTable
                  On Error Resume Next
                  Err = 0
               End If
            Next m
         End If
      End If
   Next N
   
End Sub
