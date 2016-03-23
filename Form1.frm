VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReadFile 
      Caption         =   "Read File"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateFile 
      Caption         =   "Create File"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton F4 
      Caption         =   "F4"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton F3 
      Caption         =   "F3"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.ComboBox cmbSports 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Begin VB.Menu mnuF1 
         Caption         =   "f1 test"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuTest2 
      Caption         =   "Test 2"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSports_Click()
    MsgBox (cmbSports.ItemData(cmbSports.ListIndex))
End Sub

Private Sub cmdCreateFile_Click()
    Dim intMsg As String
    Dim StudentName As String
    
    Open "C:\Users\Joe\Documents\Work\VB6\Test\Data\sample.txt" For Output As #1
    intMsg = MsgBox("File sample.txt opened")
    StudentName = InputBox("Enter the student Name")
    Print #1, StudentName
    intMsg = MsgBox("Writing a " & StudentName & " to sample.txt")
    
    Close #1
    intMsg = MsgBox("File sample.txt closed")
End Sub

Private Sub cmdReadFile_Click()
    On Error GoTo error_handler
    
    Dim variable1 As String
    Open "C:\Users\Joe\Documents\Work\VB6\Test\Data\sample.txt" For Input As #1
    Input #1, variable1
    Close #1
    MsgBox (variable1)
    
    Exit Sub
error_handler:
    MsgBox ("Error")
    MsgBox (Err.Description)
End Sub

Private Sub F3_Click()
    Form_KeyDown 114, 0
End Sub

Private Sub F4_Click()
    Form_KeyDown 115, 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 114
            RunF3
        Case 115
            RunF4
        Case Else
            'TODO
    End Select
End Sub

Private Sub RunF3()
    MsgBox ("F3")
End Sub

Private Sub RunF4()
    MsgBox ("F4")
End Sub

Private Sub Form_Load()
    Dim s As New Sport
    s.SportName = "123"
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = 'C:\Users\Joe\Documents\Work\VB6\Test\Data\VB6.accdb'"
    conn.CursorLocation = adUseClient
    conn.Open
    
    rs.Open "GetSports", conn, adOpenForwardOnly, adLockOptimistic
    
    Dim arr
    If Not rs.BOF And Not rs.EOF Then
        arr = rs.GetRows
    End If
    rs.Close
    Set rs = Nothing
    
    Dim i As Integer
    For i = 0 To UBound(arr, 2)
        cmbSports.AddItem arr(1, i)
    Next
    
'    Do While Not rs.EOF
'        cmbSports.AddItem rs("sportName")
'        cmbSports.ItemData(cmbSports.NewIndex) = rs("sportId")
'        rs.MoveNext
'    Loop
End Sub

Private Sub mnuF1_Click()
    MsgBox ("f1")
End Sub

