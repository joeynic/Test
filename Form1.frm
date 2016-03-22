VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
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

Private Sub Form_Load()
    Dim conn As New ADODB.connection
    Dim rs As New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source = 'C:\Users\Joe\Documents\Work\VB6\Test\Data\VB6.accdb'"
    conn.CursorLocation = adUseClient
    conn.Open
    
    rs.Open "GetSports", conn, adOpenForwardOnly, adLockOptimistic
    
    Do While Not rs.EOF
        cmbSports.AddItem rs("sportName")
        cmbSports.ItemData(cmbSports.NewIndex) = rs("sportId")
        rs.MoveNext
    Loop
End Sub

Private Sub mnuF1_Click()
    MsgBox ("f1")
    Dim i As Integer
    i = 1
End Sub

