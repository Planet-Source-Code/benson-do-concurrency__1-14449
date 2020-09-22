VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x, ff As Long

Dim rs As ADODB.Recordset

Dim cnn As New ADODB.Connection
   cnn.Mode = adModeShareDenyNone
   cnn.CursorLocation = adUseClient
   cnn.Provider = "Microsoft.Jet.OLEDB.4.0;"
   cnn.Open Trim(App.Path) & "\db1.MDB"
   ff = cnn.Properties("Jet OLEDB:Transaction Commit Mode")
   cnn.Properties("Jet OLEDB:Page Timeout") = 4000
   cnn.Properties("Jet OLEDB:Transaction Commit Mode") = 1
   cnn.Properties("Jet OLEDB:Lock Delay") = 120 + Int(Rnd * 80)

Set rs = New ADODB.Recordset

rs.LockType = adLockOptimistic
rs.Open "SELECT * FROM table1", cnn, adOpenKeyset, adLockOptimistic
 
    For x = 1 To 5000
            cnn.BeginTrans
                rs.AddNew
                rs.Fields("col1").Value = Str(x)
                rs.Fields("col2").Value = "ADO Concurrent Test 1"
                rs.Update
            cnn.CommitTrans
    Me.Label1.Caption = Str(x)
    DoEvents
    Next x

cnn.Properties("Jet OLEDB:Transaction Commit Mode") = ff
rs.Close
cnn.Close
Unload Me
End Sub

Private Sub Form_Load()
    Me.Show
    Command1_Click
End Sub
