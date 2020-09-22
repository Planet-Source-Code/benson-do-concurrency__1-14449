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
Dim x As Long

Dim daoDb As DAO.Database
Dim daoWs As DAO.Workspace
Dim Ws As DAO.Workspace
Dim daoRs As DAO.Recordset
Dim strSQL As String

Set Ws = DBEngine(0)
Set daoWs = Workspaces(0)
Set daoDb = daoWs.OpenDatabase(App.Path & "\db1.MDB", False, False)
    
strSQL = "SELECT * FROM table1"

Set daoRs = daoDb.OpenRecordset(strSQL, dbOpenDynaset, dbAppendOnly)

For x = 1 To 1000
    DBEngine.Idle dbRefreshCache
        Ws.BeginTrans
            daoRs.AddNew
            daoRs.Fields("col1").Value = Str(x)
            daoRs.Fields("col2").Value = "DAO Concurrent Test 4"
            daoRs.Update
        Ws.CommitTrans dbForceOSFlush
    Me.Label1.Caption = Str(x)
    DoEvents
Next x

Set daoDb = Nothing
Set daoWs = Nothing
Unload Me
End Sub

Private Sub Form_Load()
Me.Show
Command1_Click
End Sub
