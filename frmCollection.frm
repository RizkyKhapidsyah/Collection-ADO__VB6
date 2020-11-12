VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDoB 
      Height          =   345
      Left            =   1320
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   1320
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtId 
      Height          =   345
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Add"
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Get Item"
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Read"
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   345
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim o_clsTest As clsTest
    Dim o_colTest As colTest
    Dim iCount As Integer
    Dim iSelected As Integer

Private Sub Command1_Click()
    Call Form_Load
End Sub

Private Sub Command2_Click()
    With MSFlexGrid1
        iSelected = .RowData(.Row)
    End With
        Set o_clsTest = o_colTest.Item(iSelected)
        o_clsTest.Operation = scDelete
        o_colTest.Save
    Call Form_Load
End Sub

Private Sub Command3_Click()
    Command4.Caption = "&Update"
    With MSFlexGrid1
        iSelected = .RowData(.Row)
        Set o_clsTest = o_colTest.Item(iSelected)
        txtId = o_clsTest.Test_id
        txtName = o_clsTest.Test_name
        txtDoB = o_clsTest.Test_dob
        o_clsTest.Operation = scNone
    End With
    Call Form_Load
End Sub

Private Sub Command4_Click()
    If Command4.Caption = "&Add" Then
        Set o_clsTest = New clsTest
        o_clsTest.Test_dob = txtDoB.Text
        o_clsTest.Test_id = txtId.Text
        o_clsTest.Test_name = txtName.Text
        o_clsTest.Operation = scAddnew
        o_colTest.Add o_clsTest
        o_colTest.Save
        Set o_colTest = Nothing
        Set o_colTest = New colTest
        o_colTest.Read
    Else
        With MSFlexGrid1
            iSelected = .RowData(.Row)
        End With
        Set o_clsTest = o_colTest.Item(iSelected)
        o_clsTest.Test_dob = txtDoB.Text
        o_clsTest.Test_name = txtName.Text
        'o_colTest.Read
        
        o_clsTest.Operation = scUpdate
        o_colTest.Save
        Command4.Caption = "&Add"
        
    End If
    Call Form_Load
 End Sub

Private Sub Form_Load()
            
    Set o_colTest = New colTest
    Set o_clsTest = New clsTest
    o_colTest.Read
    
    iCount = 1
    MSFlexGrid1.Rows = 1
    For Each o_clsTest In o_colTest
        With MSFlexGrid1
            .TextMatrix(0, 0) = "ID"
            .TextMatrix(0, 1) = "Name"
            .TextMatrix(0, 2) = "Date"
            .Rows = .Rows + 1
            .TextMatrix(iCount, 0) = o_clsTest.Test_id
            .TextMatrix(iCount, 1) = o_clsTest.Test_name
            .TextMatrix(iCount, 2) = o_clsTest.Test_dob
            .RowData(iCount) = o_clsTest.Test_id
            iCount = iCount + 1
        End With
    Next
End Sub

Private Sub MSFlexGrid1_Click()
    Dim iSelected As Integer
    With MSFlexGrid1
        iSelected = .RowData(.Row)
        Set o_clsTest = o_colTest.Item(iSelected)
        txtId = o_clsTest.Test_id
        txtName = o_clsTest.Test_name
        txtDoB = o_clsTest.Test_dob
    End With
End Sub
