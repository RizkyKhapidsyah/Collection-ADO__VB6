VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oCnn As ADODB.Connection
Dim oRs As Recordset
Dim o_clsTest As clsTest
Dim o_colTest As Collection

Public Function Read()
    Dim sSql As String
    Set oCnn = New ADODB.Connection
    Set oRs = New ADODB.Recordset
    
    oCnn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " + App.Path + "\test.mdb"
    sSql = "select * from tblTest"
    oRs.Open sSql, oCnn, adOpenKeyset, adLockOptimistic
    While Not oRs.EOF
        Set o_clsTest = New clsTest
        o_clsTest.Test_id = oRs.Fields!Test_id
        o_clsTest.Test_name = oRs.Fields!Test_name
        o_clsTest.Test_dob = oRs.Fields!Test_dob
        o_clsTest.Operation = scNone
        Me.Add o_clsTest
        oRs.MoveNext
        
        Set o_clsTest = Nothing
    Wend
    Set oRs = Nothing
End Function

Public Function Add(obj As Object)
    o_colTest.Add obj, CStr(obj.Test_id)
End Function

Public Function Count()
    Count = o_colTest.Count
End Function


Private Sub Class_Initialize()
    Set o_colTest = New Collection
End Sub

Private Sub Class_Terminate()
    Set o_colTest = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
     Set NewEnum = o_colTest.[_NewEnum]
End Property

Public Property Let NewEnum(ByVal vNewValue As Variant)

End Property

Public Function Item(lkey As Variant) As Object
    Set Item = o_colTest.Item(CStr(lkey))
End Function

Public Function Remove(lkey As Variant)
   o_colTest.Remove lkey
End Function

Public Sub Save()
    Dim oSql As String
    Set oRs = New ADODB.Recordset
    
    For Each o_clsTest In Me
        
     
            If o_clsTest.Operation = scAddnew Then
                oSql = "insert into tblTest (test_id,test_name,test_dob) values (" _
                    & o_clsTest.Test_id & ",'" & o_clsTest.Test_name & "','" & o_clsTest.Test_dob & "')"
                oRs.Open oSql, oCnn, adOpenDynamic, adLockOptimistic
            End If
            
            If o_clsTest.Operation = scDelete Then
                oSql = "delete from tblTest where test_id = " & o_clsTest.Test_id
                
                oRs.Open oSql, oCnn, adOpenDynamic, adLockOptimistic
            End If
            
            If o_clsTest.Operation = scUpdate Then
                oSql = "update tblTest set test_name = '" & o_clsTest.Test_name & "',test_dob = '" & o_clsTest.Test_dob & "' Where Test_id = " & o_clsTest.Test_id
                
                oRs.Open oSql, oCnn, adOpenDynamic, adLockOptimistic
            End If
                
'
    Next
    Set o_colTest = Nothing

End Sub
