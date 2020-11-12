VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mTest_id As Integer
Private mTest_name As String
Private mTest_dob As Date
Private mOperation As Operation

Public Property Get Test_id() As Variant
    Test_id = mTest_id
End Property

Public Property Let Test_id(ByVal vRHS As Variant)
    mTest_id = vRHS
End Property

Public Property Get Test_name() As Variant
    Test_name = mTest_name
End Property

Public Property Let Test_name(ByVal vRHS As Variant)
    mTest_name = vRHS
End Property

Public Property Get Test_dob() As Variant
    Test_dob = mTest_dob
End Property

Public Property Let Test_dob(ByVal vRHS As Variant)
    mTest_dob = vRHS
End Property

Public Property Get Operation() As Variant
    Operation = mOperation
End Property

Public Property Let Operation(ByVal vRHS As Variant)
    mOperation = vRHS
End Property
