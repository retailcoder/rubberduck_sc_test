VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Gets an enumerator."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'Gets an enumerator.
'This property should have a procedure attribute of -4 to allow usage in For Each loop constructs
End Property
