VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CompareTo(ByVal other As Variant) As Integer
'return  1 if [other] is greater than this instance;
'return -1 if [other] is smaller than this instance;
'return  0 if [other] is equivalent than this instance.
End Function
