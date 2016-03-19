VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEquatable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Equals(ByVal other As Variant) As Boolean
'return True if [other] can be considered equal to this instance.
    Equals = (ObjPtr(Me) = ObjPtr(other))
End Function
