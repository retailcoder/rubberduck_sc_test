VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlResultRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type tRow
    ParentResult As SqlResult
    values() As Variant
    IsEmpty As Boolean
End Type

Private this As tRow

Private Sub Class_Initialize()
    ReDim this.values(0 To 0)
    this.IsEmpty = True
End Sub

Public Property Get ParentResult() As SqlResult
    Set ParentResult = this.ParentResult
End Property

Public Property Set ParentResult(value As SqlResult)
    Set this.ParentResult = value
End Property

Friend Sub AddValue(ByVal value As Variant)
    
    If Not this.IsEmpty Then ReDim Preserve this.values(0 To UBound(this.values) + 1)
    this.values(UBound(this.values)) = value
    
    this.IsEmpty = False
    
End Sub

Public Property Get FieldCount() As Integer
    FieldCount = UBound(this.values) + Abs(LBound(this.values) - 1)
End Property

Public Property Get Item(ByVal nameOrIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    
    If TypeName(nameOrIndex) = "String" Then
    
        Item = GetFieldValueByName(LCase(nameOrIndex))
    
    ElseIf IsNumeric(nameOrIndex) Then
    
        Item = GetFieldValueByIndex(nameOrIndex)
    
    Else
        
        'return empty variant
        
    End If
    
End Property

Public Property Let Item(ByVal nameOrIndex As Variant, ByVal value As Variant)
    
    If TypeName(nameOrIndex) = "String" Then
    
        SetFieldValueByName LCase(nameOrIndex), value
    
    ElseIf IsNumeric(nameOrIndex) Then
    
        SetFieldValueByIndex nameOrIndex, value
        
    Else
    
        Err.Raise 13 'type mismatch
    
    End If
        
End Property

Private Function GetFieldValueByName(ByVal name As String) As Variant
    If Not this.IsEmpty Then GetFieldValueByName = this.values(this.ParentResult.FieldNameIndex(LCase(name)))
End Function

Private Function GetFieldValueByIndex(ByVal index As Integer) As Variant
    If Not this.IsEmpty Then GetFieldValueByIndex = this.values(index)
End Function

Private Sub SetFieldValueByName(ByVal name As String, ByVal value As Variant)
    If Not this.IsEmpty Then this.values(this.ParentResult.FieldNameIndex(LCase(name))) = value
End Sub

Private Sub SetFieldValueByIndex(ByVal index As Integer, ByVal value As Variant)
    If Not this.IsEmpty Then this.values(index) = value
End Sub

Public Function Create(parent As SqlResult, fields As ADODB.fields) As SqlResultRow
    
    Dim result As New SqlResultRow
    Set result.ParentResult = parent
    
    Dim field As ADODB.field
    Dim value As Variant
    
    For Each field In fields
        If TypeName(field.value) = "String" Then
            value = LTrim(RTrim(Coalesce(field.value, vbNullString)))
        Else
            value = Coalesce(field.value, vbEmpty)
        End If
        result.AddValue value
    Next
    
    Set Create = result
    
End Function

Public Function ToString() As String
        
    If this.IsEmpty Then
        ToString = TypeName(Me)
        Exit Function
    End If
    
    Dim result As String
    result = Join(this.values, this.ParentResult.ValueSeparator)
    
    ToString = result
    
End Function
