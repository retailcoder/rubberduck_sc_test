VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type tSqlResult
    FieldNames As List
    values As List
    ToStringValueSeparator As String
End Type

Private nameIndices As New Dictionary

Private this As tSqlResult
Implements IEnumerable

Private Sub Class_Initialize()
    Set this.FieldNames = New List
    Set this.values = New List
    this.ToStringValueSeparator = ","
End Sub

Public Property Get ValueSeparator() As String
    ValueSeparator = this.ToStringValueSeparator
End Property

Public Property Let ValueSeparator(ByVal value As String)
    this.ToStringValueSeparator = value
End Property

Public Sub AddFieldName(ByVal name As String)
    this.FieldNames.Add LCase(name)
    nameIndices.Add LCase(name), this.FieldNames.Count - 1
End Sub

Public Property Get FieldNames() As List
    Dim result As New List
    If this.FieldNames.Count > 0 Then result.AddList this.FieldNames
    Set FieldNames = result
End Property

Public Function FieldNameIndex(ByVal name As String) As Long
    
    Dim i As Long
    If nameIndices.Exists(LCase(name)) Then
        FieldNameIndex = nameIndices.Item(LCase(name))
    Else
        FieldNameIndex = -1
    End If
    
End Function

Public Sub AddValue(value As SqlResultRow)
    this.values.Add value
End Sub

Public Property Get Count() As Long
    Count = this.values.Count
End Property

Public Function IndexOf(ByVal value As SqlResultRow) As Long
    IndexOf = this.values.IndexOf(value)
End Function

Public Property Get Item(ByVal index As Long) As SqlResultRow
Attribute Item.VB_UserMemId = 0
    Set Item = this.values(index + 1)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'Gets an enumerator that iterates through the List.
    
    Set NewEnum = this.values.NewEnum

End Property

Public Function First() As SqlResultRow
    Set First = Item(0)
End Function

Public Function Create(adoRecordset As ADODB.Recordset) As SqlResult
    
    Dim result As New SqlResult
    
    Dim grabFieldNames As Boolean
    grabFieldNames = True
    
    Dim rowFactory As New SqlResultRow
    Dim field As ADODB.field
    
    While Not adoRecordset.BOF And Not adoRecordset.EOF
        
        For Each field In adoRecordset.fields
            If grabFieldNames Then result.AddFieldName LCase$(Coalesce(field.name, vbNullString))
        Next
        
        result.AddValue rowFactory.Create(result, adoRecordset.fields)
        
        grabFieldNames = False
        adoRecordset.MoveNext
    Wend
    
    Set Create = result
    
End Function

Private Property Get IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = NewEnum
End Property

Public Function WhereFieldEquals(ByVal nameOrIndex As Variant, ByVal value As Variant) As SqlResult
    
    Dim result As New SqlResult
    Dim index As Integer
    For index = 1 To FieldNames.Count
        result.AddFieldName FieldNames(index)
    Next
    
    Dim row As SqlResultRow
    For Each row In this.values
    
        Set row.ParentResult = result
        If row(nameOrIndex) = value Then
            result.AddValue row
        End If
        
    Next
    
    Set WhereFieldEquals = result

End Function
