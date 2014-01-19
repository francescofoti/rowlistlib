Attribute VB_Name = "MMain"
Option Explicit

Sub Main()

End Sub

Public Sub TestCRow()
  Dim oRow    As CRow
  
  On Error GoTo TestCRow_Err
  
  Set oRow = New CRow
  'Define: name, type, size, flags
  oRow.Define "String", vbString, 0&, 0&, _
              "Boolean", vbBoolean, 0&, 0&
  oRow.AddCol "Date      ", Date, 0&, 0&
  RowDump oRow
  
  Set oRow = Nothing
  Exit Sub
TestCRow_Err:
  Debug.Print "TestCRow, Error #" & Err.Number & ": " & Err.Description
  Set oRow = Nothing
End Sub

Public Sub RowDump(poRow As CRow)
  Dim iRow      As Long
  Dim i         As Long
  Dim lCount    As Long
  Dim asColName()  As String
  Dim iLen      As Integer
  
  lCount = poRow.ColCount
  'col titles row sep
  For i = 1 To lCount
    iLen = Len(poRow.ColName(i))
    Debug.Print String$(iLen, "-"); "+";
  Next i
  Debug.Print
  For i = 1 To lCount
    Debug.Print poRow.ColName(i); "|";
  Next i
  Debug.Print
  'col titles row sep
  For i = 1 To lCount
    iLen = Len(poRow.ColName(i))
    Debug.Print String$(iLen, "-"); "+";
  Next i
  Debug.Print
  
  'dump values
  For i = 1 To poRow.ColCount
    iLen = Len(poRow.ColName(i))
    Debug.Print StrBlock(poRow.ColValue(i) & "", " ", iLen); "|";
  Next i
  Debug.Print
End Sub

Public Sub ListDump(oList As CList, Optional ByVal sTitle As String = "")
  Dim iRow      As Long
  Dim i         As Long
  Dim lCount    As Long
  Dim asColName()  As String
  Dim iLen      As Integer
  
  lCount = oList.ColCount
  If Len(sTitle) Then
    Debug.Print String$(Len(sTitle), "-"); "+"
    Debug.Print sTitle; "|"
  End If
  'col titles row sep
  For i = 1 To lCount
    iLen = Len(oList.ColName(i))
    Debug.Print String$(iLen, "-"); "+";
  Next i
  Debug.Print
  For i = 1 To lCount
    Debug.Print oList.ColName(i); "|";
  Next i
  Debug.Print
  'col titles row sep
  For i = 1 To lCount
    iLen = Len(oList.ColName(i))
    Debug.Print String$(iLen, "-"); "+";
  Next i
  Debug.Print
  
  'dump values
  lCount = oList.Count
  For iRow = 1 To lCount
    For i = 1 To oList.ColCount
      iLen = Len(oList.ColName(i))
      Debug.Print StrBlock(oList.Item(i, iRow) & "", " ", iLen); "|";
    Next i
    Debug.Print
  Next iRow
End Sub

