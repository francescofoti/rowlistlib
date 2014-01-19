VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'CRow.bas class module
'A row (CRow) is a line of a list; it's quite like a collection which
'elements are generally arranged like the columns of a list.
'
'This file is part of the DISRowList library for Visual Basic, DISRowList hereafter.
'
'THe DISRowList library is distributed under a dual license. An open source
'version is licensed under the GNU GPL v2 and a commercial,y licensed version
'can be obtained from devinfo.net either as a standalone package or as part
'of our "The 10th SDK" software library.
'
'DISRowList is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'DISRowList is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with DISRowList (license.txt); if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'When       ¦ Version  ¦ Who ¦ What
'-----------+----------+-----+-----------------------------------------------------
'           ¦          ¦     ¦
Option Explicit

Implements IObjectBytes

'class id for implementing IObjectBytes interface
Private Const klCIDRow               As Long = 1020&

Private Type TColDef
  vValue      As Variant
  iDataType   As Integer
  lDataSize   As Long
  lFlags      As Long
  sColName    As String
End Type

Private mfDirty           As Boolean
Private matColDef()       As TColDef
Private malColIndex()     As Long
Private mlColCount        As Long
Private moColIndexMap     As CMapStringToLong 'Store colname-->colindex map

'For iObjectBytes interface
Private Const ksClassVersion As String = "01.00.00"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
              lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Sub Class_Initialize()
  Set moColIndexMap = New CMapStringToLong
  moColIndexMap.Sorted = True
End Sub

Private Sub Class_Terminate()
  Set moColIndexMap = Nothing
End Sub

Public Sub Clear()
  If mlColCount Then
    Erase matColDef
    Erase malColIndex
    mlColCount = 0&
  End If
  moColIndexMap.Clear
  mfDirty = False
End Sub

Public Property Get Dirty() As Boolean
  Dirty = mfDirty
End Property

Public Property Let Dirty(ByVal pfDirty As Boolean)
  mfDirty = pfDirty
End Property

Public Property Get ColCaseSensitive() As Boolean
  ColCaseSensitive = moColIndexMap.CaseSensitive
End Property

Public Property Let ColCaseSensitive(ByVal pfColCaseSensitive As Boolean)
  moColIndexMap.CaseSensitive = pfColCaseSensitive
End Property

Public Property Get ColCount() As Long
  ColCount = mlColCount
End Property

'Define row from a flat array.
'Elements must be grouped 4 by 4 for a column definition, and in this
'order: name, type, size, flags.
Public Sub Define(ParamArray pavDefs() As Variant)
  Dim lLB       As Long
  Dim lUB       As Long
  Dim lCount    As Long
  Dim lIndex    As Long
  Dim sColName  As String
  Dim iDataType As Integer
  Dim lDataSize As Long
  Dim lFlags    As Long
  Dim lCol      As Long
  
  lLB = LBound(pavDefs)
  lUB = UBound(pavDefs)
  'There must be a multiple of 4 element count
  If (lUB - lLB + 1&) Mod 4 Then
    Err.Raise 5&, "CRow::Define", "Incorrect number of array elements."
    Exit Sub
  End If
  lCount = (lUB - lLB + 1&) \ 4
  Clear
  lCol = 1&
  For lIndex = 1& To lCount
    sColName = pavDefs(lLB + (lCol - 1&) * 4&)
    iDataType = pavDefs(lLB + (lCol - 1&) * 4& + 1&)
    lDataSize = pavDefs(lLB + (lCol - 1&) * 4& + 2&)
    lFlags = pavDefs(lLB + (lCol - 1&) * 4& + 3&)
    If moColIndexMap.Find(sColName) = 0 Then
      ReDim Preserve matColDef(1 To mlColCount + 1&)
      ReDim Preserve malColIndex(1 To mlColCount + 1&)
      'Append at end of columns
      mlColCount = mlColCount + 1&
      malColIndex(mlColCount) = mlColCount
      moColIndexMap.Add sColName, mlColCount
      'Add info
      With matColDef(mlColCount)
        .vValue = Null
        .iDataType = iDataType
        .lDataSize = lDataSize
        .lFlags = lFlags
        .sColName = sColName
      End With
    End If
    lCol = lCol + 1&
  Next lIndex
  mfDirty = False 'we've a new row, so its not dirty
End Sub

Public Sub Merge(ByRef prowSource As CRow, Optional ByVal pfAutoAddCols As Boolean = True)
  Dim lColCount     As Long
  Dim iCol          As Long
  Dim sColName      As String
  
  lColCount = prowSource.ColCount
  If lColCount = 0& Then Exit Sub
  
  For iCol = 1& To lColCount
    With prowSource
      sColName = .ColName(iCol)
      'If the column doesn't exist in our set, then
      'we merge the source column by adding it to our set.
      If Len(sColName) Then
        If ColExists(sColName) Then
          'We get the value in the row with which we merge and overwrite ours
          Me.ColValue(sColName) = .ColValue(sColName)
        Else
          If pfAutoAddCols Then
            Me.AddCol sColName, _
                      .ColValue(iCol), _
                      .ColSize(iCol), _
                      .ColFlags(iCol)
          End If
        End If
      End If
    End With
  Next iCol
  mfDirty = True
End Sub

Public Sub DefineRow(ByRef prowDest As CRow)
  prowDest.Clear
  If mlColCount = 0& Then Exit Sub
  
  Dim lDataColIndex     As Long
  Dim lDataRowIndex     As Long
  Dim lColIndex         As Long
  Dim fValidRow         As Boolean
  
  'Define row columns
  ReDim avColName(1 To mlColCount)
  ReDim avColType(1 To mlColCount)
  ReDim avColSize(1 To mlColCount)
  ReDim avColFlags(1 To mlColCount)
  ReDim avValues(1 To mlColCount)
  For lColIndex = 1& To mlColCount
    lDataColIndex = malColIndex(lColIndex)
    With matColDef(lDataColIndex)
      avColName(lColIndex) = .sColName
      avColType(lColIndex) = .iDataType
      avColSize(lColIndex) = .lDataSize
      avColFlags(lColIndex) = .lFlags
      avValues(lColIndex) = Null
      If IsNull(avValues(lColIndex)) Then
        If avColType(lColIndex) = vbString Then
          avValues(lColIndex) = ""
        End If
      End If
    End With
  Next lColIndex
  prowDest.ArrayDefine avColName, avColType, avColSize, avColFlags
  mfDirty = False
End Sub

Public Sub DefineList(ByRef plstDest As CList)
  plstDest.Clear
  If mlColCount = 0& Then Exit Sub
  
  Dim lDataColIndex     As Long
  Dim lDataRowIndex     As Long
  Dim lColIndex         As Long
  Dim fValidRow         As Boolean
  
  'Define row columns
  ReDim avColName(1 To mlColCount)
  ReDim avColType(1 To mlColCount)
  ReDim avColSize(1 To mlColCount)
  ReDim avColFlags(1 To mlColCount)
  For lColIndex = 1& To mlColCount
    lDataColIndex = malColIndex(lColIndex)
    With matColDef(lDataColIndex)
      avColName(lColIndex) = .sColName
      avColType(lColIndex) = .iDataType
      avColSize(lColIndex) = .lDataSize
      avColFlags(lColIndex) = .lFlags
    End With
  Next lColIndex
  plstDest.ArrayDefine avColName, avColType, avColSize, avColFlags
End Sub

Public Sub ArrayDefine(pavColName As Variant, _
                       Optional pavDataType As Variant, _
                       Optional pavDataSize As Variant, _
                       Optional pavDataFlags As Variant)
  Dim lLB       As Long
  Dim lUB       As Long
  Dim lIndex    As Long
  Dim sColName  As String
  Dim iDataType As Integer
  Dim lDataSize As Long
  Dim lFlags    As Long
  
  lLB = LBound(pavColName)
  lUB = UBound(pavColName)
  Clear
  For lIndex = lLB To lUB
    sColName = pavColName(lIndex)
    If Not IsMissing(pavDataType) Then iDataType = pavDataType(lIndex)
    If Not IsMissing(pavDataSize) Then lDataSize = pavDataSize(lIndex)
    If Not IsMissing(pavDataFlags) Then lFlags = pavDataFlags(lIndex)
    If moColIndexMap.Find(sColName) = 0 Then
      ReDim Preserve matColDef(1 To mlColCount + 1&)
      ReDim Preserve malColIndex(1 To mlColCount + 1&)
      'Append at end of columns
      mlColCount = mlColCount + 1&
      malColIndex(mlColCount) = mlColCount
      moColIndexMap.Add sColName, mlColCount
      'Add info
      With matColDef(mlColCount)
        .vValue = Null
        .iDataType = iDataType
        .lDataSize = lDataSize
        .lFlags = lFlags
        .sColName = sColName
      End With
    End If
  Next lIndex
  mfDirty = False
End Sub

Public Sub AddCol(ByRef psColName As String, _
                  ByVal pvColValue As Variant, _
                  ByVal plDataSize As Long, _
                  ByVal plFlags As Long, _
                  Optional ByVal plInsertAfter As Long = 0&, _
                  Optional ByVal plInsertBefore As Long = 0&)
  If Len(psColName) Then
    If moColIndexMap.Find(psColName) Then
      'This key is already associated with an element of this collection
      Err.Raise 457&, "CRow::AddCol", VBA.Error$(457)
      Exit Sub
    End If
  End If
  
  Dim lColIndex   As Long
  Dim i           As Long
  ReDim Preserve matColDef(1 To mlColCount + 1&)
  ReDim Preserve malColIndex(1 To mlColCount + 1&)
  If plInsertAfter Then
    lColIndex = plInsertAfter + 1&
    'Push down other elements
    If lColIndex <= mlColCount Then
      CopyMemory malColIndex(lColIndex + 1&), malColIndex(lColIndex), (mlColCount - plInsertAfter) * LenB(lColIndex)
    End If
    For i = 1 To moColIndexMap.Count
      If moColIndexMap.Item(i) >= lColIndex Then
        moColIndexMap.Item(i) = moColIndexMap.Item(i) + 1&
      End If
    Next i
  ElseIf plInsertBefore Then
    'Push down other elements
    lColIndex = plInsertBefore
    If mlColCount Then
      CopyMemory malColIndex(lColIndex + 1&), malColIndex(lColIndex), (mlColCount - lColIndex + 1&) * LenB(lColIndex)
      For i = 1 To moColIndexMap.Count
        If moColIndexMap.Item(i) >= lColIndex Then
          moColIndexMap.Item(i) = moColIndexMap.Item(i) + 1&
        End If
      Next i
    End If
  Else
    'Append at end of columns
    lColIndex = mlColCount + 1&
  End If
  mlColCount = mlColCount + 1&
  malColIndex(lColIndex) = mlColCount
  If Len(psColName) Then moColIndexMap.Add psColName, lColIndex
  'Add info
  With matColDef(mlColCount)
    .vValue = pvColValue
    .iDataType = VarType(pvColValue)
    .lDataSize = plDataSize
    .lFlags = plFlags
    .sColName = psColName
  End With
  mfDirty = True  'humm, adding a colon he? - That's something that dirties our row
End Sub

Public Sub RemoveCol(ByVal pvColIndex As Variant)
  Dim lIndex      As Long
  Dim lColIndex   As Long
  Dim lValueIndex As Long
  Dim i           As Long
  
  If VarType(pvColIndex) = vbString Then
    'Find column index from name
    If Len(pvColIndex) Then lIndex = moColIndexMap.Find((pvColIndex))
    If lIndex Then
      lColIndex = moColIndexMap.Item(lIndex)
      'Remove from map
      moColIndexMap.Remove lIndex
    Else
      'Error 5 (Invalid procedure call)
      Err.Raise 5&, "CRow::RemoveCol", "<" & pvColIndex & ">: Item not found"
      Exit Sub
    End If
  Else
    lColIndex = pvColIndex
    'Remove from map
    For i = 1 To moColIndexMap.Count
      If moColIndexMap.Item(i) = lColIndex Then
        moColIndexMap.Remove i
        Exit For
      End If
    Next i
  End If
  
  lValueIndex = malColIndex(lColIndex)
  'Remove from value array
  For i = lValueIndex To mlColCount - 1&
    matColDef(i) = matColDef(i + 1&)
  Next i
  'Remove from colindex
  For i = lColIndex To mlColCount - 1&
    malColIndex(i) = malColIndex(i + 1&)
  Next i
  'Eventually an index pointing on the last element must be corrected
  For i = 1 To mlColCount
    If malColIndex(i) > (mlColCount - 1&) Then
      malColIndex(i) = malColIndex(i) - 1&
      Exit For
    End If
  Next i
  'Decrement each column index value in map
  For i = 1 To moColIndexMap.Count
    If moColIndexMap.Item(i) > lColIndex Then
      moColIndexMap.Item(i) = moColIndexMap.Item(i) - 1&
    End If
  Next i
  mlColCount = mlColCount - 1&
  
  mfDirty = True
End Sub

Public Function ColPos(ByVal psColName As String) As Long
  Dim lColIndex As Long
  If Left$(psColName, 1) <> "#" Then
    lColIndex = moColIndexMap.Find(psColName)
    If lColIndex = 0& Then
      'Error 5 (Invalid procedure call)
      Err.Raise 5&, "CRow::Value [Get]", "<" & psColName & ">: Item not found"
      Exit Function
    End If
    ColPos = moColIndexMap.Item(lColIndex)
  Else
    ColPos = CLng(Val(Right$(psColName, Len(psColName) - 1)))
  End If
End Function

Public Function ColExists(ByVal psColName As String) As Boolean
  ColExists = CBool(moColIndexMap.Find(psColName))
End Function

Public Property Get ColValue(ByVal pvIndex As Variant) As Variant
Attribute ColValue.VB_UserMemId = 0
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  If Not IsObject(matColDef(malColIndex(lColIndex)).vValue) Then
    ColValue = matColDef(malColIndex(lColIndex)).vValue
  Else
    Set ColValue = matColDef(malColIndex(lColIndex)).vValue
  End If
End Property

Public Property Let ColValue(ByVal pvIndex As Variant, ByVal pvNewValue As Variant)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  matColDef(malColIndex(lColIndex)).vValue = pvNewValue
  mfDirty = True
End Property

Public Property Set ColValue(ByVal pvIndex As Variant, ByRef pvNewValue As Variant)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  Set matColDef(malColIndex(lColIndex)).vValue = pvNewValue
  mfDirty = True
End Property

Public Property Get ColName(ByVal plColIndex As Long) As String
  ColName = matColDef(malColIndex(plColIndex)).sColName
End Property

Public Property Let ColName(ByVal plColIndex As Long, ByVal psNewName As String)
  Dim lMappedLong   As Long
  Dim lMapIndex     As Long
  Dim i             As Long
  
  'Column name must be given, you can't make a named column an unnamed one.
  If Len(psNewName) = 0 Then
    Err.Raise 5&, "CRow::ColName [Let]", "Missing column name."
    Exit Property
  End If
  If moColIndexMap.Find(psNewName) Then
    'This key is already associated with an element of this collection
    Err.Raise 457&, "CRow::ColName [Let]", "Duplicate column names not allowed"
    Exit Property
  End If
  'Sequentially search for the index in the map
  For i = 1 To moColIndexMap.Count
    If moColIndexMap.Item(i) = plColIndex Then
      lMapIndex = i
      Exit For
    End If
  Next i
  If lMapIndex Then
    moColIndexMap.Remove lMapIndex
    moColIndexMap.Add psNewName, plColIndex
    matColDef(malColIndex(plColIndex)).sColName = psNewName
  Else
    'Bad column index, error 9: "Subscript out of range"
    Err.Raise 9&, "CRow::ColName [Let]", VBA.Error$(9&)
  End If
End Property

Public Property Get ColType(ByVal pvIndex As Variant) As Integer
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  ColType = matColDef(malColIndex(lColIndex)).iDataType
End Property

Public Property Let ColType(ByVal pvIndex As Variant, ByVal piNewType As Integer)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  matColDef(malColIndex(lColIndex)).iDataType = piNewType
End Property

Public Property Get ColSize(ByVal pvIndex As Variant) As Long
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  ColSize = matColDef(malColIndex(lColIndex)).lDataSize
End Property

Public Property Let ColSize(ByVal pvIndex As Variant, ByVal plNewSize As Long)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  matColDef(malColIndex(lColIndex)).lDataSize = plNewSize
End Property

Public Property Get ColFlags(ByVal pvIndex As Variant) As Long
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  ColFlags = matColDef(malColIndex(lColIndex)).lFlags
End Property

Public Property Let ColFlags(ByVal pvIndex As Variant, ByVal plNewFlags As Long)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  matColDef(malColIndex(lColIndex)).lFlags = plNewFlags
End Property

Public Sub Assign(ParamArray pavValues() As Variant)
  Dim lLB       As Long
  Dim lUB       As Long
  Dim lIndex    As Long
  Dim lListCol  As Long
  
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    matColDef(malColIndex(lListCol)).vValue = pavValues(lIndex)
    lListCol = lListCol + 1&
  Next lIndex
  mfDirty = True
End Sub

Public Sub ArrayAssign(ByRef pavValues As Variant)
  Dim lLB       As Long
  Dim lUB       As Long
  Dim lIndex    As Long
  Dim lListCol  As Long
  
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    matColDef(malColIndex(lListCol)).vValue = pavValues(lIndex)
    lListCol = lListCol + 1&
  Next lIndex
  mfDirty = True
End Sub

Public Function Clone() As CRow
  Dim oClone    As CRow
  
  Set oClone = New CRow
  oClone.CopyFrom Me
  Set Clone = oClone
  Set oClone = Nothing
End Function

Public Sub CopyFrom(ByRef prowSource As CRow)
  Dim lColCount     As Long
  Dim iCol          As Long
  
  'Does the same as DefineRow, but using another technique
  Clear
  lColCount = prowSource.ColCount
  If lColCount = 0& Then Exit Sub
  
  For iCol = 1& To lColCount
    With prowSource
      Me.AddCol .ColName(iCol), _
                .ColValue(iCol), _
                .ColSize(iCol), _
                .ColFlags(iCol)
    End With
  Next iCol
  mfDirty = False 'we've got a completely new row
End Sub

'
' IObjectBytes implementation
'

Private Property Get IObjectBytes_ByteSize() As Long
  IObjectBytes_ByteSize = GetInstanceByteSize()
End Property

Private Function GetInstanceByteSize() As Long
  Dim lsize       As Long
  Dim i           As Long
  
  'We store class id and class version (long + long)
  lsize = 2& * klSizeOfLong
  
  'Behaviour
  lsize = lsize + klSizeOfBool  'mfDirty
  lsize = lsize + klSizeOfLong  'mlColCount
  
  'Column indices (malColIndex())
  lsize = lsize + mlColCount * klSizeOfLong 'malColIndex items
  'Column items (TColDef items in matColDef())
  For i = 1 To mlColCount
    With matColDef(i)
      lsize = lsize + CalcByteSize(.vValue)
      lsize = lsize + klSizeOfInt   '.iDataType
      lsize = lsize + klSizeOfLong  '.lDataSize
      lsize = lsize + klSizeOfLong  '.lFlags
      lsize = lsize + klSizeOfLong + Len(.sColName) '.sColName
    End With
  Next i
  lsize = lsize + moColIndexMap.IIObjectBytes.ByteSize()
  
  GetInstanceByteSize = lsize
End Function

'Easily expose implemented IObjectBytes interface
Public Property Get IIObjectBytes() As IObjectBytes
  Set IIObjectBytes = Me
End Property

Private Sub IObjectBytes_GetDataBytes(abRetObject() As Byte)
  'It's our responsability to allocate space (ie redim abRetObject)
  Dim lByte         As Long
  Dim lByteSize     As Long
  Dim i             As Long
  Dim iIndex        As Long
  
  lByteSize = GetInstanceByteSize()
  
  On Error Resume Next
  Erase abRetObject
  On Error GoTo 0
  If lByteSize = 0& Then Exit Sub
  
  'transform instance data to byte array
  lByte = 1&
  ReDim abRetObject(1 To lByteSize)
  
  'Write class id and class version
  PokeLong klCIDRow, abRetObject, lByte
  PokeLong MAKE_VERSIONLONG(ksClassVersion), abRetObject, lByte
  
  'Write behaviour members
  PokeBoolean mfDirty, abRetObject, lByte
  PokeLong mlColCount, abRetObject, lByte
  
  'Column items (TColDef items in matColDef())
  For i = 1 To mlColCount
    PokeLong malColIndex(i), abRetObject, lByte
    With matColDef(i)
      MoveVariant abRetObject, lByte, .vValue
      PokeInteger .iDataType, abRetObject, lByte
      PokeLong .lDataSize, abRetObject, lByte
      PokeLong .lFlags, abRetObject, lByte
      PokeString .sColName, abRetObject, lByte
    End With
  Next i
  
  'The index map must be first saved in an independent byte array,
  'then we can copy the indep byte array into our destination byte array.
  Dim abMapBytes() As Byte
  moColIndexMap.IIObjectBytes.GetDataBytes abMapBytes()
  CopyMemory abRetObject(lByte), abMapBytes(1), UBound(abMapBytes)
  lByte = lByte + UBound(abMapBytes)
  Erase abMapBytes
End Sub

Private Sub IObjectBytes_SetDataBytes(abObjectData() As Byte, Optional plByte As Long = 1&)
  Dim i             As Long
  Dim iIndex        As Long
  Dim lClassID      As Long
  Dim lClassVer     As Long
  
  Clear 'load all defaults
  
  'ClassID and version
  lClassID = PeekLong(abObjectData(), plByte)
  lClassVer = PeekLong(abObjectData, plByte)
  If lClassID <> klCIDRow Then
    Err.Raise MAKE_OBJECTERROR(kErrBadClassIDBytes), "CRow", "CRow data bytes class ID doesn't match"
  End If
  If lClassVer <> MAKE_VERSIONLONG(ksClassVersion) Then
    Err.Raise MAKE_OBJECTERROR(kErrBadClassVerBytes), "CRow", "CRow data bytes version doesn't match"
  End If
  
  'behaviour members
  mfDirty = PeekBoolean(abObjectData(), plByte)
  mlColCount = PeekLong(abObjectData(), plByte)
  If mlColCount Then
    ReDim malColIndex(1 To mlColCount)
    ReDim matColDef(1 To mlColCount)
    For i = 1 To mlColCount
      malColIndex(i) = PeekLong(abObjectData(), plByte)
      With matColDef(i)
        GetVariant .vValue, plByte, abObjectData()
        .iDataType = PeekInteger(abObjectData(), plByte)
        .lDataSize = PeekLong(abObjectData(), plByte)
        .lFlags = PeekLong(abObjectData(), plByte)
        .sColName = PeekString(abObjectData(), plByte)
      End With
    Next i
  End If
  
  moColIndexMap.IIObjectBytes.SetDataBytes abObjectData, plByte
End Sub


