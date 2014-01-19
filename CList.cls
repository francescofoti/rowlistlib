VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'CList.bas class module
'This the CList class, that handle two dimensional indexed arrays which
'can be searched and sorted on multiple colums.
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
Private Const klCIDList               As Long = 1010&

Private Const kiDefaultGrowSize As Integer = 20
Private mlGrowSize          As Long

'We expose publicly the inner data array, to be able to
'let ADO populate it automatically with a GetRows() call.
Public DataArray              As Variant  '1st dim: Columns, 2nd dim: rows
Private mlDataArraySize       As Long     'Size of 2nd dim (rows) of DataArray
Private mlDataArraySlotCount  As Long     'Number of used rows (some may have been freed)

'current row index
Private mlCurRow              As Long

'Row handles storage
Private malIndex()          As Long
Private mlIdxArraySize      As Long     'Size of index array (upper bound)
Private mlIdxCount          As Long     'Number of elements in the DataArray

'For sorting the list on one or multiple columns
Private Type TSortInfo
  sColName        As String
  lColIndex       As Long
  fCaseSensitive  As Boolean
  fDesc           As Boolean
End Type
Private mfSorted            As Boolean  'If true, the class keeps the list elements sorted
Private msSortColumns       As String     'The sort definitions string ( [!]<colname>[+|-] )
Private matSortInfo()       As TSortInfo  'Information for each column on which we sort
Private miSortColCt         As Integer    'The number of columns on which we sort
'For searching the list
Enum eSearchListAlgorithm
  eSearchAlgoSequential
  eSearchAlgoDichotomic
End Enum
'[f]ind[i]nfo [f]lag values
Private Const klfifCaseSensitive    As Long = 1&  'applies only to text fields.
Private Const klfifHasJokers        As Long = 2&
Private Const klfifHasWildcards     As Long = 4&
Private Const klfifExactSearch      As Long = 8&
Private Const klfifRootSearch       As Long = 16&
Private Const klfifSortedDesc       As Long = 32&
Private Type TFindInfo
  fUseSortKeys      As Boolean  'True if we can benefit from a sorted list, because we search on key fields
  lSearchFieldsCt   As Long     'Number of fields on which we search
  asSearchField()   As String   '1 to lSearchFieldsCt
  alColIndex()      As Long     '1 to lSearchFieldsCt
  avSearchValue()   As Variant  '1 to lSearchFieldsCt
  alFlags()         As Long     '1 to lSearchFieldsCt
  eSearchAlgo       As eSearchListAlgorithm
End Type
Private mtFindInfo  As TFindInfo

'The garbage queue for the rows array (circular array queue)
Private malGarbageQ()       As Long
Private mlGarbQSize         As Long
Private mlGarbQHead         As Long
Private mlGarbQTail         As Long
Private mlGarbQCount        As Long

'
' Columns
'
Private Type TColDef
  iDataType   As Integer
  lDataSize   As Long
  lFlags      As Long
  sColName    As String
End Type
Private matColDef()       As TColDef
Private malColIndex()     As Long
Private mlColCount        As Long
Private moColIndexMap     As CMapStringToLong 'Store colname-->colindex map

'Special column flags
Private Const klColFlagSortedCaseSensitive     As Long = &H10000
Private Const klColFlagSortedCaseInsensitive   As Long = &H20000

'For iObjectBytes interface
Private Const ksClassVersion As String = "01.10.00"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
              lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'
' Constructor / destructor
'
Private Sub Class_Initialize()
  mlGrowSize = kiDefaultGrowSize
  mfSorted = False
  'columns
  Set moColIndexMap = New CMapStringToLong  'case sensitive=false, by default
  moColIndexMap.Sorted = True
  Clear
End Sub

Private Sub Class_Terminate()
  Set moColIndexMap = Nothing
End Sub

'Clear removes the rows *AND* the column definitions.
'Use .Reset to remove the rows only
Public Sub Clear()
  Reset
  'clear sort info
  If miSortColCt Then Erase matSortInfo
  miSortColCt = 0
  mlCurRow = 0&
  'clear columns info
  If mlColCount Then
    Erase matColDef
    Erase malColIndex
    mlColCount = 0&
  End If
  moColIndexMap.Clear
End Sub

Public Sub Reset()
  If mlIdxCount Then Erase DataArray
  mlDataArraySlotCount = 0&
  mlDataArraySize = 0&
  'Clear index array
  mlIdxArraySize = mlGrowSize
  ReDim malIndex(1 To mlIdxArraySize)
  mlIdxCount = 0&
  mlCurRow = 0&
  'Clear garbage queue
  mlGarbQSize = mlGrowSize
  ReDim malGarbageQ(1 To mlGarbQSize)
  mlGarbQHead = 1&
  mlGarbQTail = 0&
  mlGarbQCount = 0&
  'clear special column flags
  ClearAllSysColFlags
End Sub

Public Property Get ColCaseSensitive() As Boolean
  ColCaseSensitive = moColIndexMap.CaseSensitive
End Property

Public Property Let ColCaseSensitive(ByVal pfColCaseSensitive As Boolean)
  moColIndexMap.CaseSensitive = pfColCaseSensitive
End Property

Public Sub SyncWithDataArray()
  Dim i     As Long
  On Error Resume Next
  mlIdxCount = UBound(DataArray, 2)
  If Err.Number = 0 Then
    mlIdxCount = mlIdxCount - LBound(DataArray) + 1&
  End If
  On Error GoTo 0
  If mlIdxCount Then
    If mlIdxCount > mlGrowSize Then
      mlIdxArraySize = ((mlIdxCount \ mlGrowSize) * mlGrowSize) + (mlIdxCount Mod mlGrowSize)
    Else
      mlIdxArraySize = mlGrowSize
    End If
    ReDim malIndex(1 To mlIdxArraySize)
    'For i = 1& To mlIdxCount
    For i = 1& To mlIdxCount
      malIndex(i) = i - 1&
    Next i
    mlDataArraySize = mlIdxCount
    mlDataArraySlotCount = mlIdxCount
  Else
    mlDataArraySize = 0&
    mlDataArraySlotCount = 0&
    mlIdxArraySize = mlGrowSize
    ReDim malIndex(1 To mlIdxArraySize)
  End If
  ClearAllSysColFlags
  mlCurRow = 0&
End Sub

Public Property Get ColCount() As Long
  ColCount = mlColCount
End Property

Public Sub AddCol(ByRef psColName As String, _
                  ByVal pvTemplateValue As Variant, _
                  ByVal plDataSize As Long, _
                  ByVal plFlags As Long, _
                  Optional ByVal plInsertAfter As Long = 0&, _
                  Optional ByVal plInsertBefore As Long = 0&)
  If Len(psColName) Then
    If moColIndexMap.Find(psColName) Then
      'This key is already associated with an element of this collection
      Err.Raise 457&, "CRow::AddCol", "Duplicate column names not allowed"
      Exit Sub
    End If
  End If
  
  'All column contents lost, because we modify the first dim
  If mlColCount Then
    ReDim DataArray(0 To mlColCount, 0 To 0)
  Else
    ReDim DataArray(0 To 0, 0 To 0)
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
      For i = 1 To moColIndexMap.Count
        If moColIndexMap.Item(i) >= lColIndex Then
          moColIndexMap.Item(i) = moColIndexMap.Item(i) + 1&
        End If
      Next i
    End If
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
    .iDataType = VarType(pvTemplateValue)
    .lDataSize = plDataSize
    .lFlags = plFlags And &HFFFF&
    .sColName = psColName
  End With
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

Public Property Get ColName(ByVal plColIndex As Long) As String
  ColName = matColDef(malColIndex(plColIndex)).sColName
End Property

Public Property Let ColName(ByVal plColIndex As Long, ByVal psNewName As String)
  Dim lMappedLong   As Long
  Dim lMapIndex     As Long
  Dim i             As Long
  
  'Column name must be given
  If Len(psNewName) = 0 Then
    Err.Raise 5&, "CRow::ColName [Let]", "Missing column name."
    Exit Property
  End If
  If moColIndexMap.Find(psNewName) Then
    'This key is already associated with an element of this collection
    Err.Raise 457&, "CRow::ColName [Let]", "Duplicate column names not allowed."
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
  'bit values above &HFFFF& are reserved for class internal use,
  'so we just mask the passed flags with &HFFFF&
  matColDef(malColIndex(lColIndex)).lFlags = plNewFlags And &HFFFF&
End Property

'Define list from a flat array.
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
    Err.Raise 5&, "CList::Define", "Incorrect number of array elements."
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
        .iDataType = iDataType
        .lDataSize = lDataSize
        .lFlags = lFlags And &HFFFF&
        .sColName = sColName
      End With
    End If
    lCol = lCol + 1&
  Next lIndex
  ReDim DataArray(0 To mlColCount - 1, 0 To 0)
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
    If Not IsMissing(pavDataType) Then
      iDataType = pavDataType(lIndex)
    Else
      iDataType = VbVarType.vbString
    End If
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
        .iDataType = iDataType
        .lDataSize = lDataSize
        .lFlags = lFlags And &HFFFF&
        .sColName = sColName
      End With
    End If
  Next lIndex
  ReDim DataArray(0 To mlColCount - 1, 0 To 0)
End Sub

'
' CList specific methods
'

Public Property Get CurrentRow() As Long
  CurrentRow = mlCurRow
End Property

Public Property Let CurrentRow(ByVal plNewCurRow As Long)
  If (plNewCurRow >= 0) And (plNewCurRow <= Me.Count) Then
    mlCurRow = plNewCurRow
  End If
End Property

Public Property Get GrowSize() As Long
  GrowSize = mlGrowSize
End Property

Public Property Let GrowSize(ByVal plGrowSize As Long)
  mlGrowSize = plGrowSize
End Property

Public Property Get Count() As Long
  Count = mlIdxCount
End Property

'
' Sorting, public methods
'

'Sort columns string format: [!]Column name[+|-][, [!]ColumnName2[+|-] [, ...] ]
Private Sub ParseSortColumns(ByRef psSortColumns As String)
  Dim asColName()     As String
  Dim sColName        As String
  Dim i               As Integer
  Dim sCaseSensitive  As String
  Dim fCaseSensitive  As Boolean
  Dim sAscDesc        As String
  
  'If the sort columns definition is empty, we reinit and exit
  If Len(psSortColumns) = 0 Then
    If miSortColCt Then Erase matSortInfo()
    miSortColCt = 0
    mfSorted = False
    Exit Sub
  End If
  'each column on which we'll have to sort is separated by a comma
  miSortColCt = SplitString(asColName(), psSortColumns, ",")
  ReDim matSortInfo(1 To miSortColCt)
  For i = 1 To miSortColCt
    sColName = asColName(i)
    sCaseSensitive = Left$(sColName, 1)
    sAscDesc = Right$(sColName, 1)
    With matSortInfo(i)
      'The column name must be banged (preceeded by "!") for the sort to be case sensitive
      If sCaseSensitive = "!" Then
        .fCaseSensitive = True
        sColName = Right$(sColName, Len(sColName) - 1)
      Else
        .fCaseSensitive = False
      End If
      'Last char of column name may indicate asc or desc sorting
      If (sAscDesc = "+") Or (sAscDesc = "-") Then
        .fDesc = CBool(sAscDesc = "-")
        sColName = Left$(sColName, Len(sColName) - 1)
      Else
        .fDesc = False
      End If
      .sColName = sColName
      On Error Resume Next
      .lColIndex = ColPos(sColName)
      If .lColIndex = 0& Then
        'this can't be a good thing: we can't sort on an unknown column
        '(we borrow VB error 5 (invalid procedure call); we could also
        'have used error 9 (subscript out of range))
        Err.Raise 5&, "CList::ParseSortColumns", sColName & ": is not an existing column; can't sort list"
        Exit Sub
      End If
    End With
  Next i
End Sub

'(Complementary function for handling sort info)
'Need to now if a column is part of the sort key
Private Function IsSortColumn(ByVal plColIndex As Long) As Boolean
  Dim i               As Integer
  Dim sCompColName    As String
  Dim eCompMethod     As VbCompareMethod
  
  If miSortColCt Then
    If moColIndexMap.CaseSensitive Then
      eCompMethod = vbBinaryCompare
    Else
      eCompMethod = vbTextCompare
    End If
    sCompColName = ColName(plColIndex)
    For i = 1 To miSortColCt
      If StrComp(matSortInfo(i).sColName, sCompColName, eCompMethod) = 0 Then
        IsSortColumn = True
        Exit Function
      End If
    Next i
  End If
End Function

Public Property Get Sorted() As Boolean
  Sorted = mfSorted
End Property

Public Sub SetUnsorted()
  If miSortColCt Then Erase matSortInfo()
  mfSorted = False: miSortColCt = 0: ClearAllSysColFlags
End Sub

Public Sub SetSorted(ByVal psSortColumns As String)
  Dim i         As Integer
  Dim iColPos   As Long
  
  ParseSortColumns psSortColumns
  msSortColumns = psSortColumns
  If miSortColCt Then
    For i = 1 To miSortColCt
      iColPos = matSortInfo(i).lColIndex
      If matSortInfo(i).fCaseSensitive Then
        SetSysColFlags iColPos, klColFlagSortedCaseSensitive
      Else
        SetSysColFlags iColPos, klColFlagSortedCaseInsensitive
      End If
    Next i
    mfSorted = True
  End If
End Sub

Public Sub Sort(ByVal psSortColumns As String, Optional ByVal plStartRow As Long = 1&, Optional ByVal plEndRow As Long = 0&)
  Dim i         As Integer
  Dim iColPos   As Long
  
  If plEndRow = 0& Then plEndRow = mlIdxCount
  ClearAllSysColFlags
  ParseSortColumns psSortColumns 'Will raise an error if a column name doesn't exist or map
  msSortColumns = psSortColumns
  
  If miSortColCt Then
    mfSorted = False
    'resolve sort columns indices
    QuickSortRows plStartRow, plEndRow
    If (plStartRow = 1&) And (plEndRow = mlIdxCount) Then
      'set columns sys flags
      For i = 1 To miSortColCt
        iColPos = matSortInfo(i).lColIndex
        If matSortInfo(i).fCaseSensitive Then
          SetSysColFlags iColPos, klColFlagSortedCaseSensitive
        Else
          SetSysColFlags iColPos, klColFlagSortedCaseInsensitive
        End If
      Next i
      mfSorted = True
      mlCurRow = 0&
    End If
  End If
End Sub

Public Property Get SortColumns()
  SortColumns = msSortColumns
End Property

Public Property Get Item(ByVal pvColIndex As Variant, ByVal plRowIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
  Dim lColIndex     As Long
  If VarType(pvColIndex) = vbString Then
    lColIndex = ColPos(pvColIndex)
  Else
    lColIndex = pvColIndex
  End If
  Item = DataArray(lColIndex - 1, malIndex(plRowIndex))
End Property

Public Property Let Item(ByVal pvColIndex As Variant, ByVal plRowIndex As Long, ByRef pvCellValue As Variant)
  Dim lColIndex     As Long
  If VarType(pvColIndex) = vbString Then
    lColIndex = ColPos(pvColIndex)
  Else
    lColIndex = pvColIndex
  End If
  DataArray(lColIndex - 1, malIndex(plRowIndex)) = pvCellValue
  If IsSortColumn(lColIndex) Then mfSorted = False
End Property

'
' Add row data
'

Public Function AddRow(prowToAdd As CRow, _
                       Optional ByVal plInsertAfter As Long = 0&, _
                       Optional ByVal plInsertBefore As Long = 0&) As Long
  Dim lDataIndex    As Long
  Dim lCol          As Long
  Dim lRowIndex     As Long
  
  'Grow index array if necessary
  If mlIdxCount = mlIdxArraySize Then
    ReDim Preserve malIndex(1 To mlIdxArraySize + mlGrowSize)
    mlIdxArraySize = mlIdxArraySize + mlGrowSize
  End If
  'Find a slot where to insert the new row data
  lDataIndex = AllocRow(prowToAdd)
  'Adding a row breaks sorted state
  SetUnsorted
  If plInsertAfter Then
    lRowIndex = plInsertAfter + 1&
    'Push down other index pointers
    If lRowIndex < mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - plInsertAfter) * LenB(lRowIndex)
    End If
  ElseIf plInsertBefore Then
    lRowIndex = plInsertBefore
    If mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - lRowIndex + 1&) * LenB(lRowIndex)
    End If
  Else
    'Simply insert the new element at the end of the index array
    lRowIndex = mlIdxCount + 1&
  End If
  mlIdxCount = mlIdxCount + 1&
  malIndex(lRowIndex) = lDataIndex
  'copy row data
  For lCol = 1& To mlColCount
    If Len(matColDef(lCol).sColName) Then
      If prowToAdd.ColExists(matColDef(lCol).sColName) Then
        DataArray(lCol - 1&, lDataIndex) = prowToAdd(matColDef(lCol).sColName)
      Else
        DataArray(lCol - 1&, lDataIndex) = Null
      End If
    Else
      DataArray(lCol - 1&, lDataIndex) = Null
    End If
  Next lCol
  AddRow = lRowIndex
End Function

'Return the index in DataArray() for the new row copy it
Private Function AllocRow(ByRef prowNew As CRow) As Long
  Dim lRetIndex   As Long
  
  lRetIndex = GarbQPop()
  If lRetIndex = -1& Then
    'No free slot in the garbage queue, add a new row in the DataArray
    If mlDataArraySlotCount = mlDataArraySize Then
      If mlColCount Then
        ReDim Preserve DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      Else
        ReDim DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      End If
      mlDataArraySize = mlDataArraySize + mlGrowSize
    End If
    lRetIndex = mlDataArraySlotCount
    mlDataArraySlotCount = mlDataArraySlotCount + 1&
  End If
  
  AllocRow = lRetIndex
End Function

'Return the index in DataArray() for a new row of data
Private Function AllocRowSlot() As Long
  Dim lRetIndex   As Long
  Dim lCol        As Long
  
  lRetIndex = GarbQPop()
  If lRetIndex = -1& Then
    'No free slot in the garbage queue, add a new row in the DataArray
    If mlDataArraySlotCount = mlDataArraySize Then
      If mlColCount Then
        ReDim Preserve DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      Else
        ReDim DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      End If
      mlDataArraySize = mlDataArraySize + mlGrowSize
    End If
    lRetIndex = mlDataArraySlotCount
    mlDataArraySlotCount = mlDataArraySlotCount + 1&
  End If
  
  AllocRowSlot = lRetIndex
End Function

Public Function AddValues(ParamArray pavValues() As Variant) As Long
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  Dim lRowIndex     As Long
  
  'Grow index array if necessary
  If mlIdxCount = mlIdxArraySize Then
    ReDim Preserve malIndex(1 To mlIdxArraySize + mlGrowSize)
    mlIdxArraySize = mlIdxArraySize + mlGrowSize
  End If
  'Find a slot where to insert the new row data
  lDataIndex = AllocRowSlot()
  'Adding a row breaks sorted state
  SetUnsorted
  'Simply insert the new element at the end of the index array
  lRowIndex = mlIdxCount + 1&
  mlIdxCount = mlIdxCount + 1&
  malIndex(lRowIndex) = lDataIndex
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    lListCol = lListCol + 1&
  Next lIndex
  AddValues = lRowIndex
End Function

Public Function AddValuesArray(pavValues As Variant, _
                               Optional ByVal plInsertAfter As Long = 0&, _
                               Optional ByVal plInsertBefore As Long = 0&) As Long
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  Dim lRowIndex     As Long
  
  'Grow index array if necessary
  If mlIdxCount = mlIdxArraySize Then
    ReDim Preserve malIndex(1 To mlIdxArraySize + mlGrowSize)
    mlIdxArraySize = mlIdxArraySize + mlGrowSize
  End If
  'Find a slot where to insert the new row data
  lDataIndex = AllocRowSlot()
  'Adding a row breaks sorted state
  SetUnsorted
  If plInsertAfter Then
    lRowIndex = plInsertAfter + 1&
    'Push down other index pointers
    If lRowIndex < mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - plInsertAfter) * LenB(lRowIndex)
    End If
  ElseIf plInsertBefore Then
    lRowIndex = plInsertBefore
    If mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - lRowIndex + 1&) * LenB(lRowIndex)
    End If
  Else
    'Simply insert the new element at the end of the index array
    lRowIndex = mlIdxCount + 1&
  End If
  mlIdxCount = mlIdxCount + 1&
  malIndex(lRowIndex) = lDataIndex
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    lListCol = lListCol + 1&
  Next lIndex
  AddValuesArray = lRowIndex
End Function

'
' Update row data
'

Public Sub AssignRow(ByVal plRowIndex As Long, ByRef prowSource As CRow)
  Dim lDataIndex    As Long
  Dim lCol          As Long
  
  lDataIndex = malIndex(plRowIndex)
  'copy row data
  For lCol = 1& To mlColCount
    If Len(matColDef(lCol).sColName) Then
      If prowSource.ColExists(matColDef(lCol).sColName) Then
        DataArray(lCol - 1&, lDataIndex) = prowSource(matColDef(lCol).sColName)
        'If the sorted column value is changed, then the sorted state is broken
        If IsSortColumn(lCol) And mfSorted Then SetUnsorted
      End If
    End If
  Next lCol
End Sub

Public Sub AssignValues(ByVal plRowIndex As Long, ParamArray pavValues() As Variant)
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  
  lDataIndex = malIndex(plRowIndex)
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    'If the sorted column value is changed, then the sorted state is broken
    If IsSortColumn(lListCol) And mfSorted Then SetUnsorted
    lListCol = lListCol + 1&
  Next lIndex
End Sub

Public Sub AssignValuesArray(ByVal plRowIndex As Long, ByRef pavValues() As Variant)
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  
  lDataIndex = malIndex(plRowIndex)
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    'If the sorted column value is changed, then the sorted state is broken
    If IsSortColumn(lListCol) And mfSorted Then SetUnsorted
    lListCol = lListCol + 1&
  Next lIndex
End Sub

'
' Remove rows
'

Public Sub Remove(ByVal plIndex As Long)
  Dim lSlot       As Long
  
  'decrement current row if needed
  If plIndex <= mlCurRow Then
    mlCurRow = mlCurRow - 1&
    If mlCurRow < 0& Then mlCurRow = 0&
  End If
  
  'Remove from the index array
  lSlot = malIndex(plIndex)
  
  If plIndex < mlIdxCount Then
    CopyMemory malIndex(plIndex), malIndex(plIndex + 1&), (mlIdxCount - 1&) * 4&
  End If
  mlIdxCount = mlIdxCount - 1&
  
  GarbQPush lSlot
End Sub

Private Sub FreeString(ByVal plSlot As Long)
  GarbQPush plSlot
End Sub

'
' Find data
'

Private Sub ClearFindInfo()
  With mtFindInfo
    .fUseSortKeys = False
    If .lSearchFieldsCt Then
      Erase .asSearchField
      Erase .alColIndex
      Erase .avSearchValue
      Erase .alFlags
    End If
    .lSearchFieldsCt = 0&
    .eSearchAlgo = eSearchAlgoSequential 'this one can always be applied
  End With
End Sub

'Warning, parameters may be 0 based arrays !
Private Sub ParseSearchColumns(ByRef pvSearchColumns As Variant, ByRef pvSearchCriteria As Variant)
  Dim iDims           As Integer
  Dim avBounds        As Variant
  Dim i               As Integer
  Dim eCompMethod     As VbCompareMethod
  Dim lParamIndex     As Long
  
  If moColIndexMap.CaseSensitive Then
    eCompMethod = vbBinaryCompare
  Else
    eCompMethod = vbTextCompare
  End If
  With mtFindInfo
    'If pvSearchColumns is an array, then we search on multiple columns
    GetVarArrayBounds pvSearchColumns, iDims, avBounds
    If iDims Then
      lParamIndex = avBounds(0)
      .lSearchFieldsCt = avBounds(1) - avBounds(0) + 1&
      ReDim .asSearchField(1 To .lSearchFieldsCt)
      ReDim .alColIndex(1 To .lSearchFieldsCt)
      ReDim .alFlags(1 To .lSearchFieldsCt)
      .asSearchField(1&) = pvSearchColumns(lParamIndex)
    Else
      ReDim .asSearchField(1 To 1)
      ReDim .alColIndex(1 To 1)
      ReDim .alFlags(1 To 1)
      .lSearchFieldsCt = 1&
      .asSearchField(1&) = pvSearchColumns
      lParamIndex = 1&
    End If
    'transform field names and compute derived infos
    .fUseSortKeys = True
    For i = 1& To .lSearchFieldsCt
      If i > 1& Then .asSearchField(i) = pvSearchColumns(lParamIndex)
      If Left$(.asSearchField(i), 1) = "!" Then
        .alFlags(i) = .alFlags(i) Or klfifCaseSensitive
        .asSearchField(i) = Right$(.asSearchField(i), Len(.asSearchField(i)) - 1&)
      Else
        .alFlags(i) = .alFlags(i) And (Not klfifCaseSensitive)
      End If
      'see if we're in the sorted key fields sequence
      If .fUseSortKeys And (i <= miSortColCt) Then
        .fUseSortKeys = CBool(StrComp(.asSearchField(i), matSortInfo(i).sColName, eCompMethod) = 0)
        If .fUseSortKeys Then
          If matSortInfo(i).fDesc Then
            .alFlags(i) = .alFlags(i) Or klfifSortedDesc
          Else
            .alFlags(i) = .alFlags(i) And (Not klfifSortedDesc)
          End If
        Else
          .alFlags(i) = .alFlags(i) And (Not klfifSortedDesc)
        End If
      Else
        .fUseSortKeys = False
      End If
      On Error Resume Next
      'We just test for a valid column name by computing its index
      .alColIndex(i) = ColPos(.asSearchField(i))
      If .alColIndex(i) = 0& Then
        'We can't accept a non matching column; we reuse VB error 5 (invalid proc call)
        Err.Raise 5&, "CList::ParseSearchColumns", .asSearchField(i) & " is not a valid column name"
      End If
      On Error GoTo 0
      lParamIndex = lParamIndex + 1&
    Next i
    'We must have a corresponding number of values
    GetVarArrayBounds pvSearchCriteria, iDims, avBounds
    If iDims Then
      lParamIndex = avBounds(0)
      If (avBounds(1) - avBounds(0) + 1&) < .lSearchFieldsCt Then
        'there is a problem, there are less search values than
        'there are search fields. We just decrement the number
        'of search fields we take into account, to synchronize.
        .lSearchFieldsCt = iDims
      End If
      ReDim .avSearchValue(1 To .lSearchFieldsCt)
      .avSearchValue(1&) = pvSearchCriteria(lParamIndex)
    Else
      ReDim .avSearchValue(1 To 1)
      .avSearchValue(1&) = pvSearchCriteria
      lParamIndex = 1&
    End If
    'interpret field values
    For i = 1& To .lSearchFieldsCt
      If i > 1& Then .avSearchValue(i) = pvSearchCriteria(lParamIndex)
      'Compute flags
      If VarType(.avSearchValue(i)) = vbString Then
        If CBool(InStr(1, .avSearchValue(i), "?")) Then
          .alFlags(i) = .alFlags(i) Or klfifHasJokers
          'jokers/wildcards supported only on last search column
          If i < .lSearchFieldsCt Then .fUseSortKeys = False
        End If
        If CBool(InStr(1, .avSearchValue(i), "*")) Then
          .alFlags(i) = .alFlags(i) Or klfifHasWildcards
          'jokers/wildcards supported only on last search column
          If i < .lSearchFieldsCt Then .fUseSortKeys = False
        End If
        If (.alFlags(i) And klfifHasJokers) = 0& Then
          If (.alFlags(i) And klfifHasWildcards) = 0& Then
            .alFlags(i) = .alFlags(i) Or klfifExactSearch
          Else
            If Left$(.avSearchValue(i), 1) <> "*" Then
              .alFlags(i) = .alFlags(i) Or klfifRootSearch
              .avSearchValue(i) = Left$(.avSearchValue(i), Len(.avSearchValue(i)) - 1&)
            End If
          End If
        End If
      End If
      lParamIndex = lParamIndex + 1&
    Next i
  End With 'mtFindInfo
End Sub

'Find a specific value and return its index. Returns 0 if not found.
'Note: when there are duplicates, any one of the duplicate's
'index may be returned. To get the first row of the list, when
'there are duplicates, use the FindFirst method.
Public Function Find(ByVal pvSearchColumns As Variant, ByVal pvSearchCriteria As Variant, Optional ByVal plStartFrom As Long = 1&, Optional ByVal plEndTo As Long = 0&) As Long
  Dim lMidIndex       As Long
  Dim lMinIndex       As Long
  Dim lMaxIndex       As Long
  Dim iComp           As Integer
  
  'If no items in array then immediately exit
  If mlIdxCount = 0& Then Exit Function
  'Fix last index to search to
  If plEndTo = 0& Then plEndTo = mlIdxCount
  'build find info
  ClearFindInfo
  ParseSearchColumns pvSearchColumns, pvSearchCriteria
  
  'use a sequential search if we can't benefit from the sorted state
  If (Not mfSorted) Or (Not mtFindInfo.fUseSortKeys) Then
    For lMidIndex = plStartFrom To plEndTo
      iComp = CompareRowForFind(malIndex(lMidIndex))
      If iComp = 0 Then
        Find = lMidIndex  'Item found, exit
        Exit Function
      End If
    Next lMidIndex
    'not found sequentially, the function terminates here.
    Exit Function
  End If
  
  'dichotomic search is possible
  lMinIndex = plStartFrom
  lMaxIndex = plEndTo
  While True
    lMidIndex = (lMinIndex + lMaxIndex) \ 2&
    If lMaxIndex < lMinIndex Then Exit Function
    iComp = CompareRowForFind(malIndex(lMidIndex))
    If iComp = 1 Then
      lMinIndex = lMidIndex + 1&
    Else
      If iComp = -1 Then
        lMaxIndex = lMidIndex - 1&
      Else
        Find = lMidIndex
        Exit Function
      End If
    End If
  Wend
End Function

'Find the first string and return its index. Returns 0 if not found.
Public Function FindFirst(ByVal pvSearchColumns As Variant, ByVal pvSearchCriteria As Variant, Optional ByVal plStartFrom As Long = 1&, Optional ByVal plEndTo As Long = 0&) As Long
  Dim lMidIndex       As Long
  Dim lMinIndex       As Long
  Dim lMaxIndex       As Long
  Dim lSaveIndex      As Long
  Dim iComp           As Integer
  
  'If no items in array then immediately exit
  If mlIdxCount = 0& Then Exit Function
  'Fix last index to search to
  If plEndTo = 0& Then plEndTo = mlIdxCount
  'build find info
  ClearFindInfo
  ParseSearchColumns pvSearchColumns, pvSearchCriteria
  
  'use a sequential search if we can't benefit from the sorted state
  If (Not mfSorted) Or (Not mtFindInfo.fUseSortKeys) Then
    For lMidIndex = plStartFrom To plEndTo
      iComp = CompareRowForFind(malIndex(lMidIndex))
      If iComp = 0 Then
        FindFirst = lMidIndex  'Item found, exit
        Exit Function
      End If
    Next lMidIndex
    'not found sequentially, the function terminates here.
    Exit Function
  End If
  
  'dichotomic search is possible
  lMinIndex = plStartFrom
  lMaxIndex = plEndTo
  While True
    lMidIndex = (lMinIndex + lMaxIndex) \ 2&
    If lMaxIndex < lMinIndex Then Exit Function
    iComp = CompareRowForFind(malIndex(lMidIndex))
    If iComp = 1 Then
      lMinIndex = lMidIndex + 1&
    Else
      If iComp = -1 Then
        lMaxIndex = lMidIndex - 1&
      Else
        If iComp = 1 Then
          lMaxIndex = lMidIndex - 1&
        Else
          If iComp = 0 Then
            'We've found a corresponding string. Now we bubble up
            'sequentially, until we reach the first one of its
            'duplicates (if any).
            Do While lMidIndex > 1&
              lSaveIndex = lMidIndex
              lMidIndex = lMidIndex - 1&
              iComp = CompareRowForFind(malIndex(lMidIndex))
              If iComp Then
                lMidIndex = lSaveIndex
                Exit Do
              End If
            Loop
            FindFirst = lMidIndex
            Exit Function
          Else
            lMinIndex = lMidIndex + 1&
          End If
        End If
      End If
    End If
  Wend
End Function

'The array must be sorted.
Public Sub RemoveDuplicates()
  If Not mfSorted Then
    'Raise "Invalid procedure call" standard vb error
    Err.Raise 5&, "CList::RemoveDuplicates", VBA.Error$(5&) & ". List must be sorted."
    Exit Sub
  End If
  
  'To have a duplicate we must have at least two elements...
  If mlIdxCount < 2& Then Exit Sub
  
  'First pass: identify duplicates and put a -1 index in their slot pointer (sentinels)
  Dim lIndex      As Long
  Dim vCompValue  As Variant
  Dim vCellValue  As Variant
  Dim iComp       As Integer
  Dim lNewIndex   As Long
  Dim lRefIndex   As Long
  
  lRefIndex = 1&
  For lIndex = 2& To mlIdxCount
    iComp = CompareRowsForSort(malIndex(lRefIndex), malIndex(lIndex))
    If iComp Then
      lRefIndex = lIndex
    Else
      'duplicate: free string and zero its index
      GarbQPush lIndex
      malIndex(lIndex) = -1&
    End If
  Next lIndex
  
  'second pass: compress the array and free string duplicates
  lNewIndex = 0&
  For lIndex = 1& To mlIdxCount
    If malIndex(lIndex) > -1& Then
      lNewIndex = lNewIndex + 1&
      malIndex(lNewIndex) = malIndex(lIndex)
    End If
  Next lIndex
  mlIdxCount = lNewIndex
  mlCurRow = 0& 'the current row is lost
End Sub

Public Property Get Row(ByVal plRowIndex As Long) As CRow
  Dim oRetRow           As CRow
  Set oRetRow = New CRow
  GetRow oRetRow, plRowIndex
  Set Row = oRetRow
  Set oRetRow = Nothing
End Property

Public Sub GetRow(prowNewRet As CRow, ByVal plRowIndex As Long)
  Dim lDataColIndex     As Long
  Dim lDataRowIndex     As Long
  Dim lColIndex         As Long
  Dim fValidRow         As Boolean
  lDataRowIndex = malIndex(plRowIndex)
  fValidRow = ((plRowIndex > 0&) And (plRowIndex <= mlIdxCount))
  
  Set prowNewRet = New CRow
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
      If fValidRow Then
        avValues(lColIndex) = DataArray(lDataColIndex - 1&, lDataRowIndex)
      End If
    End With
  Next lColIndex
  prowNewRet.ArrayDefine avColName, avColType, avColSize, avColFlags
  
  If fValidRow Then
    prowNewRet.ArrayAssign avValues
  End If
End Sub

Public Sub DefineRow(prowRet As CRow)
  prowRet.Clear
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
      avColFlags(lColIndex) = .lFlags And &HFFFF&
      avValues(lColIndex) = Null
    End With
  Next lColIndex
  prowRet.ArrayDefine avColName, avColType, avColSize, avColFlags
End Sub

Public Sub DefineList(plstRet As CList)
  plstRet.Clear
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
      avColFlags(lColIndex) = .lFlags And &HFFFF&
    End With
  Next lColIndex
  plstRet.ArrayDefine avColName, avColType, avColSize, avColFlags
End Sub

Public Sub CopyFrom(ByRef plstSource As CList)
  Dim lColCount     As Long
  Dim iCol          As Long
  Dim lRow          As Long
  
  Clear
  lColCount = plstSource.ColCount
  If lColCount = 0& Then Exit Sub
  
  For iCol = 1& To lColCount
    With plstSource
      'vTemplateValue cannot be provided, so...
      Me.AddCol .ColName(iCol), _
                "", _
                .ColSize(iCol), _
                .ColFlags(iCol)
    End With
  Next iCol
  '...column types are synchronized afterwards
  For iCol = 1& To lColCount
    Me.ColType(iCol) = plstSource.ColType(iCol)
  Next iCol
  'There's no other solution than copying row by row
  If plstSource.Count Then
    For lRow = 1 To plstSource.Count
      Me.AddRow plstSource.Row(lRow)
    Next lRow
    If Len(plstSource.SortColumns) Then SetSorted plstSource.SortColumns
  End If
  mlCurRow = plstSource.CurrentRow
End Sub

Public Property Let Row(ByVal plRowIndex As Long, ByRef prowValues As CRow)
  If (plRowIndex < 1&) Or (plRowIndex > mlIdxCount) Then
    Exit Property
  End If
  Dim lCol          As Long
  Dim lDataIndex    As Long
  Dim lColPos       As Long
  'copy row data
  On Error Resume Next
  lDataIndex = malIndex(plRowIndex)
  For lCol = 1 To prowValues.ColCount
    lColPos = 0&
    If Len(prowValues.ColName(lCol)) Then lColPos = ColPos(prowValues.ColName(lCol))
    If lColPos Then DataArray(lColPos - 1&, lDataIndex) = prowValues(lCol)
  Next lCol
End Property

'
' Garbage queue
'

Private Function GarbQPop() As Long
  If mlGarbQCount Then
    GarbQPop = malGarbageQ(mlGarbQHead)
    If mlGarbQHead < mlGarbQSize Then
      mlGarbQHead = mlGarbQHead + 1&
    Else
      mlGarbQHead = 1&
    End If
    mlGarbQCount = mlGarbQCount - 1&
  Else
    GarbQPop = -1&
  End If
End Function

Private Sub GarbQPush(ByVal plIndexValue As Long)
  If mlGarbQCount = mlGarbQSize Then
    Dim lMoveCount    As Long
    Dim lMoveIndex    As Long
    Dim lOldSize      As Long
    lOldSize = mlGarbQSize
    lMoveCount = mlGarbQSize - mlGarbQHead + 1
    'Grow the queue array
    ReDim Preserve malGarbageQ(1 To mlGarbQSize + mlGrowSize)
    mlGarbQSize = mlGarbQSize + mlGrowSize
    If mlGarbQTail < mlGarbQHead Then
      'Move at end of queue array
      'This throws a GPF: CopyMemory malGarbageQ(mlGarbQHead + mlGrowSize), malGarbageQ(mlGarbQHead), mlGrowSize * 4&
      For lMoveIndex = mlGarbQSize To mlGarbQSize - lMoveCount Step -1&
        malGarbageQ(lMoveIndex) = malGarbageQ(lMoveIndex - lOldSize)
      Next lMoveIndex
      mlGarbQHead = mlGarbQHead + mlGrowSize
    End If
  End If
  
  If mlGarbQTail < mlGarbQSize Then
    mlGarbQTail = mlGarbQTail + 1&
  Else
    mlGarbQTail = 1&
  End If
  malGarbageQ(mlGarbQTail) = plIndexValue
  mlGarbQCount = mlGarbQCount + 1&
End Sub

'
' Sorting
'

Private Sub QuickSortRows(ByVal plLowBound As Long, ByVal plUpBound As Long)
  Dim lPivot  As Long
  Dim lTemp   As Long
  Dim lFirst  As Long
  Dim lLast   As Long
  Dim lMiddle As Long
  Dim iComp   As Integer
  
  'Locate Pivot
  lFirst = plLowBound
  lLast = plUpBound
  lMiddle = (lFirst + lLast) / 2&
  lPivot = malIndex(lMiddle)
  
  Do  'Move pointers against each other
    iComp = CompareRowsForSort(malIndex(lFirst), lPivot)
    While iComp = -1
      lFirst = lFirst + 1&
      iComp = CompareRowsForSort(malIndex(lFirst), lPivot)
    Wend
    iComp = CompareRowsForSort(malIndex(lLast), lPivot)
    While iComp = 1
      lLast = lLast - 1&
      iComp = CompareRowsForSort(malIndex(lLast), lPivot)
    Wend
    
    If lFirst <= lLast Then
      'Swap string (faked) pointers
      lTemp = malIndex(lFirst)
      malIndex(lFirst) = malIndex(lLast)
      malIndex(lLast) = lTemp
      lFirst = lFirst + 1&
      lLast = lLast - 1&
    End If
  Loop Until lFirst > lLast
  
  If plLowBound < lLast Then
    Call QuickSortRows(plLowBound, lLast)
  End If
  If lFirst < plUpBound Then
    Call QuickSortRows(lFirst, plUpBound)
  End If
End Sub

Private Function CompValues(pvValue1 As Variant, pvValue2 As Variant) As Integer
  If pvValue1 < pvValue2 Then
    CompValues = -1
  ElseIf pvValue1 > pvValue2 Then
    CompValues = 1
  Else
    CompValues = 0
  End If
End Function

Private Function CompStrings(ByVal psString1 As String, ByVal psString2 As String, peCompMethod As VbCompareMethod) As Integer
  CompStrings = StrComp(psString1, psString2, peCompMethod)
End Function

'Used to sort rows. Returns -1, 0 or 1, as StrComp does
Private Function CompareRowsForSort(ByVal plRow1 As Long, ByVal plRow2 As Long) As Integer
  'we compare in the matSortInfo fields order.
  Dim iFld            As Long
  Dim lColIndex       As Long
  Dim iType           As String
  Dim iComp           As Integer
  Dim iDir            As Integer
  Dim eStrCompMethod  As VbCompareMethod
  
  'sequential search for this column value
  For iFld = 1 To miSortColCt
    If matSortInfo(iFld).fDesc Then iDir = -1 Else iDir = 1
    lColIndex = malColIndex(matSortInfo(iFld).lColIndex)
    iType = matColDef(lColIndex).iDataType
    lColIndex = lColIndex - 1&
    If iType = VbVarType.vbString Then
      If matSortInfo(iFld).fCaseSensitive Then
        eStrCompMethod = vbBinaryCompare
      Else
        eStrCompMethod = vbTextCompare
      End If
      iComp = CompStrings(DataArray(lColIndex, plRow1), DataArray(lColIndex, plRow2), eStrCompMethod)
    Else
      iComp = CompValues(DataArray(lColIndex, plRow1), DataArray(lColIndex, plRow2))
    End If
    'if there is a difference between comparison operators then we can stop comparing
    If iComp Then Exit For
  Next iFld
  CompareRowsForSort = iDir * iComp
End Function

Private Function CompareRowForFind(ByVal plRowIndex As Long) As Integer
  Dim iFld            As Long
  Dim lColIndex       As Long
  Dim iComp           As Integer
  Dim iDir            As Integer
  Dim eStrCompMethod  As VbCompareMethod
  Dim lLenSearch      As Long
  Dim vTemp           As Variant
  Dim iType           As Integer
  
  With mtFindInfo
    'We compare in the search fields order.
    'As soon as a field by field compare is <>0 then exit
    For iFld = 1& To .lSearchFieldsCt
      If .alFlags(iFld) And klfifSortedDesc Then iDir = -1 Else iDir = 1
      lColIndex = malColIndex(.alColIndex(iFld))
      iType = matColDef(lColIndex).iDataType
      lColIndex = lColIndex - 1&
      vTemp = DataArray(lColIndex, plRowIndex)
      If iType = VbVarType.vbString Then
        'compare method for this search field derived from klfifCaseSensitive flag
        If .alFlags(iFld) And klfifCaseSensitive Then
          eStrCompMethod = vbBinaryCompare
        Else
          eStrCompMethod = vbTextCompare
        End If
        If iFld = .lSearchFieldsCt Then 'only on last search field we look for joker/wildcards
          If .alFlags(iFld) And klfifExactSearch Then
            'raw string compare
            iComp = CompStrings(.avSearchValue(iFld), vTemp & "", eStrCompMethod)
          ElseIf .alFlags(iFld) And klfifRootSearch Then
            lLenSearch = Len(.avSearchValue(iFld)) - 1&
            iComp = StrComp(.avSearchValue(iFld), Left$(vTemp & "", lLenSearch), eStrCompMethod)
          Else
            iComp = 1
            If StrMatch(vTemp, .avSearchValue(iFld), eStrCompMethod) Then iComp = 0
          End If
        Else
          'raw string compare
          iComp = CompStrings(.avSearchValue(iFld), vTemp & "", eStrCompMethod)
        End If
      Else
        iComp = CompValues(.avSearchValue(iFld), vTemp)
      End If
      If iComp Then Exit For
    Next iFld
  End With 'mtFindInfo
  CompareRowForFind = iDir * iComp
End Function

'
' A routine for comparing strings using simple pattern matching operators,
' like "*" for root or suffix expansion and "?" as character joker.
'
'Returns True if psToMatch matches the psModel pattern.
Private Function StrMatch(ByVal psToMatch As String, ByVal psModel As String, ByVal piCompMethod As VbCompareMethod) As Boolean
  'If no wildcard or joker then simply return StrComp
  If (InStr(1, psModel, "*") = 0) And (InStr(1, psModel, "?") = 0) Then
    StrMatch = CBool(StrComp(psModel, psToMatch, piCompMethod) = 0)
    Exit Function
  End If
  'If no model or nothing to match, no match
  If Len(psModel) = 0 Then Exit Function
  If Len(psToMatch) = 0 Then Exit Function
  'If model = "*" then match is ok
  If psModel = "*" Then StrMatch = True: Exit Function
    
  'Wildcard at start of string or at end. If anywhere in the middle, then ignore.
  Dim fRootExpand     As Boolean  'True if wildcard at end
  Dim fSuffixExpand   As Boolean  'True if wildcard at start
  Dim sExpandedModel  As String
  Dim sShrinkedModel  As String
  Dim iLgToMatch      As Integer
  Dim iLgModel        As Integer
  Dim iJoker          As Integer
    
  iLgToMatch = Len(psToMatch)
  iLgModel = Len(psModel)
  fRootExpand = (Right$(psModel, 1) = "*")
  fSuffixExpand = (Left$(psModel, 1) = "*")
    
  'Different algorithms for both expansions or for only one.
  If fRootExpand And fSuffixExpand Then
    'Note: This would be the slowest case if we handled possible jokers in psModel.
    iLgModel = iLgModel - 2
    If iLgToMatch > iLgModel Then
      'Example: A call like StrMatch("something", "*th*", vbTextCompare) has been made.
      'In this implementation, jokers ("?") are not supported in this case.
      sShrinkedModel = Mid$(psModel, 2, iLgModel)  'ie: "th" in our example
      StrMatch = CBool(InStr(1, psToMatch, sShrinkedModel, piCompMethod))
    ElseIf iLgToMatch < iLgModel Then
      'Example: A call like StrMatch("text", "*textual*", vbTextCompare) has been made.
      StrMatch = False
    Else 'If iLgToMatch = iLgModel Then
      'Example: A call like StrMatch("something", "*something*", vbTextCompare) has been made.
      StrMatch = CBool(StrComp(Mid$(psModel, 2, iLgModel), psToMatch, piCompMethod) = 0)
    End If
  Else
    iLgModel = iLgModel - 1 'don't count the *.
    If fRootExpand Then
      'Example: A call like StrMatch("something", "some*", vbTextCompare) has been made.
      If iLgToMatch > iLgModel Then
        iJoker = InStr(1, sExpandedModel, "?", vbBinaryCompare)
        If iJoker Then
          'We expand the model with missing characters borrowed to psToMatch.
          sExpandedModel = Left$(psModel, iLgModel) & Right$(psToMatch, iLgToMatch - iLgModel)
          'Now replace all jokers ("?") with corresponding char from psToMatch
          While iJoker
            Mid$(sExpandedModel, iJoker, 1) = Mid$(psToMatch, iJoker, 1)
            iJoker = InStr(iJoker + 1, sExpandedModel, "?", vbBinaryCompare)
          Wend
          'And finally compare
          StrMatch = CBool(StrComp(sExpandedModel, psToMatch, piCompMethod) = 0)
        Else
          StrMatch = CBool(StrComp(Left$(psModel, iLgModel), Left$(psToMatch, iLgModel), piCompMethod) = 0)
        End If
      ElseIf iLgToMatch < iLgModel Then
        StrMatch = False
      Else 'If iLgToMatch = iLgModel Then
        StrMatch = CBool(StrComp(Left$(psModel, iLgModel), psToMatch, piCompMethod) = 0)
      End If
    Else 'can only be suffix expand now.
      'Example: A call like StrMatch("something", "*thing", vbTextCompare) has been made.
      If iLgToMatch > iLgModel Then
        'We expand the model with missing characters borrowed to psToMatch.
        sExpandedModel = Left$(psToMatch, iLgToMatch - iLgModel) & Right$(psModel, iLgModel)
        'Now replace all jokers ("?") with corresponding char from psToMatch
        iJoker = InStr(1, sExpandedModel, "?")
        While iJoker
          Mid$(sExpandedModel, iJoker, 1) = Mid$(psToMatch, iJoker, 1)
          iJoker = InStr(iJoker + 1, sExpandedModel, "?", vbBinaryCompare)
        Wend
        'And finally compare
        StrMatch = CBool(StrComp(sExpandedModel, psToMatch, piCompMethod) = 0)
      ElseIf iLgToMatch < iLgModel Then
        StrMatch = False
      Else 'If iLgToMatch = iLgModel Then
        StrMatch = CBool(StrComp(Left$(psModel, iLgModel), psToMatch, piCompMethod) = 0)
      End If
    End If
  End If
End Function

'
' Handle system column flags (not used outside the class)
'

Private Sub ClearAllSysColFlags()
  Dim i     As Long
  On Error Resume Next
  For i = 1& To mlColCount
    matColDef(malColIndex(i)).lFlags = matColDef(malColIndex(i)).lFlags And &HFFFF&
  Next i
End Sub

Private Sub SetAllSysColFlags(ByVal plSysFlags As Long)
  Dim i     As Long
  On Error Resume Next
  plSysFlags = plSysFlags And &HFFFF0000
  For i = 1& To mlColCount
    matColDef(malColIndex(i)).lFlags = matColDef(malColIndex(i)).lFlags Or plSysFlags
  Next i
End Sub

Private Sub SetSysColFlags(ByVal plColIndex As Long, ByVal plSysFlags As Long)
  plSysFlags = plSysFlags And &HFFFF0000
  matColDef(malColIndex(plColIndex)).lFlags = matColDef(malColIndex(plColIndex)).lFlags Or plSysFlags
End Sub

Private Sub ClearSysColFlags(ByVal plColIndex As Long)
  matColDef(malColIndex(plColIndex)).lFlags = matColDef(malColIndex(plColIndex)).lFlags And &HFFFF&
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
  
  'mlGrowSize
  lsize = lsize + klSizeOfLong
  'Behaviour
  lsize = lsize + klSizeOfBool  'mfSorted
  lsize = lsize + klSizeOfLong + Len(msSortColumns) 'msSortColumns
  
  'DataArray
  'mlDataArraySize, mlDataArraySlotCount
  lsize = lsize + 2& * klSizeOfLong
  lsize = lsize + CalcByteSize(DataArray)
  
  'Row handles
  'mlIdxArraySize, mlIdxCount,mlCurRow
  lsize = lsize + 2& * klSizeOfLong
  lsize = lsize + mlIdxCount * klSizeOfLong
  lsize = lsize + klSizeOfLong
  
  'Column definitions
  lsize = lsize + klSizeOfLong              'mlColCount
  lsize = lsize + mlColCount * klSizeOfLong 'all the malColIndex() array elements
  'All the TColDef items of matColDef, member by member
  lsize = lsize + mlColCount * klSizeOfInt  'iDataType
  lsize = lsize + mlColCount * klSizeOfLong 'lDataSize
  lsize = lsize + mlColCount * klSizeOfLong 'lFlags
  For i = 1 To mlColCount
    lsize = lsize + klSizeOfLong + Len(matColDef(i).sColName)
  Next i
  lsize = lsize + moColIndexMap.IIObjectBytes.ByteSize()
  
  'Finally, the garbage queue
  lsize = lsize + 4& * klSizeOfLong 'mlGarbQCount, mlGarbQSize, mlGarbQHead, mlGarbQTail
  If mlGarbQCount Then
    lsize = lsize + mlGarbQCount * klSizeOfLong
  End If
  
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
  PokeLong klCIDList, abRetObject, lByte
  PokeLong MAKE_VERSIONLONG(ksClassVersion), abRetObject, lByte
  
  'Write behaviour members
  PokeLong mlGrowSize, abRetObject, lByte
  PokeBoolean mfSorted, abRetObject, lByte
  PokeString msSortColumns, abRetObject, lByte
    
  'DataArray
  PokeLong mlDataArraySize, abRetObject, lByte
  PokeLong mlDataArraySlotCount, abRetObject, lByte
  MoveVariant abRetObject(), lByte, DataArray
  
  'Row handles and current row
  PokeLong mlIdxArraySize, abRetObject, lByte
  PokeLong mlIdxCount, abRetObject, lByte
  For i = 1 To mlIdxCount
    PokeLong malIndex(i), abRetObject, lByte
  Next i
  PokeLong mlCurRow, abRetObject, lByte
  
  'Column definitions
  PokeLong mlColCount, abRetObject, lByte
  For i = 1 To mlColCount
    PokeLong malColIndex(i), abRetObject, lByte
  Next i
  
  'All the TColDef items of matColDef, member by member
  For i = 1 To mlColCount
    With matColDef(i)
      PokeInteger .iDataType, abRetObject, lByte
      PokeLong .lDataSize, abRetObject, lByte
      PokeLong .lFlags, abRetObject, lByte
      PokeString .sColName, abRetObject, lByte
    End With
  Next i
  
  Dim abMapBytes() As Byte
  moColIndexMap.IIObjectBytes.GetDataBytes abMapBytes()
  CopyMemory abRetObject(lByte), abMapBytes(1), UBound(abMapBytes)
  lByte = lByte + UBound(abMapBytes)
  Erase abMapBytes
  
  'Finally, the garbage queue
  PokeLong mlGarbQCount, abRetObject, lByte
  PokeLong mlGarbQSize, abRetObject, lByte
  PokeLong mlGarbQHead, abRetObject, lByte
  PokeLong mlGarbQTail, abRetObject, lByte
  If mlGarbQCount Then
    iIndex = mlGarbQHead
    For i = 1 To mlGarbQCount
      PokeLong malGarbageQ(iIndex), abRetObject, lByte
      iIndex = iIndex + 1&
      If iIndex > mlGarbQSize Then iIndex = 1&
    Next i
  End If
End Sub

Private Sub IObjectBytes_SetDataBytes(abObjectData() As Byte, Optional ByRef plByte As Long = 1&)
  Dim i             As Long
  Dim iIndex        As Long
  Dim lClassID      As Long
  Dim lClassVer     As Long
  
  Clear 'load all defaults
  
  'ClassID and version
  lClassID = PeekLong(abObjectData(), plByte)
  lClassVer = PeekLong(abObjectData, plByte)
  If lClassID <> klCIDList Then
    Err.Raise MAKE_OBJECTERROR(kErrBadClassIDBytes), "CList", "CList data bytes class ID doesn't match"
  End If
  If lClassVer <> MAKE_VERSIONLONG(ksClassVersion) Then
    Err.Raise MAKE_OBJECTERROR(kErrBadClassVerBytes), "CList", "CList data bytes version doesn't match"
  End If
  
  'behaviour members
  mlGrowSize = PeekLong(abObjectData, plByte)
  mfSorted = PeekBoolean(abObjectData, plByte)
  msSortColumns = PeekString(abObjectData, plByte)
    
  'DataArray
  mlDataArraySize = PeekLong(abObjectData, plByte)
  mlDataArraySlotCount = PeekLong(abObjectData, plByte)
  GetVariant DataArray, plByte, abObjectData()
  
  'Row handles and current row
  mlIdxArraySize = PeekLong(abObjectData, plByte)
  mlIdxCount = PeekLong(abObjectData, plByte)
  ReDim malIndex(1 To mlIdxArraySize) As Long
  If mlIdxCount Then
    For i = 1 To mlIdxCount
      malIndex(i) = PeekLong(abObjectData, plByte)
    Next i
  End If
  mlCurRow = PeekLong(abObjectData, plByte)
  
  'Column definitions
  mlColCount = PeekLong(abObjectData, plByte)
  If mlColCount Then
    ReDim malColIndex(1 To mlColCount)
    ReDim matColDef(1 To mlColCount)
    For i = 1 To mlColCount
      malColIndex(i) = PeekLong(abObjectData, plByte)
    Next i
    'All the TColDef items of matColDef, member by member
    For i = 1 To mlColCount
      With matColDef(i)
        .iDataType = PeekInteger(abObjectData, plByte)
        .lDataSize = PeekLong(abObjectData, plByte)
        .lFlags = PeekLong(abObjectData, plByte)
        .sColName = PeekString(abObjectData, plByte)
      End With
    Next i
  End If
  
  moColIndexMap.IIObjectBytes.SetDataBytes abObjectData, plByte
  
  'Finally, read the garbage queue
  mlGarbQCount = PeekLong(abObjectData, plByte)
  mlGarbQSize = PeekLong(abObjectData, plByte)
  mlGarbQHead = PeekLong(abObjectData, plByte)
  mlGarbQTail = PeekLong(abObjectData, plByte)
  ReDim malGarbageQ(1 To mlGarbQSize)
  If mlGarbQCount Then
    iIndex = mlGarbQHead
    For i = 1 To mlGarbQCount
      malGarbageQ(iIndex) = PeekLong(abObjectData, plByte)
      iIndex = iIndex + 1&
      If iIndex > mlGarbQSize Then iIndex = 1&
    Next i
  End If
End Sub


