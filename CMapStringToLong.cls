VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMapStringToLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'CMapStringToLong.bas class module
'This class associates strings (keys) with long integer (32bits) values.
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
'When       � Version  � Who � What
'-----------+----------+-----+-----------------------------------------------------
'           �          �     �
Option Explicit

Implements IObjectBytes

'class id for implementing IObjectBytes interface
Private Const klCIDMapStringToLong     As Long = 1000&

'Behaviour
Private Const kiDefaultGrowSize As Integer = 20
Private mlGrowSize          As Long
Private mfSorted            As Boolean  'If true, the class keeps the string array sorted
Private mfCaseSensitive     As Boolean
Private miCompareMethod     As VbCompareMethod

'String allocator storage
Private masString()         As String
Private mlStrArraySize      As Long     'Size of string array
Private mlStrArraySlotCount As Long     'Number of used elements (some may have been freed)
'String allocator: garbage queue (circular array queue)
Private malGarbageQ()       As Long
Private mlGarbQSize         As Long
Private mlGarbQHead         As Long
Private mlGarbQTail         As Long
Private mlGarbQCount        As Long

'Map item
Private Type TMapItem
  lIndex      As Long
  lLongValue  As Long
End Type
'Map memory
Private mlMapItemSize       As Long
Private matMap()            As TMapItem
Private mlMapArraySize      As Long     'Size of index array (upper bound)
Private mlMapCount          As Long     'Number of elements in the string array

'For iObjectBytes interface
Private Const ksClassVersion As String = "01.00.00"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
              lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Sub Class_Initialize()
  Dim tCalcSize   As TMapItem
  mlMapItemSize = LenB(tCalcSize)
  mlGrowSize = kiDefaultGrowSize
  Clear
  mfSorted = False
  mfCaseSensitive = False
  miCompareMethod = vbBinaryCompare
End Sub

Private Sub Class_Terminate()
  '...
End Sub

Public Sub Clear()
  'String allocator: clear string array
  If mlMapCount Then Erase masString()
  mlStrArraySlotCount = 0&
  mlStrArraySize = 0&
  'String allocator: clear garbage queue
  mlGarbQSize = mlGrowSize
  ReDim malGarbageQ(1 To mlGarbQSize)
  mlGarbQHead = 1&
  mlGarbQTail = 0&
  mlGarbQCount = 0&
  'Map data: clear mapping array
  mlMapArraySize = mlGrowSize
  ReDim matMap(1 To mlMapArraySize)
  mlMapCount = 0&
End Sub

Public Property Get GrowSize() As Long
  GrowSize = mlGrowSize
End Property

Public Property Let GrowSize(ByVal plGrowSize As Long)
  mlGrowSize = plGrowSize
End Property

Public Property Get Count() As Long
  Count = mlMapCount
End Property

Public Property Get Key(ByVal plIndex As Long) As String
Attribute Key.VB_UserMemId = 0
  Key = masString(matMap(plIndex).lIndex)
End Property

Public Property Let Key(ByVal plIndex As Long, ByRef psNewValue As String)
  If mfSorted Then
    'Raise "Invalid procedure call" standard vb error
    Err.Clear: Err.Raise 5&, "CMapStringToLong::Key[Let]", VBA.Error$(5&) & ". Can't change a key when Sorted property is True."
  Else
    masString(matMap(plIndex).lIndex) = psNewValue
  End If
End Property

Public Property Get Item(ByVal plIndex As Long) As Long
  Item = matMap(plIndex).lLongValue
End Property

Public Property Let Item(ByVal plIndex As Long, ByVal plNewLong As Long)
  matMap(plIndex).lLongValue = plNewLong
End Property

Public Property Get Sorted() As Boolean
  Sorted = mfSorted
End Property

Public Property Let Sorted(ByVal pfSorted As Boolean)
  If pfSorted <> mfSorted Then
    If Not mfSorted Then
      If mlMapCount Then
        'Sort current array
        QuickSort
      End If
    End If
    mfSorted = pfSorted
  End If
End Property

Public Property Get CaseSensitive() As Boolean
  CaseSensitive = mfCaseSensitive
End Property

'Property should not be changed if array is sorted
Public Property Let CaseSensitive(ByVal pfCaseSensitive As Boolean)
  If Not mfSorted Then
    If mfCaseSensitive <> pfCaseSensitive Then
      mfCaseSensitive = pfCaseSensitive
      If mfCaseSensitive Then
        miCompareMethod = vbBinaryCompare
      Else
        miCompareMethod = vbTextCompare
      End If
    End If
  Else
    'Raise "Invalid procedure call" standard vb error
    Err.Clear: Err.Raise 5&, "CMapStringToLong::CaseSensitive[Let]", VBA.Error$(5&) & ". Array already sorted."
  End If
End Property

Public Sub Add(ByRef psKey As String, ByVal plItem As Long)
  Dim lStrIndex   As Long
  
  'Grow index array if necessary
  If mlMapCount = mlMapArraySize Then
    ReDim Preserve matMap(1 To mlMapArraySize + mlGrowSize)
    mlMapArraySize = mlMapArraySize + mlGrowSize
  End If
  
  lStrIndex = AllocString(psKey)
  
  If (Not mfSorted) Or (mlMapCount = 0) Then
    'Simply insert the new element at the end of the index array
    mlMapCount = mlMapCount + 1&
    With matMap(mlMapCount)
      .lIndex = lStrIndex
      .lLongValue = plItem
    End With
  Else
    'Use insertion sort with dichotomic search for place where to insert.
    Dim iInsertIndex As Long
    
    'Find the index where to insert the string
    iInsertIndex = FindStringPos(psKey)
    
    'Push down existing indices
    If iInsertIndex < mlMapArraySize Then
      CopyMemory matMap(iInsertIndex + 1&), matMap(iInsertIndex), mlMapItemSize * (mlMapCount - iInsertIndex + 1&)
    End If
    With matMap(iInsertIndex)
      .lIndex = lStrIndex
      .lLongValue = plItem
    End With
    mlMapCount = mlMapCount + 1&
  End If
End Sub

'Dichotomic search
Private Function FindStringPos(ByRef psNewString As String) As Long
  Dim lMin        As Long
  Dim lMax        As Long
  Dim lMiddle     As Long
  Dim iComp       As Integer
  Dim iComp2      As Integer
  
  lMin = 1: lMax = mlMapCount
  Do While lMin <= lMax
    ' Don't divide by zero.
    If StrComp(masString(matMap(lMin).lIndex), masString(matMap(lMax).lIndex), miCompareMethod) = 0 Then
      iComp = StrComp(masString(matMap(lMin).lIndex), psNewString, miCompareMethod)
      If (iComp >= 0) Then
        FindStringPos = lMin
      Else
        FindStringPos = lMax + 1&
      End If
      Exit Do
    End If
    
    ' Compute the dividing point.
    lMiddle = (lMax - lMin) \ 2& + 1&
    
    ' Make sure we stay in bounds.
    If lMiddle < lMin Then lMiddle = lMin
    If lMiddle > lMax Then lMiddle = lMax
    
    iComp = StrComp(masString(matMap(lMiddle).lIndex), psNewString, miCompareMethod)
    If iComp = 0 Then ' We found it.
      FindStringPos = lMiddle
      Exit Do
    ElseIf iComp = -1 Then ' Search the right half.
      lMin = lMiddle + 1&
    Else ' Search the left half.
      lMax = lMiddle - 1&
    End If
  Loop
  
  If FindStringPos < lMin Then
    ' At this point lMax <= lMin.
    If lMax < lMin Then
      lMax = lMin
    ElseIf lMin > lMax Then
      lMin = lMax
    End If
    iComp = StrComp(masString(matMap(lMax).lIndex), psNewString, miCompareMethod)
    iComp2 = StrComp(masString(matMap(lMin).lIndex), psNewString, miCompareMethod)
    If (iComp >= 0) Then
      FindStringPos = lMax
    ElseIf (iComp2 >= 0) Then
      FindStringPos = lMin
    Else
      FindStringPos = lMin + 1&
    End If
  End If
End Function

Public Sub Remove(ByVal plIndex As Long)
  Dim lSlot       As Long
  
  'Remove from the index array
  lSlot = matMap(plIndex).lIndex
  
  If plIndex < mlMapCount Then
    If plIndex < mlMapCount Then
      CopyMemory matMap(plIndex), matMap(plIndex + 1&), (mlMapCount - plIndex) * mlMapItemSize
    End If
  End If
  mlMapCount = mlMapCount - 1&
  
  FreeString lSlot
End Sub

'Remove all entries which are associated to the specified value
Public Sub RemoveMappingsFor(ByVal plLongValue As Long)
  Dim lFoundIndex   As Long
  Dim i             As Long
  Do
    lFoundIndex = 0&
    For i = 1 To mlMapCount
      If matMap(i).lLongValue = plLongValue Then
        lFoundIndex = i
        Exit For
      End If
    Next i
    If lFoundIndex Then
      Remove lFoundIndex
    End If
  Loop Until lFoundIndex = 0&
End Sub

'Find a specific string and return its index. Returns 0 if not found.
'Note: when there are duplicate string, any one of the duplicate's
'index may be returned. To get the first string in the array, when
'there are duplicates, use the FindFirst method.
Public Function Find(ByRef psSearch As String) As Long
  Dim lMidIndex   As Long
  Dim lMinIndex   As Long
  Dim lMaxIndex   As Long
  Dim lLenSearch  As Long
  Dim iComp       As Integer
  
  'If no items in array then immediately exit
  If mlMapCount = 0& Then Exit Function
  'Cannot find an item if not sorted
  If Not mfSorted Then
    'Raise "Invalid procedure call" standard vb error
    Err.Clear: Err.Raise 5&, "CMapStringToLong::Find()", VBA.Error$(5&) & ". Can't find key if not Sorted."
    Exit Function
  End If
  
  lMinIndex = 1&
  lMaxIndex = mlMapCount
  
  lLenSearch = Len(psSearch)
  While True
    lMidIndex = (lMinIndex + lMaxIndex) \ 2&
      If lMaxIndex < lMinIndex Then Exit Function
      iComp = StrComp(psSearch, masString(matMap(lMidIndex).lIndex), miCompareMethod)
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
Public Function FindFirst(ByRef psSearch As String, Optional ByVal pfRootSearch As Boolean = False) As Long
  Dim lMidIndex   As Long
  Dim lMinIndex   As Long
  Dim lMaxIndex   As Long
  Dim lLenSearch  As Long
  Dim sTemp       As String
  Dim iComp       As Integer
  Dim lSaveIndex  As Long
  
  'If no items in array then immediately exit
  If mlMapCount = 0& Then Exit Function
  'Cannot find an item if not sorted
  If Not mfSorted Then
    'Raise "Invalid procedure call" standard vb error
    Err.Clear: Err.Raise 5&, "CMapStringToLong::FindFirst()", VBA.Error$(5&) & ". Can't find key if not Sorted."
    Exit Function
  End If
  
  lMinIndex = 1&
  lMaxIndex = mlMapCount
  
  lLenSearch = Len(psSearch)
  While True
    lMidIndex = (lMinIndex + lMaxIndex) \ 2&
    If lMaxIndex < lMinIndex Then Exit Function
      
    sTemp = masString(matMap(lMidIndex).lIndex)
    iComp = StrComp(psSearch, sTemp, miCompareMethod)
    If iComp = 1 Then
      lMinIndex = lMidIndex + 1&
    Else
      If pfRootSearch Then
        iComp = StrComp(psSearch, Left$(sTemp, lLenSearch), miCompareMethod)
      End If
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
              If pfRootSearch Then
                iComp = StrComp(psSearch, Left$(masString(matMap(lMidIndex).lIndex), lLenSearch), miCompareMethod)
              Else
                iComp = StrComp(psSearch, masString(matMap(lMidIndex).lIndex), miCompareMethod)
              End If
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
    Err.Clear: Err.Raise 5&, "CMapStringToLong::RemoveDuplicates", VBA.Error$(5&) & ". Can't remove duplicates if not Sorted."
    Exit Sub
  End If
  
  'To have a duplicate we must have at least two elements...
  If mlMapCount < 2& Then Exit Sub
  
  'First pass: identify duplicates and put a zero index in their slot pointer (sentinels)
  Dim lIndex      As Long
  Dim sCompValue  As String
  Dim sCellValue  As String
  Dim iComp       As Integer
  Dim lNewIndex   As Long
  
  sCompValue = "Z" & masString(matMap(1).lIndex)   'Handle empty element
  For lIndex = 1& To mlMapCount
    sCellValue = masString(matMap(lIndex).lIndex)
    iComp = StrComp(sCellValue, sCompValue, miCompareMethod)
    If iComp Then
      sCompValue = sCellValue
    Else
      'duplicate: free string and zero its index
      GarbQPush lIndex
      matMap(lIndex).lIndex = 0&
    End If
  Next lIndex
  
  'second pass: compress the array and free string duplicates
  lNewIndex = 0&
  For lIndex = 1& To mlMapCount
    If matMap(lIndex).lIndex Then
      lNewIndex = lNewIndex + 1&
      matMap(lNewIndex) = matMap(lIndex)
    End If
  Next lIndex
  mlMapCount = lNewIndex
End Sub

'
' String allocator
'

'Return the index in masString() for the new string and copy it
Private Function AllocString(ByRef psNewString As String) As Long
  Dim lRetIndex   As Long
  
  lRetIndex = GarbQPop()
  If lRetIndex = 0 Then
    'No free slot in the garbage queue, add a new string in string array
    If mlStrArraySlotCount = mlStrArraySize Then
      ReDim Preserve masString(1 To mlStrArraySize + mlGrowSize)
      mlStrArraySize = mlStrArraySize + mlGrowSize
    End If
    mlStrArraySlotCount = mlStrArraySlotCount + 1&
    lRetIndex = mlStrArraySlotCount
  End If
  masString(lRetIndex) = psNewString
  
  AllocString = lRetIndex
End Function

Private Sub FreeString(ByVal plSlot As Long)
  GarbQPush plSlot
End Sub

'
' String allocator: Garbage queue
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

Private Sub QuickSort()
  If mlMapCount Then
    QuickSortProc 1&, mlMapCount
  End If
End Sub

Private Sub QuickSortProc(ByVal plLowBound As Long, ByVal plUpBound As Long)
  Dim sPivot  As String
  Dim tTemp   As TMapItem
  Dim lFirst  As Long
  Dim lLast   As Long
  Dim lMiddle As Long
  Dim iComp   As Integer
  
  'Locate Pivot
  lFirst = plLowBound
  lLast = plUpBound
  lMiddle = (lFirst + lLast) / 2&
  sPivot = masString(matMap(lMiddle).lIndex)
  
  Do  'Move pointers against each other
    iComp = StrComp(masString(matMap(lFirst).lIndex), sPivot, miCompareMethod)
    'Debug.Print "StrComp("; masString(matMap(lFirst)); ","; sPivot; ")="; iComp
    While iComp = -1
      lFirst = lFirst + 1&
      iComp = StrComp(masString(matMap(lFirst).lIndex), sPivot, miCompareMethod)
      'Debug.Print "StrComp("; masString(matMap(lFirst)); ","; sPivot; ")="; iComp
    Wend
    iComp = StrComp(masString(matMap(lLast).lIndex), sPivot, miCompareMethod)
    'Debug.Print "StrComp("; masString(matMap(lLast)); ","; sPivot; ")="; iComp
    While iComp = 1
      lLast = lLast - 1&
      iComp = StrComp(masString(matMap(lLast).lIndex), sPivot, miCompareMethod)
      'Debug.Print "StrComp("; masString(matMap(lLast)); ","; sPivot; ")="; iComp
    Wend
    
    'Debug.Print "lFirst="; lFirst; ", lLast="; lLast
    If lFirst <= lLast Then
      'Swap string (faked) pointers
      'Debug.Print "Swap("; matMap(lFirst); ","; matMap(lLast); ") <--> Swap("; masString(matMap(lFirst)); ","; masString(matMap(lLast)); ")"
      tTemp = matMap(lFirst)
      matMap(lFirst) = matMap(lLast)
      matMap(lLast) = tTemp
      lFirst = lFirst + 1&
      lLast = lLast - 1&
    End If
  Loop Until lFirst > lLast
  
  If plLowBound < lLast Then
    Call QuickSortProc(plLowBound, lLast)
  End If
  If lFirst < plUpBound Then
    Call QuickSortProc(lFirst, plUpBound)
  End If
End Sub

'
' dynaBasic5 specialisation
'
Private Property Get IObjectBytes_ByteSize() As Long
  IObjectBytes_ByteSize = GetInstanceByteSize()
End Property

Private Function GetInstanceByteSize() As Long
  Dim lsize       As Long
  Dim i           As Long
  
  'We store class id and class version (long + long)
  lsize = 2 * klSizeOfLong
  
  'mlGrowSize, mfSorted, mfCaseSensitive, miCompareMethod
  lsize = lsize + klSizeOfLong + klSizeOfBool + klSizeOfBool + klSizeOfInt 'behaviour
  'string array
  lsize = lsize + 2& * klSizeOfLong 'mlStrArraySize, mlStrArraySlotCount
  If mlStrArraySlotCount Then
    For i = 1& To mlStrArraySlotCount
      lsize = lsize + klSizeOfLong  'string length
      lsize = lsize + Len(masString(i))
    Next i
  End If
  'garbage queue
  lsize = lsize + 4& * klSizeOfLong 'mlGarbQCount, mlGarbQSize, mlGarbQHead, mlGarbQTail
  If mlGarbQCount Then
    lsize = lsize + mlGarbQCount * klSizeOfLong
  End If
  'map array
  lsize = lsize + 2& * klSizeOfLong  'mlMapCount, mlMapArraySize
  If mlMapCount Then
    lsize = lsize + mlMapCount * klSizeOfLong * 2& 'elements in matMap() are long+long
  End If
  GetInstanceByteSize = lsize
End Function

'Easily expose implemented IObjectBytes interface
Public Property Get IIObjectBytes() As IObjectBytes
  Set IIObjectBytes = Me
End Property

Private Sub IObjectBytes_GetDataBytes(abRetObject() As Byte)
  'It's our responsability to allocate space (ie redim abRetObject)
  Dim lByteSize     As Long
  Dim lByte         As Long
  Dim i             As Long
  Dim iIndex        As Long
  
  lByteSize = GetInstanceByteSize()
  
  On Error Resume Next
  Erase abRetObject
  On Error GoTo 0
  If lByteSize = 0& Then Exit Sub
  
  'transform instance data to byte array
  ReDim abRetObject(1 To lByteSize)
  lByte = 1&
  
  'Write class id and class version
  PokeLong klCIDMapStringToLong, abRetObject, lByte
  PokeLong MAKE_VERSIONLONG(ksClassVersion), abRetObject, lByte
  
  'Write behaviour members
  PokeLong mlGrowSize, abRetObject, lByte
  PokeBoolean mfSorted, abRetObject, lByte
  PokeBoolean mfCaseSensitive, abRetObject, lByte
  PokeInteger miCompareMethod, abRetObject, lByte
  'Write string table
  PokeLong mlStrArraySize, abRetObject, lByte
  PokeLong mlStrArraySlotCount, abRetObject, lByte
  If mlStrArraySlotCount Then
    For i = 1 To mlStrArraySlotCount
      PokeString masString(i), abRetObject, lByte
    Next i
  End If
  'Write garbage queue
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
  'Write map table
  PokeLong mlMapCount, abRetObject(), lByte
  PokeLong mlMapArraySize, abRetObject(), lByte
  If mlMapCount Then
    For i = 1 To mlMapCount
      PokeLong matMap(i).lIndex, abRetObject(), lByte
      PokeLong matMap(i).lLongValue, abRetObject(), lByte
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
  lClassID = PeekLong(abObjectData, plByte)
  lClassVer = PeekLong(abObjectData, plByte)
  If lClassID <> klCIDMapStringToLong Then
    Err.Raise MAKE_OBJECTERROR(kErrBadClassIDBytes), "CMapStringToLong", "CMapStringToLong data bytes class ID doesn't match"
  End If
  If lClassVer <> MAKE_VERSIONLONG(ksClassVersion) Then
    Err.Raise MAKE_OBJECTERROR(kErrBadClassVerBytes), "CMapStringToLong", "CMapStringToLong data bytes version doesn't match"
  End If
  
  'behaviour members
  mlGrowSize = PeekLong(abObjectData, plByte)
  mfSorted = PeekBoolean(abObjectData, plByte)
  mfCaseSensitive = PeekBoolean(abObjectData, plByte)
  miCompareMethod = PeekInteger(abObjectData, plByte)
  'read string table
  mlStrArraySize = PeekLong(abObjectData, plByte)
  mlStrArraySlotCount = PeekLong(abObjectData, plByte)
  ReDim masString(1 To mlStrArraySize)
  If mlStrArraySlotCount Then
    For i = 1 To mlStrArraySlotCount
      masString(i) = PeekString(abObjectData, plByte)
    Next i
  End If
  'Read garbage queue
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
  'Read map table
  mlMapCount = PeekLong(abObjectData(), plByte)
  mlMapArraySize = PeekLong(abObjectData(), plByte)
  ReDim matMap(1 To mlMapArraySize)
  If mlMapCount Then
    For i = 1 To mlMapCount
      matMap(i).lIndex = PeekLong(abObjectData(), plByte)
      matMap(i).lLongValue = PeekLong(abObjectData(), plByte)
    Next i
  End If
End Sub



