Attribute VB_Name = "MMain"
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'MMain.bas module
'Where all starts up.
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

'Needed for chronometer functions
#If Win64 Then
  'Win32 64bits API
  Type UINT64
      LowPart As Long
      HighPart As Long
  End Type
  Private Const BSHIFT_32 = 4294967296# ' 2 ^ 32
  Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As UINT64) As Long
  Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As UINT64) As Long
  'Needed for chronometer functions
  Private mcurFrequency   As UINT64
  Private mcurChronoStart As UINT64
#Else
  'Win32 API
  Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
  Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
  'Needed for chronometer functions
  Private mcurFrequency   As Currency
  Private mcurChronoStart As Currency
  
#End If


'Output file handle
Private mhOutput        As Integer

Sub Main()
  Dim fIsOpen     As Boolean
  Dim sOutputFile As String
  Dim sMsg        As String
  Dim sChoice     As String
  
  On Error GoTo Main_Err
  
  sOutputFile = CurrentProject.Path & "\output.txt"
  mhOutput = FreeFile
  Open sOutputFile For Output As #mhOutput
  
  sMsg = "Select test to run:" & vbCrLf & vbCrLf
  sMsg = sMsg & "----- CMapStringToLong Class -----" & vbCrLf & vbCrLf
  sMsg = sMsg & "1. Collection vs CMapStringToLong" & vbCrLf
  sMsg = sMsg & "2. Prove that collection keys are case insensitive" & vbCrLf
  sMsg = sMsg & "3. Duplicates in a CMapStringToLong object" & vbCrLf
  sMsg = sMsg & vbCrLf & "----- CRow Class -----" & vbCrLf & vbCrLf
  sMsg = sMsg & "4. Perfom all CRow test" & vbCrLf
  sMsg = sMsg & vbCrLf & "----- CList Class -----" & vbCrLf & vbCrLf
  sMsg = sMsg & "5. Add/update/remove rows test" & vbCrLf
  sMsg = sMsg & "6. Test list sort methods" & vbCrLf
  sMsg = sMsg & "7. Test find methods" & vbCrLf
  sMsg = sMsg & "10. Perform all list tests" & vbCrLf
  sMsg = sMsg & vbCrLf & "----- All tests -----" & vbCrLf & vbCrLf
  sMsg = sMsg & "99. Perform all tests" & vbCrLf
  'sMsg = sMsg & "" & vbCrLf
  'sMsg = sMsg & "" & vbCrLf
  'sMsg = sMsg & "" & vbCrLf
  'sMsg = sMsg & "" & vbCrLf
  sMsg = sMsg & vbCrLf & "or enter 0 (zero) to exit."
  
  sChoice = "1"
  Do
    sChoice = LongChooseBox(sMsg, "DISRowList library test driver", sChoice, 99, 0)
    If Len(sChoice) And (sChoice <> "0") Then
      ShowProgressDialog
      Select Case Val(sChoice)
      Case 1
        Test1
      Case 2
        Test2
      Case 3
        Test3
      Case 4
        Row1
        Row2
        Row3
        Row4
        Row5
        Row6
        Row7
      Case 5
        List1
      Case 6
        List2
      Case 7
        List3
      Case 10
        List1
        List2
        List3
      Case 99
        Test1
        Test2
        Test3
        Row1
        Row2
        Row3
        Row4
        Row5
        Row6
        Row7
        List1
        List2
        List3
      End Select
      MsgBox "Done.", vbOKOnly + vbInformation
      CloseProgressDialog
    End If
  Loop Until (Len(sChoice) = 0) Or (sChoice = "0")
  
  Close #mhOutput
  Shell "notepad " & sOutputFile, vbMaximizedFocus
  Exit Sub
Main_Err:
  MsgBox Err.Number & ": " & Err.Description, vbCritical
  If fIsOpen Then Close #mhOutput
End Sub

'
' CMapStringToLong test
'

Public Sub Test1()
  OutputBanner "Test1", _
    "Compare performance of a collection against the performance " & _
    "of the CMapStringToLong class"
  
  Dim cCollection   As Collection
  Dim mslMap        As CMapStringToLong
  Dim i             As Long
  Dim j             As Long
  Dim k             As Long
  Dim lRandom       As Long
  Dim lValue        As Long
  
  Set cCollection = New Collection
  Set mslMap = New CMapStringToLong
  
  '
  ' Adding a huge number of elements
  '
  Const klMaxElements As Long = 100000
  
  'Collection
  Output "Collection: Adding " & klMaxElements & "... "
  ChronoStart
  For i = 1 To klMaxElements
    cCollection.Add i, CStr(i)
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Adding " & klMaxElements & "... "
  ChronoStart
  For i = 1 To klMaxElements
    mslMap.Add CStr(i), i
  Next i
  mslMap.Sorted = True
  OutputLn ChronoTime() & " seconds."
  
  '
  ' Retrieval #1: by numerical index
  '
  Const klMaxTestElements As Long = 1000&
  Dim alIndex(1 To klMaxTestElements) As Long
  'Get random elements
  For i = 1 To klMaxTestElements
    k = 0&
    Do
      lRandom = GetRandom(1&, klMaxElements)
      For j = 1 To klMaxTestElements
        If alIndex(i) = lRandom Then k = i: Exit For
      Next j
    Loop Until k = 0&
    alIndex(i) = lRandom
  Next i
  
  'Collection
  Output "Collection: Retrieving (by numerical index) " & klMaxTestElements & " elements... "
  ChronoStart
  For i = 1 To klMaxTestElements
    lValue = cCollection(alIndex(i))
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Retrieving (by numerical index) " & klMaxTestElements & " elements... "
  ChronoStart
  For i = 1 To klMaxTestElements
    lValue = mslMap.Item(alIndex(i))
  Next i
  OutputLn ChronoTime() & " seconds."
  
  '
  ' Retrieval #2: by alphabetical index
  '
  'Collection
  Output "Collection: Retrieving (by key) " & klMaxTestElements & " elements... "
  ChronoStart
  For i = 1 To klMaxTestElements
    lValue = cCollection(CStr(alIndex(i)))
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Retrieving (by key) " & klMaxTestElements & " elements... "
  ChronoStart
  For i = 1 To klMaxTestElements
    lValue = mslMap.Item(mslMap.Find(CStr(alIndex(i))))
  Next i
  OutputLn ChronoTime() & " seconds."
   
  '
  ' Removing elements
  '
  'Get random elements again, but take them at the beginning of the
  'full range, otherwise, there is the risk that the element index
  'that will be removed is out of the valid bounds.
  For i = 1 To klMaxTestElements
    k = 0&
    Do
      lRandom = GetRandom(1&, klMaxElements \ 100&)
      For j = 1 To klMaxTestElements
        If alIndex(i) = lRandom Then k = i: Exit For
      Next j
    Loop Until k = 0&
    alIndex(i) = lRandom
  Next i
  
  'Collection
  Output "Collection: Removing " & klMaxTestElements & " elements... "
  ChronoStart
  For i = 1 To klMaxTestElements
    cCollection.Remove alIndex(i)
  Next i
  OutputLn ChronoTime() & " seconds."
  
  'CMapStringToLong
  Output "CMapStringToLong: Removing " & klMaxTestElements & " elements... "
  ChronoStart
  For i = 1 To klMaxTestElements
    mslMap.Remove alIndex(i)
  Next i
  OutputLn ChronoTime() & " seconds."
  
  '
  ' Destruction
  '
  Output "Collection: destroy... "
  ChronoStart
  Set cCollection = Nothing
  OutputLn ChronoTime() & " seconds."
  
  Output "CMapStringToLong: destroy... "
  ChronoStart
  Set mslMap = Nothing
  OutputLn ChronoTime() & " seconds."
End Sub

Sub Test2()
  OutputBanner "Test2", _
    "Prove that collection keys are case insensitive"
  
  'Prove that collection keys are case insensitive
  Dim cCollection As New Collection
  cCollection.Add Item:="Item1", Key:="iTeM1"
  On Error Resume Next
  cCollection.Add Item:="Item1", Key:="item1" 'this generates a runtime error
  If Err.Number Then
    OutputLn "Error #" & Err.Number & " occured while trying to add a duplicate key (but with different letter case) in a collection."
  End If
End Sub

Sub Test3()
  OutputBanner "Test3", _
    "Add some empty and duplicate keys to a CMapStringToInteger object, " & _
    "then remove duplicates."
  
  Dim mslMap    As New CMapStringToLong
  Dim i         As Long
  Dim k         As Long
  
  mslMap.Add "", 1&
  mslMap.Add "", 2&
  mslMap.Add "", 3&
  mslMap.Add "abc", 10&
  mslMap.Add "abc", 100&
  mslMap.Add "def", 20&
  mslMap.Add "def", 200&
  mslMap.Add "abc", 101&
  mslMap.Add "abc", 102&
  mslMap.Add "abc", 103&
  mslMap.Sorted = True
  
  OutputLn mslMap.Count & " items in set."
  OutputLn "one of the 'abc' found at position: " & mslMap.Find("abc")
  
  'print all items which key is "abc"
  OutputLn "All items which key is 'abc': "
  k = mslMap.FindFirst("abc")
  If k Then
    For i = k To mslMap.Count
      If mslMap.Key(i) = "abc" Then
        OutputLn i & ": " & mslMap.Key(i) & " --> " & mslMap.Item(i)
      Else
        Exit For
      End If
    Next i
  Else
    OutputLn "key 'abc' was not found by the FindFirst Function() ???"
  End If
  
  OutputLn "Removing duplicates..."
  mslMap.RemoveDuplicates
  OutputLn mslMap.Count & " items remaining in set: "
  For i = 1 To mslMap.Count
    OutputLn i & ": " & mslMap.Key(i) & " --> " & mslMap.Item(i)
  Next i
  
End Sub

'
' CRow test
'

'Populate column set of a CRow object using the AddCol method
Sub Row1()
  OutputBanner "Row1", _
    "Populate column set of a CRow object using the AddCol method"
  
  Dim oRow          As New CRow
  Dim iColPos       As Long
  
  'Define using the AddCol method
  With oRow
    .AddCol "ClientID", 1468&, 8&, 1&
    .AddCol "Name", "John Doe", 0&, 0&
    .AddCol "Address", "47 Main Street", 0&, 0&
    .AddCol "City", "Geneva", 0&, 0&
    .AddCol "State", "Switzerland", 0&, 0&
    'insert the Zip column after the Address column
    iColPos = .ColPos("Address")
    .AddCol "Zip", "12345", 0&, 0&, plInsertAfter:=iColPos
  End With
  RowDump oRow, "AddCol() method"
End Sub

'Defining column set Using "on the fly" arrays with the VB Array() method
Sub Row2()
  OutputBanner "Row2", _
    "Defining column set Using ""on the fly"" arrays with the VB Array() method"
  
  Dim oRow          As New CRow
  oRow.ArrayDefine Array("ClientID", "Name", "Address", "Zip", "City", "State"), _
                   Array(vbLong, vbString, vbString, vbString, vbString, vbString), _
                   Array(8&, 0&, 0&, 0&, 0&, 0&), _
                   Array(1&, 0&, 0&, 0&, 0&, 0&)
  oRow.Assign 1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland"
  RowDump oRow, "ArrayDefine() method"
End Sub

'Defining column set Using the Define method
Sub Row3()
  OutputBanner "Row3", _
    "Defining column set Using the Define method"
  
  Dim oRow          As New CRow
  'name, type, size, flags
  oRow.Define "ClientID", vbLong, 8&, 1&, _
              "Name", vbString, 0&, 0&, _
              "Address", vbString, 0&, 0&, _
              "Zip", vbString, 0&, 0&, _
              "City", vbString, 0&, 0&, _
              "State", vbString, 0&, 0&
  oRow.Assign 1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland"
  RowDump oRow, "Define() method"
End Sub

'Populate column set Using the ArrayAssign method
Sub Row4()
  OutputBanner "Row4", _
    "Populate column set Using the ArrayAssign method"
  
  Dim oRow          As New CRow
  'name, type, size, flags
  oRow.Define "ClientID", vbLong, 8&, 1&, _
              "Name", vbString, 0&, 0&, _
              "Address", vbString, 0&, 0&, _
              "Zip", vbString, 0&, 0&, _
              "City", vbString, 0&, 0&, _
              "State", vbString, 0&, 0&
  oRow.ArrayAssign Array(1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland")
  RowDump oRow, "Define() method"
End Sub

'Copying and cloning rows
Sub Row5()
  OutputBanner "Row5", _
    "Copying and cloning rows"
  
  Dim oRow1     As New CRow
  Dim oRow2     As New CRow
  Dim oRowClone As CRow
  
  oRow1.Define "ClientID", vbLong, 8&, 1&, _
               "Name", vbString, 0&, 0&, _
               "Address", vbString, 0&, 0&, _
               "Zip", vbString, 0&, 0&, _
               "City", vbString, 0&, 0&, _
               "State", vbString, 0&, 0&
  oRow1.Assign 1468&, "John Doe", "47 Main Street", "12345", "Geneva", "Switzerland"
  
  'Copy row
  oRow2.CopyFrom oRow1
  RowDump oRow2, "oRow2 copied from oRow1"
  
  'Clone row
  Set oRowClone = oRow1.Clone()
  RowDump oRowClone, "oRowClone created from oRow1"
End Sub

'Merging rows
Sub Row6()
  OutputBanner "Row6", _
    "Merging rows"
  
  Dim oRow1     As New CRow
  Dim oRow2     As New CRow
  
  oRow1.Define "ClientID", vbLong, 8&, 1&, _
               "Name", vbString, 0&, 0&, _
               "Address", vbString, 0&, 0&
  oRow1.Assign 1468&, "John Doe", "47 Main Street"
  
  oRow2.Define "Zip", vbString, 0&, 0&, _
               "Name", vbString, 0&, 0&, _
               "City", vbString, 0&, 0&, _
               "State", vbString, 0&, 0&
  oRow2.Assign "12345", "Patrick Doe", "Geneva", "Switzerland"
  
  'Merge Row2 into Row1
  'The name column of row1 wll be kept.
  oRow1.Merge oRow2
  RowDump oRow1, "oRow1 merged with oRow2"
End Sub

'Populate column set of a CRow object and access columns with # notation
Sub Row7()
  OutputBanner "Row7", _
    "Populate column set of a CRow object and access columns with # notation"
  
  Dim oRow          As New CRow
  Dim iColPos       As Long
  
  'Define using the AddCol method
  With oRow
    '.AddCol "ClientID", 1468&, 8&, 1&
    .AddCol "", 1468&, 8&, 1&
    .AddCol "Address", "47 Main Street", 0&, 0&
    '.AddCol "Name", "John Doe", 0&, 0&, plInsertAfter:=1&
    .AddCol "", "John Doe", 0&, 0&, plInsertAfter:=1&
    .AddCol "State", "Switzerland", 0&, 0&
    .AddCol "City", "Geneva", 0&, 0&, plInsertAfter:=3&
    'insert the Zip column after the Address column
    iColPos = .ColPos("Address")
    .AddCol "Zip", "12345", 0&, 0&, plInsertAfter:=iColPos
    
    RowDump oRow, "Row7 Sub"
    
    For iColPos = 1 To .ColCount
      OutputLn "ColValue('#" & iColPos & "')=" & .ColValue("#" & iColPos)
    Next iColPos
  End With
End Sub

'
' Support functions
'

'Get a random long
Function GetRandom(ByVal iLo As Long, ByVal iHi As Long) As Long
  GetRandom = Int(iLo + (Rnd * (iHi - iLo + 1)))
End Function

'
' A precise chronometer for accurate timing
'

#If Win64 Then
Private Function U64Dbl(U64 As UINT64) As Double
    Dim lDbl As Double, hDbl As Double
    lDbl = U64.LowPart
    hDbl = U64.HighPart
    If lDbl < 0 Then lDbl = lDbl + BSHIFT_32
    If hDbl < 0 Then hDbl = hDbl + BSHIFT_32
    U64Dbl = lDbl + BSHIFT_32 * hDbl
End Function
#End If

Public Sub ChronoStart()
  #If Win64 Then
    If (mcurFrequency.HighPart = 0) And (mcurFrequency.LowPart = 0) Then QueryPerformanceFrequency mcurFrequency
  #Else
    If mcurFrequency = 0 Then QueryPerformanceFrequency mcurFrequency
  #End If
  QueryPerformanceCounter mcurChronoStart
End Sub

Public Function ChronoTime() As String
#If Win64 Then
  Dim curFrequency As UINT64
  Dim dblElapsed   As Double
  QueryPerformanceCounter curFrequency
  If (mcurFrequency.LowPart = 0) And (mcurFrequency.HighPart = 0) Then
    curFrequency.LowPart = 0
    curFrequency.HighPart = 0
  Else
    dblElapsed = (U64Dbl(curFrequency) - U64Dbl(mcurChronoStart)) / U64Dbl(mcurFrequency)
  End If
  ChronoTime = CStr(dblElapsed)
#Else
  Dim curFrequency As Currency
  QueryPerformanceCounter curFrequency
  If mcurFrequency = 0 Then
    curFrequency = 0
  Else
    curFrequency = (curFrequency - mcurChronoStart) / mcurFrequency
  End If
  ChronoTime = CStr(curFrequency)
#End If
End Function

'
' Dumping CRow and CList objects
'

Sub ShowProgressDialog()
  'Note: Not used in this project
  'frmMessage.Show
End Sub

Sub CloseProgressDialog()
  'Note: Not used in this project
  'Unload frmMessage
End Sub

Sub ProgressMsg(ByRef sMessage As String)
  'Note: Not used in this project
  'frmMessage.lblMessage = sMessage
  'frmMessage.Refresh
  DoEvents
End Sub

Sub OutputLn(Optional ByRef sOutput As String = "")
  If mhOutput Then
    Print #mhOutput, sOutput
  Else
    Debug.Print sOutput
  End If
End Sub

Sub Output(ByRef sOutput As String)
  If mhOutput Then
    Print #mhOutput, sOutput;
  Else
    Debug.Print sOutput;
  End If
End Sub

Sub OutputBanner(ByVal sSubName As String, ByVal sDescr As String)
  Dim sBanner   As String
  
  sBanner = vbCrLf & String$(60, "=") & vbCrLf & sSubName & vbCrLf & sDescr & vbCrLf & String$(60, "=") & vbCrLf
  OutputLn sBanner
  ProgressMsg "Running " & sSubName & vbCrLf & vbCrLf & sDescr
End Sub

Public Sub RowDump(oRow As CRow, Optional ByVal sTitle As String = "")
  Dim iRow      As Long
  Dim i         As Long
  Dim lCount    As Long
  Dim asColName()  As String
  Dim iLen      As Integer
  
  lCount = oRow.ColCount
  If Len(sTitle) Then
    OutputLn String$(Len(sTitle), "-") & "+"
    OutputLn sTitle & "|"
  End If
  'col titles row sep
  For i = 1 To lCount
    iLen = Len(oRow.ColName(i))
    Output String$(iLen, "-") & "+"
  Next i
  OutputLn
  For i = 1 To lCount
    Output oRow.ColName(i) & "|"
  Next i
  OutputLn
  'col titles row sep
  For i = 1 To lCount
    iLen = Len(oRow.ColName(i))
    Output String$(iLen, "-") & "+"
  Next i
  OutputLn
  
  'dump values
  For i = 1 To oRow.ColCount
    iLen = Len(oRow.ColName(i))
    If Not IsNull(oRow.ColValue(i)) Then
      Output StrBlock(oRow.ColValue(i) & "", " ", iLen) & "|"
    Else
      Output StrBlock("#NULL", " ", iLen) & "|"
    End If
  Next i
  OutputLn
End Sub

Public Sub ListDump(oList As CList, Optional ByVal sTitle As String = "", Optional ByVal psColWidths As String = "", Optional ByVal plStartRow As Long = 0&, Optional ByVal plEndRow As Long = 0&)
  Dim iRow      As Long
  Dim i         As Long
  Dim lCount    As Long
  Dim asColName()  As String
  Dim aiColWidth() As Integer
  Dim iLen      As Integer
  Dim iStart    As Long
  Dim iEnd      As Long
  
  On Error GoTo ListDump_Err
  
  lCount = oList.ColCount
  If lCount = 0& Then Exit Sub
  ReDim aiColWidth(1 To lCount)
  If Len(sTitle) Then
    OutputLn String$(Len(sTitle), "-") & "+"
    OutputLn sTitle & "|"
  End If
  If Len(psColWidths) Then
    Dim iColWidthCt       As Integer
    Dim asColWidthSpec()  As String
    Dim iCol              As Integer
    Dim sColName          As String
    Dim sWidth            As String
    Dim iColon            As Integer
    iColWidthCt = SplitString(asColWidthSpec(), psColWidths, ";")
    For i = 1 To iColWidthCt
      iColon = InStr(1, asColWidthSpec(i), ":")
      If iColon Then
        sColName = Left$(asColWidthSpec(i), iColon - 1)
        sWidth = Right$(asColWidthSpec(i), Len(asColWidthSpec(i)) - iColon)
        If Len(sWidth) > 0 Then
          If Val(sWidth) > 0 Then
            iCol = oList.ColPos(sColName)
            If iCol Then
              aiColWidth(iCol) = Val(sWidth)
            End If
          End If
        End If
      End If
    Next i
  End If
  'col titles row sep
  For i = 1 To lCount
    iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
    Output String$(iLen, "-") & "+"
  Next i
  OutputLn
  For i = 1 To lCount
    iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
    Output StrBlock(oList.ColName(i), " ", iLen) & "|"
  Next i
  OutputLn
  'col titles row sep
  For i = 1 To lCount
    iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
    Output String$(iLen, "-") & "+"
  Next i
  OutputLn
  
  'dump values
  iStart = 1&
  iEnd = oList.Count
  If plStartRow > 0& Then
    If plStartRow <= oList.Count Then
      iStart = plStartRow
    End If
  End If
  If plEndRow > 0& Then
    If plEndRow <= oList.Count Then
      iEnd = plEndRow
    End If
  End If
  If iStart > iEnd Then
    'Swap
    Dim iTemp As Long
    iTemp = iEnd
    iEnd = iStart
    iStart = iTemp
  End If
  
  For iRow = iStart To iEnd
    For i = 1 To oList.ColCount
      iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
      If Not IsNull(oList.Item(i, iRow)) Then
        Output StrBlock(oList.Item(i, iRow) & "", " ", iLen) & "|"
      Else
        Output StrBlock("#NULL", " ", iLen) & "|"
      End If
    Next i
    OutputLn
  Next iRow
  
  Exit Sub
ListDump_Err:
  'Stop
  'Resume
End Sub

Public Function StrBlock(ByVal sText As String, ByVal sPadChar As String, ByVal iMaxLen As Integer) As String
  Dim iLen      As Integer
  
  iLen = Len(sText)
  If iLen < iMaxLen Then
    StrBlock = sText & String$(iMaxLen - iLen, sPadChar)
  Else
    If iMaxLen > 6 Then
      StrBlock = Left$(sText, iMaxLen - 3) & "..."
    Else
      StrBlock = Left$(sText, iMaxLen)
    End If
  End If
End Function

'Returns a string representing a valid long integer (typed by the user), or
'an empty string if the user cancels the action.
'To get the returned string into a long (if not empty), you can safely use:
'Clng(sReturnedString) which will never raise an error.
Public Function LongChooseBox(ByVal sText As String, _
                              ByVal sTitle As String, _
                              ByVal sDefault As String, _
                              ByVal lMax As Long, _
                              Optional ByVal lMin As Long = 1&) As String
  Dim sInput    As String
  Dim lRet      As Long
  
  Do
    sInput = InputBox$(sText, sTitle, sDefault)
    If Len(sInput) Then
      If CheckLong(sInput) = 0& Then
        lRet = CDbl(sInput)
        'Its a long
        If (lRet >= lMin) And (lRet <= lMax) Then
          LongChooseBox = CStr(lRet)
          Exit Function
        Else
          MsgBox "Please type a number between " & lMin & " and " & lMax, vbCritical
        End If
      Else
        Select Case lRet
        Case 2& 'String is too long to represent a signed long
          MsgBox "The number you typed is too big", vbCritical
        Case 3& 'bad character in string
          MsgBox "There's an invalid character in the string you typed", vbCritical
        End Select
      End If
    End If
    'Present the user what he previously type
    sDefault = sInput
  Loop Until (sInput = "")
End Function

'Returns:
'0& : if sValue represents a valid long integer
'1& : if sValue is empty
'2& : String is too long to represent a signed long
'3& : bad character in string
Public Function CheckLong(ByVal sValue As String) As Long
  Dim iLen      As Integer
  Dim i         As Integer
  Dim sChar     As String
  Dim iAsc      As Integer
  Dim iAscZero  As Integer
  Dim iAscNine  As Integer
  Dim iAscPlus  As Integer
  Dim iAscMinus As Integer
  
  iAscZero = Asc("0")
  iAscNine = Asc("9")
  iAscPlus = Asc("+")
  iAscMinus = Asc("-")
  
  iLen = Len(sValue)
  If iLen = 0 Then
    CheckLong = 1&  'string is empty
    Exit Function
  End If
  If iLen > 11 Then
    CheckLong = 2&  'string too long
    Exit Function
  End If
  
  For i = 1 To iLen
    sChar = Mid$(sValue, i, 1)
    iAsc = Asc(sChar)
    If i = 1 Then
      '((iAsc<iAscZero) or (iAsc>iAscNine)) : char is not a number
      '(iAsc<>iAscPlus) and (iAsc<>iAscMinus) but can be + or -
      If ((iAsc < iAscZero) Or (iAsc > iAscNine)) And _
         (iAsc <> iAscPlus) And (iAsc <> iAscMinus) Then
        CheckLong = 3&  'bad character
        Exit Function
      End If
    Else
      '((iAsc<iAscZero) or (iAsc>iAscNine)) : char is not a number
      If ((iAsc < iAscZero) Or (iAsc > iAscNine)) Then
        CheckLong = 3&  'bad character
        Exit Function
      End If
    End If
  Next i
End Function

