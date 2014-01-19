Attribute VB_Name = "MStrings"
Option Explicit

Public Const sBSlash    As String = "\"

Public Function IsWhite(ByVal sChar As String) As Boolean
  Dim iAsc    As Integer
  iAsc = Chr$(sChar)
  If (Len(sChar) > "") Then
    iAsc = Chr$(sChar)
    IsWhite = (iAsc = 32) Or (iAsc = 9) Or (iAsc = 13) Or (iAsc = 10) Or (iAsc = 160)
  End If
End Function

Public Function TrimAllWhites(ByVal sStrip As String) As String
  If sStrip > "" Then
    While IsWhite(Left$(sStrip, 1))
      sStrip = Right$(sStrip, Len(sStrip) - 1)
    Wend
    While IsWhite(Right$(sStrip, 1))
      sStrip = Left$(sStrip, Len(sStrip) - 1)
    Wend
  End If
  TrimAllWhites = sStrip
End Function

Public Function TrimChar(ByVal sStrip As String, ByVal sTrimChar As String) As String
  If sStrip > "" Then
    While Left$(sStrip, 1) = sTrimChar
      sStrip = Right$(sStrip, Len(sStrip) - 1)
    Wend
    While Right$(sStrip, 1) = sTrimChar
      sStrip = Left$(sStrip, Len(sStrip) - 1)
    Wend
  End If
  TrimChar = sStrip
End Function

Function NbObj(ByRef sSource As String, ByRef sSeps As String) As Integer
  Dim n           As Integer
  Dim p           As Integer
  Dim sBreak      As String
  Dim iBreakLen   As Integer
  
  sBreak = sSeps
  iBreakLen = Len(sBreak)

  'Remove any leading / trailing sBreak
  While Left$(sSource, iBreakLen) = sBreak
      sSource = Right$(sSource, Len(sSource) - iBreakLen)
  Wend
  While Right$(sSource, iBreakLen) = sBreak
    sSource = Left$(sSource, Len(sSource) - iBreakLen)
  Wend

  'Count
  p = InStr(sSource, sBreak)
  While p
    n = n + 1
    p = InStr(p + iBreakLen, sSource, sBreak)
  Wend
  If n = 0 Then
    If sSource <> "" Then
      NbObj = 1
    End If
  Else
    NbObj = n + 1
  End If
End Function

Function GetObj(ByVal iObjIndex As Integer, ByRef sSeps As String, ByRef sObjects As String) As String
  Dim i           As Integer
  Dim p           As Integer
  Dim pp          As Integer
  Dim iBad        As Integer
  Dim sBreak      As String
  Dim sTmpObjs    As String
  Dim iBreakLen   As Integer
  
  sTmpObjs = sObjects
  sBreak = sSeps
  iBreakLen = Len(sBreak)
  
  'par sécurité, je remove les leading et trailing sBreak
  While Left$(sTmpObjs, iBreakLen) = sBreak
    sTmpObjs = Right$(sTmpObjs, Len(sTmpObjs) - iBreakLen)
  Wend
  While Right$(sTmpObjs, iBreakLen) = sBreak
    sTmpObjs = Left$(sTmpObjs, Len(sTmpObjs) - iBreakLen)
  Wend

  If iObjIndex > 1 Then
    For i = 1 To iObjIndex - 1
      p = InStr(p + 1, sTmpObjs, sBreak)
      If p = 0 Then iBad = True: Exit For
    Next i
    If Not iBad Then
      pp = InStr(p + iBreakLen, sTmpObjs, sBreak)
      If pp Then
        GetObj = Mid$(sTmpObjs, p + iBreakLen, pp - p - iBreakLen)
      Else
        GetObj = Right$(sTmpObjs, Len(sTmpObjs) - p - iBreakLen + 1)
      End If
    End If
  Else
    p = InStr(sTmpObjs, sBreak)
    If p Then
      GetObj = Left$(sTmpObjs, p - 1)
    Else
      GetObj = sTmpObjs
    End If
  End If
End Function

Sub sReplace(ByRef sSource As String, ByVal sPattern As String, ByVal sReplace As String)
  Dim i         As Integer
  Dim lR        As Integer
  Dim lgReplace As Integer
  
  If sSource = "" Then Exit Sub
  If sPattern = "" Then Exit Sub
  
  i = InStr(sSource, sPattern)
  lR = Len(sPattern)
  lgReplace = Len(sReplace): If lgReplace = 0 Then lgReplace = 1
  While i
    sSource = Left$(sSource, i - 1) + sReplace + Right$(sSource, Len(sSource) - (i + lR - 1))
    i = InStr(i + lgReplace, sSource, sPattern)
  Wend
End Sub

Public Function CtoVB(sBuf) As String
  Dim i     As Integer
  i = InStr(sBuf, Chr$(0))
  If i > 0 Then
    CtoVB = Left$(sBuf, i - 1)
  Else
    CtoVB = sBuf
  End If
End Function

Public Function USASCII(ByRef psMixedCase As String) As String
  Dim i       As Integer
  Dim iLen    As Integer
  Dim sRet    As String
  Dim sChar   As String
  
  iLen = Len(psMixedCase)
  If iLen Then
    For i = 1 To iLen
      sChar = Mid$(psMixedCase, i, 1)
      Select Case sChar
      Case "à", "á", "ä", "â"
          sChar = "A"
      Case "é", "è", "ë", "ê"
          sChar = "E"
      Case "í", "ì", "ï", "î"
          sChar = "I"
      Case "ó", "ò", "ö", "ô"
          sChar = "O"
      Case "ú", "ù", "ü", "û"
          sChar = "U"
      Case "ç"
          sChar = "C"
      Case Else
          sChar = UCase$(sChar)
      End Select
      sRet = sRet & sChar
    Next i
    USASCII = sRet
  End If
End Function

' Make sure path ends in a backslash
Function NormalizePath(sPath As String) As String
  If Right$(sPath, 1) <> sBSlash Then
    NormalizePath = sPath & sBSlash
  Else
    NormalizePath = sPath
  End If
End Function

' Make sure path doesn't end in a backslash
Function DenormalizePath(sPath As String) As String
  If Right$(sPath, 1) = sBSlash Then
    DenormalizePath = Left$(sPath, Len(sPath) - 1)
  Else
    DenormalizePath = sPath
  End If
End Function

Public Function StripFileExt(ByVal sFile As String) As String
  Dim sExt    As String
  Dim iDot    As Integer
  
  On Error Resume Next
  sExt = Right$(sFile, 4)
  iDot = InStr(1, sExt, ".")
  If iDot Then
    sExt = Right$(sExt, Len(sExt) - iDot)
    StripFileExt = Left$(sFile, Len(sFile) - Len(sExt) - 1)
  Else
    StripFileExt = sFile
  End If
End Function

Function StripFileName(ByVal sPath As String) As String
  Dim iIndex    As Integer
  Dim iLoop     As Integer
  Dim sChar     As String * 1
  
  If (InStr(sPath, ":") = 0) And (Len(sPath) < 13) Then Exit Function
  
  iIndex = Len(sPath)
  If iIndex Then iLoop = True
  While iLoop
    If iIndex > 0 Then
      sChar = Mid$(sPath, iIndex, 1)
      If (sChar = "\") Or (sChar = ":") Then
        iLoop = False
      End If
    End If
    If iIndex > 1 Then
      iIndex = iIndex - 1
    Else
      iIndex = 0
      iLoop = False
    End If
  Wend
  If iIndex Then
    StripFileName = Left$(sPath, iIndex)
  Else
    StripFileName = ""
  End If
End Function

Function StripFilePath(ByVal sFileName As String) As String
  Dim iIndex    As Integer
  Dim sChar     As String * 1
  iIndex = Len(sFileName)
  If iIndex Then
    sChar = Mid$(sFileName, iIndex, 1)
    While (sChar <> ":") And (sChar <> "\") And (iIndex > 0)
      iIndex = iIndex - 1
      If iIndex Then
          sChar = Mid$(sFileName, iIndex, 1)
      Else
          sChar = sBSlash
      End If
    Wend
    If iIndex Then
      StripFilePath = Right$(sFileName, Len(sFileName) - iIndex)
    Else
      StripFilePath = sFileName
    End If
  End If
End Function

Function StrLPad(ByVal s As String, ByVal PadChar As String, ByVal iLen As Integer) As String
  If iLen Then
    If Len(s) < iLen Then
      StrLPad = String$(iLen - Len(s), PadChar) & s
    Else
      StrLPad = Left$(s, iLen)
    End If
  End If
End Function

Function StrRPad(ByVal s As String, ByVal PadChar As String, ByVal iLen As Integer) As String
  If Len(s) < iLen Then
    StrRPad = s & String$(iLen - Len(s), Asc(PadChar))
  Else
    StrRPad = Left$(s, iLen)
  End If
End Function

'Split         Split a string into a variant array.
'
'InStrRev      Similar to InStr but searches from end of string.
'
'Replace       To find a particular string and replace it.
'
'Reverse       To reverse a string.

Public Function InStrRev(ByVal sIn As String, ByVal _
   sFind As String, Optional nStart As Long = 1, _
    Optional bCompare As VbCompareMethod = vbBinaryCompare) _
    As Long

    Dim nPos As Long
    
    sIn = Reverse(sIn)
    sFind = Reverse(sFind)
    
    nPos = InStr(nStart, sIn, sFind, bCompare)
    If nPos = 0 Then
        InStrRev = 0
    Else
        InStrRev = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function

Public Function Reverse(ByVal sIn As String) As String
    Dim nC As Long
    Dim sOut As String

    For nC = Len(sIn) To 1 Step -1
        sOut = sOut & Mid(sIn, nC, 1)
    Next nC
    
    Reverse = sOut
End Function

Public Function Join(Source() As String, _
    Optional sDelim As String = " ") As String

    Dim nC As Long
    Dim sOut As String
    
    For nC = LBound(Source) To UBound(Source) - 1
        sOut = sOut & Source(nC) & sDelim
    Next
    
    Join = sOut & Source(nC)
End Function

Public Function Replace(ByVal sIn As String, ByVal sFind As _
    String, ByVal sReplace As String, Optional nStart As _
     Long = 1, Optional nCount As Long = -1, _
     Optional bCompare As VbCompareMethod = vbBinaryCompare) As _
     String

  Dim nC As Long, nPos As Long
  Dim nFindLen As Long, nReplaceLen As Long

  nFindLen = Len(sFind)
  nReplaceLen = Len(sReplace)
  
  If (sFind <> "") And (sFind <> sReplace) Then
    nPos = InStr(nStart, sIn, sFind, bCompare)
    Do While nPos
        nC = nC + 1
        sIn = Left(sIn, nPos - 1) & sReplace & _
         Mid(sIn, nPos + nFindLen)
        If nCount <> -1 And nC >= nCount Then Exit Do
        nPos = InStr(nPos + nReplaceLen, sIn, sFind, _
          bCompare)
    Loop
  End If

  Replace = sIn
End Function

Public Function Split(ByVal sIn As String, _
  Optional sDelim As String = " ", _
  Optional nLimit As Long = -1, _
  Optional bCompare As VbCompareMethod = vbBinaryCompare) _
  As Variant

  Dim nC As Long, nPos As Long, nDelimLen As Long
  Dim sOut() As String
  
  If sDelim <> "" Then
    nDelimLen = Len(sDelim)
    nPos = InStr(1, sIn, sDelim, bCompare)
    Do While nPos
      nC = nC + 1
      ReDim Preserve sOut(1 To nC)
      sOut(nC) = Left(sIn, nPos - 1)
      sIn = Mid(sIn, nPos + nDelimLen)
      If nLimit <> -1 And nC >= nLimit Then Exit Do
      nPos = InStr(1, sIn, sDelim, bCompare)
    Loop
  End If

  ReDim Preserve sOut(1 To nC)
  sOut(nC) = sIn

  Split = sOut
End Function

Public Function DecimalSeparator() As String
  DecimalSeparator = Mid$(1 / 2, 2, 1)
End Function

'Splitte une chaine de car. en fonction de sSep, dans asDest et retourne le nbre d'élem.
Public Function SplitString(asDest() As String, ByVal sSplit As String, ByVal sSep As String) As Integer
  Dim nObj        As Integer
  Dim i           As Integer
  
  nObj = NbObj(sSplit, sSep)
  If nObj Then
    ReDim asDest(1 To nObj)
    For i = 1 To nObj
      asDest(i) = GetObj(i, sSep, sSplit)
    Next i
  Else
    On Error Resume Next
    Erase asDest
  End If
  
  SplitString = nObj
End Function

Public Function StrBlock(ByVal sText As String, ByVal sPadChar As String, ByVal iMaxLen As Integer) As String
  Dim iLen      As Integer
  
  iLen = Len(sText)
  If iLen <= iMaxLen Then
    StrBlock = sText & String$(iMaxLen - iLen, sPadChar)
  Else
    If iMaxLen > 6 Then
      StrBlock = Left$(sText, iMaxLen - 3) & "..."
    Else
      StrBlock = Left$(sText, iMaxLen)
    End If
  End If
End Function
