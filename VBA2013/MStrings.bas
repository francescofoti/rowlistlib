Attribute VB_Name = "MStrings"
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'MStrings.bas module
'This module contains general purpose functions working on strings
'available for internal use by the library.
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
Option Compare Database
Option Explicit

Function LPad(ByVal psToPad As String, ByVal psPadChar As String, ByVal piLen As Integer) As String
  If piLen Then
    If Len(psToPad) < piLen Then
      LPad = String$(piLen - Len(psToPad), psPadChar) & psToPad
    Else
      LPad = Left$(psToPad, piLen)
    End If
  End If
End Function

Function CountSplittableItems(ByRef psSource As String, ByRef psSeps As String) As Integer
  Dim n           As Integer
  Dim p           As Integer
  Dim sBreak      As String
  Dim iBreakLen   As Integer
  
  sBreak = psSeps
  iBreakLen = Len(sBreak)

  'Remove any leading / trailing sBreak
  While Left$(psSource, iBreakLen) = sBreak
    psSource = Right$(psSource, Len(psSource) - iBreakLen)
  Wend
  While Right$(psSource, iBreakLen) = sBreak
  psSource = Left$(psSource, Len(psSource) - iBreakLen)
  Wend

  'Count
  p = InStr(psSource, sBreak)
  While p
    n = n + 1
    p = InStr(p + iBreakLen, psSource, sBreak)
  Wend
  If n = 0 Then
    If Len(psSource) > 0 Then
      CountSplittableItems = 1
    End If
  Else
    CountSplittableItems = n + 1
  End If
End Function

Function SplittedItem(ByVal piSplitItem As Integer, ByRef psSeps As String, ByRef psItems As String, Optional ByVal pfTrimSeps As Boolean = True) As String
  Dim i           As Long
  Dim iBreak      As Long
  Dim iBreak2     As Long
  Dim fBad        As Boolean
  Dim sBreak      As String
  Dim sItemsCopy  As String
  Dim lBreakLen   As Long
  
  sItemsCopy = psItems
  sBreak = psSeps
  lBreakLen = Len(sBreak)
  
  If pfTrimSeps Then
    While Left$(sItemsCopy, lBreakLen) = sBreak
      sItemsCopy = Right$(sItemsCopy, Len(sItemsCopy) - lBreakLen)
    Wend
    While Right$(sItemsCopy, lBreakLen) = sBreak
      sItemsCopy = Left$(sItemsCopy, Len(sItemsCopy) - lBreakLen)
    Wend
  End If

  If piSplitItem > 1& Then
    For i = 1 To piSplitItem - 1&
      iBreak = InStr(iBreak + 1&, sItemsCopy, sBreak)
      If iBreak = 0& Then fBad = True: Exit For
    Next i
    If Not fBad Then
      iBreak2 = InStr(iBreak + lBreakLen, sItemsCopy, sBreak)
      If iBreak2 Then
        SplittedItem = Mid$(sItemsCopy, iBreak + lBreakLen, iBreak2 - iBreak - lBreakLen)
      Else
        SplittedItem = Right$(sItemsCopy, Len(sItemsCopy) - iBreak - lBreakLen + 1&)
      End If
    End If
  Else
    iBreak = InStr(sItemsCopy, sBreak)
    If iBreak Then
      SplittedItem = Left$(sItemsCopy, iBreak - 1&)
    Else
      SplittedItem = sItemsCopy
    End If
  End If
End Function

'Splits astring on psSep separator, putting splitted items into pasDest and returning the # of items
Public Function SplitString(pasDest() As String, ByVal psSplit As String, ByVal psSep As String) As Integer
  Dim iItemCount  As Integer
  Dim i           As Integer
  
  iItemCount = CountSplittableItems(psSplit, psSep)
  If iItemCount Then
    ReDim pasDest(1 To iItemCount)
    For i = 1 To iItemCount
      pasDest(i) = SplittedItem(i, psSep, psSplit)
    Next i
  Else
    On Error Resume Next
    Erase pasDest
  End If
  
  SplitString = iItemCount
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

