Attribute VB_Name = "MRowList"
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'MRowList.bas module
'This module contains general purpose functions available for internal
'use by the library.
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

' 5001 - 5010: Shared between object implementing IObjectBytes
Public Const kErrBadClassIDBytes      As Long = 5001& '%1%=class name
Public Const kErrBadClassVerBytes     As Long = 5002& '%1%=class name

Public Const klObjectErrBase          As Long = 6000&   'Which leaves us from 6000& to 29000& for our errors

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub CopyMemoryToString Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpstrDest As String, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub CopyMemoryFromString Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpstrSource As String, ByVal cbCopy As Long)

'A version string has 3 parts : XX.YY.ZZ
' XX: major
' YY: minor
' ZZ: revision
Public Function MAKE_VERSIONLONG(ByVal psVersion As String) As Long
  MAKE_VERSIONLONG = CLng("1" & LPad$(Replace(psVersion, ".", ""), "0", 6))
End Function

'Call in VB error trapping routines
Public Function MAKE_VBERROR(ByVal plErrCode As Long) As Long
  MAKE_VBERROR = vbObjectError Or (plErrCode And &HFFFF&)
End Function

'Call when setting a custom object err code
Public Function MAKE_OBJECTERROR(ByVal plErrCode As Long) As Long
  MAKE_OBJECTERROR = vbObjectError Or (plErrCode + klObjectErrBase)
End Function

'Split a string into a new array.
'Returns the number of elements in the array.
Public Function SplitString(ByRef asRetItems() As String, _
  ByVal sToSplit As String, _
  Optional sSep As String = " ", _
  Optional lMaxItems As Long = 0&, _
  Optional eCompare As VbCompareMethod = vbBinaryCompare) _
  As Long

  Dim lPos        As Long
  Dim lDelimLen   As Long
  Dim lRetCount   As Long
  
  On Error Resume Next
  Erase asRetItems
  On Error GoTo SplitString_Err
  
  If Len(sToSplit) Then
    lDelimLen = Len(sSep)
    If lDelimLen Then
      lPos = InStr(1, sToSplit, sSep, eCompare)
      Do While lPos
        lRetCount = lRetCount + 1&
        ReDim Preserve asRetItems(1& To lRetCount)
        asRetItems(lRetCount) = Left$(sToSplit, lPos - 1&)
        sToSplit = Mid$(sToSplit, lPos + lDelimLen)
        If lMaxItems Then
          If lRetCount = lMaxItems - 1& Then Exit Do
        End If
        lPos = InStr(1, sToSplit, sSep, eCompare)
      Loop
    End If
    lRetCount = lRetCount + 1&
    ReDim Preserve asRetItems(1& To lRetCount)
    asRetItems(lRetCount) = sToSplit
  End If
  SplitString = lRetCount
SplitString_Err:
End Function

'VB6 compatible function
Public Function Replace(ByVal sText As String, _
                        ByVal sReplaceWhat As String, _
                        ByVal sReplaceBy As String, _
                        Optional lStartPos As Long = 1&, _
                        Optional lMaxReplaces As Long = 0&, _
                        Optional eCompare As VbCompareMethod = vbBinaryCompare) As String
  Dim lCount      As Long
  Dim lFindPos    As Long
  Dim lFindLen    As Long
  Dim lReplaceLen As Long

  lFindLen = Len(sReplaceWhat)
  lReplaceLen = Len(sReplaceBy)
  
  If CBool(Len(sReplaceWhat)) And CBool(StrComp(sReplaceWhat, sReplaceBy, eCompare)) Then
    lFindPos = InStr(lStartPos, sText, sReplaceWhat, eCompare)
    Do While lFindPos
      lCount = lCount + 1&
      sText = Left(sText, lFindPos - 1&) & sReplaceBy & Mid(sText, lFindPos + lFindLen)
      If lMaxReplaces Then
        If lCount = lMaxReplaces - 1& Then Exit Do
      End If
      lFindPos = InStr(lFindPos + lReplaceLen, sText, sReplaceWhat, eCompare)
    Loop
  End If

  Replace = sText
End Function

Function LPad(ByVal s As String, ByVal PadChar As String, ByVal iLen As Integer) As String
  If iLen Then
    If Len(s) < iLen Then
      LPad = String$(iLen - Len(s), PadChar) & s
    Else
      LPad = Left$(s, iLen)
    End If
  End If
End Function


