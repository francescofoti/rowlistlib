Attribute VB_Name = "MMemory"
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'MMemory.bas module
'This code module is an API for reading and writing VB variables from
'and to byte arrays (peek and poke versions).
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
'09/06/2007 ¦ 01.00.00 ¦ FFO ¦ Tagging as inserting into new shared library make up.
'01/08/2013 ¦ 01.01.00 ¦ FFO ¦ 32/64 bits compatible edition for VBA
'           ¦          ¦     ¦
'           ¦          ¦     ¦
'           ¦          ¦     ¦
Option Explicit

Public Const klSizeOfLong      As Long = 4&
Public Const klSizeOfInt       As Long = 2&
Public Const klSizeOfBool      As Long = 2&
Public Const klSizeOfByte      As Long = 1&

'Peek a variant value from a byte array and advance

Public Function PeekByte(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Byte
  PeekByte = pabSource(plBytePtr)
  plBytePtr = plBytePtr + klSizeOfByte
End Function

Public Function PeekInteger(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Integer
  Dim iInt As Integer
  CopyMemory iInt, pabSource(plBytePtr), klSizeOfInt
  plBytePtr = plBytePtr + klSizeOfInt
  PeekInteger = iInt
End Function

Public Function PeekLong(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Long
  Dim lLong As Long
  CopyMemory lLong, pabSource(plBytePtr), klSizeOfLong
  plBytePtr = plBytePtr + klSizeOfLong
  PeekLong = lLong
End Function

Public Function PeekSingle(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Single
  Dim sngValue As Single
  CopyMemory sngValue, pabSource(plBytePtr), LenB(sngValue)
  plBytePtr = plBytePtr + LenB(sngValue)
  PeekSingle = sngValue
End Function

Public Function PeekDouble(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Double
  Dim dblDouble As Double
  CopyMemory dblDouble, pabSource(plBytePtr), LenB(dblDouble)
  plBytePtr = plBytePtr + LenB(dblDouble)
  PeekDouble = dblDouble
End Function

Public Function PeekCurrency(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Currency
  Dim curValue As Currency
  CopyMemory curValue, pabSource(plBytePtr), LenB(curValue)
  plBytePtr = plBytePtr + LenB(curValue)
  PeekCurrency = curValue
End Function

Public Function PeekDate(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Date
  Dim dtValue As Date
  CopyMemory dtValue, pabSource(plBytePtr), LenB(dtValue)
  plBytePtr = plBytePtr + LenB(dtValue)
  PeekDate = dtValue
End Function

Public Function PeekBoolean(ByRef pabSource() As Byte, ByRef plBytePtr As Long) As Boolean
  Dim fBool As Boolean
  CopyMemory fBool, pabSource(plBytePtr), LenB(fBool)
  plBytePtr = plBytePtr + LenB(fBool)
  PeekBoolean = fBool
End Function

Public Function PeekString(ByRef pabData() As Byte, ByRef plBytePtr As Long) As String
  Dim lLen      As Long
  Dim iByte     As Long
  Dim sBuffer   As String
  
  CopyMemory lLen, pabData(plBytePtr), klSizeOfLong
  plBytePtr = plBytePtr + klSizeOfLong
  If lLen Then
    sBuffer = Space$(lLen)
    CopyMemoryToString sBuffer, pabData(plBytePtr), lLen
    PeekString = sBuffer
    plBytePtr = plBytePtr + lLen
  End If
End Function

Public Sub GetVariant(ByRef pvVariant As Variant, ByRef plBytePtr As Long, ByRef pabSource() As Byte)
  Dim iVarType  As Integer
  
  CopyMemory iVarType, pabSource(plBytePtr), klSizeOfInt
  plBytePtr = plBytePtr + klSizeOfInt
  If (iVarType And vbArray) = 0 Then
    Select Case iVarType
    Case vbEmpty      '0    Empty (uninitialized)
    Case vbNull       '1    Null (no valid data)
    Case vbInteger    '2    Integer
      pvVariant = PeekInteger(pabSource(), plBytePtr)
    Case vbLong       '3    Long integer
      pvVariant = PeekLong(pabSource(), plBytePtr)
    Case vbSingle     '4    Single-precision floating-point number
      pvVariant = PeekSingle(pabSource(), plBytePtr)
    Case vbDouble     '5    Double-precision floating-point number
      pvVariant = PeekDouble(pabSource(), plBytePtr)
    Case vbCurrency   '6    Currency value
      pvVariant = PeekCurrency(pabSource(), plBytePtr)
    Case vbDate       '7    Date value
      pvVariant = PeekDate(pabSource(), plBytePtr)
    Case vbString     '8    String
      pvVariant = PeekString(pabSource(), plBytePtr)
    Case vbBoolean    '11   Boolean value
      pvVariant = PeekBoolean(pabSource(), plBytePtr)
    Case vbDecimal    '14   Decimal value
      'Unsupported, borrow VB error 13&
      Err.Raise MAKE_VBERROR(13&), "MMemory::GetVariant", "Variant Decimal subtype unsupported"
    Case vbByte       '17   Byte value
      pvVariant = pabSource(plBytePtr)
      plBytePtr = plBytePtr + 1&
    End Select
  Else
    Dim alDims()    As Long
    Dim iDimCt      As Integer
    Dim i           As Long
    Dim j           As Long
    ' number of dims:
    iDimCt = PeekInteger(pabSource(), plBytePtr)
    ' -  the dims
    ReDim Preserve alDims(1 To (iDimCt * 2))
    For i = 1& To iDimCt * 2&
      alDims(i) = PeekLong(pabSource(), plBytePtr)
    Next i
    'Now we loop the whole array and we'll recurse
    If iDimCt > 1& Then
      ReDim pvVariant(alDims(1) To alDims(2), alDims(3) To alDims(4))
    Else
      ReDim pvVariant(alDims(1) To alDims(2))
    End If
    For i = CLng(alDims(1)) To CLng(alDims(2))
      If iDimCt > 1& Then
        For j = CLng(alDims(3)) To CLng(alDims(4))
          GetVariant pvVariant(i, j), plBytePtr, pabSource()
        Next j
      Else
        GetVariant pvVariant(i), plBytePtr, pabSource()
      End If
    Next i
  End If
End Sub

Public Sub PokeByte(ByVal pbByte As Byte, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  pabDest(plBytePtr) = pbByte
  plBytePtr = plBytePtr + klSizeOfByte
End Sub

Public Sub PokeInteger(ByVal piInt As Integer, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  CopyMemory pabDest(plBytePtr), piInt, klSizeOfInt
  plBytePtr = plBytePtr + klSizeOfInt
End Sub

Public Sub PokeLong(ByVal plLong As Long, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  CopyMemory pabDest(plBytePtr), plLong, klSizeOfLong
  plBytePtr = plBytePtr + klSizeOfLong
End Sub

Public Sub PokeSingle(ByVal psngValue As Single, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  CopyMemory pabDest(plBytePtr), psngValue, LenB(psngValue)
  plBytePtr = plBytePtr + LenB(psngValue)
End Sub

Public Sub PokeDouble(ByVal pdblValue As Double, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  CopyMemory pabDest(plBytePtr), pdblValue, LenB(pdblValue)
  plBytePtr = plBytePtr + LenB(pdblValue)
End Sub

Public Sub PokeCurrency(ByVal pcurValue As Currency, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  CopyMemory pabDest(plBytePtr), pcurValue, LenB(pcurValue)
  plBytePtr = plBytePtr + LenB(pcurValue)
End Sub

Public Sub PokeDate(ByVal pdtValue As Date, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  CopyMemory pabDest(plBytePtr), pdtValue, LenB(pdtValue)
  plBytePtr = plBytePtr + LenB(pdtValue)
End Sub

Public Sub PokeBoolean(ByVal pfValue As Boolean, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  CopyMemory pabDest(plBytePtr), pfValue, LenB(pfValue)
  plBytePtr = plBytePtr + LenB(pfValue)
End Sub

Public Sub PokeString(ByRef psString As String, ByRef pabDest() As Byte, ByRef plBytePtr As Long)
  Dim lstrlen   As Long
  lstrlen = Len(psString)
  CopyMemory pabDest(plBytePtr), lstrlen, klSizeOfLong
  plBytePtr = plBytePtr + klSizeOfLong
  If lstrlen Then
    CopyMemoryFromString pabDest(plBytePtr), psString, lstrlen
    plBytePtr = plBytePtr + lstrlen
  End If
End Sub

Public Sub MoveVariant(ByRef pabDest() As Byte, ByRef plBytePtr As Long, ByRef pvVariant As Variant)
  Dim iVarType  As Integer
  
  iVarType = VarType(pvVariant)
  'save vartype
  If plBytePtr < LBound(pabDest) Then 'fool guard (for those who confuse base index)
    Err.Raise MAKE_VBERROR(9&), "MMemory::MoveVariant", "Start byte smaller than array base"
  End If
  CopyMemory pabDest(plBytePtr), iVarType, klSizeOfInt
  plBytePtr = plBytePtr + klSizeOfInt
  If (iVarType And vbArray) = 0 Then
    Select Case iVarType
    Case vbEmpty      '0    Empty (uninitialized)
    Case vbNull       '1    Null (no valid data)
    Case vbInteger    '2    Integer
      PokeInteger pvVariant, pabDest(), plBytePtr
    Case vbLong       '3    Long integer
      PokeLong pvVariant, pabDest(), plBytePtr
    Case vbSingle     '4    Single-precision floating-point number
      PokeSingle pvVariant, pabDest(), plBytePtr
    Case vbDouble     '5    Double-precision floating-point number
      PokeDouble pvVariant, pabDest(), plBytePtr
    Case vbCurrency   '6    Currency value
      PokeCurrency pvVariant, pabDest(), plBytePtr
    Case vbDate       '7    Date value
      PokeDate pvVariant, pabDest(), plBytePtr
    Case vbString     '8    String
      PokeString (pvVariant), pabDest(), plBytePtr
    Case vbBoolean    '11   Boolean value
      PokeBoolean pvVariant, pabDest(), plBytePtr
    Case vbDecimal    '14   Decimal value
      'Unsupported, borrow VB error 13&
      Err.Raise MAKE_VBERROR(13&), "MMemory::MoveVariant", "Variant Decimal subtype unsupported"
    Case vbByte       '17   Byte value
      pabDest(plBytePtr) = pvVariant
      plBytePtr = plBytePtr + 1&
    End Select
  Else
    'We have a variant containing an array. We'll handle single and bidimensional arrays only
    'Note that each element of the array can also be an array (in a variant)
    'and that we must later reload the variant by making an array in a variant.
    'Note also, that we rely on the assertion that each element of an array contained
    'in a variant, is a variant.
    Dim avDims      As Variant
    Dim iDimCt      As Integer
    Dim i           As Long
    Dim j           As Long
    GetVarArrayBounds pvVariant, iDimCt, avDims
    'we copy the infos needed to rebuild the array, ie:
    ' - the number of dims:
    PokeInteger iDimCt, pabDest(), plBytePtr
    ' -  the dims
    For i = 1& To iDimCt * 2&
      PokeLong avDims(i - 1&), pabDest(), plBytePtr
    Next i
    'Now we loop the whole array and we'll recurse
    For i = CLng(avDims(0)) To CLng(avDims(1))
      If iDimCt > 1& Then
        For j = CLng(avDims(2)) To CLng(avDims(3))
          MoveVariant pabDest(), plBytePtr, pvVariant(i, j)
        Next j
      Else
        MoveVariant pabDest(), plBytePtr, pvVariant(i)
      End If
    Next i
  End If
End Sub

Public Function CalcByteSize(ByRef pvVariant As Variant) As Long
  Dim sng       As Single
  Dim dbl       As Double
  Dim cur       As Currency
  Dim dt        As Date
  Dim bool      As Boolean
  Dim iInt      As Integer
  Dim lng       As Long
  Dim str       As String
  Dim iVarType  As Integer
  
  Dim lRetSize  As Long
  
  iVarType = VarType(pvVariant)
  'save vartype
  lRetSize = lRetSize + klSizeOfInt
  If (iVarType And vbArray) = 0 Then
    Select Case iVarType
    Case vbEmpty      '0    Empty (uninitialized)
    Case vbNull       '1    Null (no valid data)
    Case vbInteger    '2    Integer
      lRetSize = lRetSize + klSizeOfInt
    Case vbLong       '3    Long integer
      lRetSize = lRetSize + klSizeOfLong
    Case vbSingle     '4    Single-precision floating-point number
      lRetSize = lRetSize + LenB(sng)
    Case vbDouble     '5    Double-precision floating-point number
      lRetSize = lRetSize + LenB(dbl)
    Case vbCurrency   '6    Currency value
      lRetSize = lRetSize + LenB(cur)
    Case vbDate       '7    Date value
      lRetSize = lRetSize + LenB(dt)
    Case vbString     '8    String
      lRetSize = lRetSize + klSizeOfLong + Len(pvVariant)
    Case vbBoolean    '11   Boolean value
      lRetSize = lRetSize + LenB(bool)
    Case vbDecimal    '14   Decimal value
      'Unsupported, borrow VB error 13&
      Err.Raise MAKE_VBERROR(13&), "MMemory::CalcByteSize", "Variant Decimal subtype unsupported"
    Case vbByte       '17   Byte value
      lRetSize = lRetSize + 1&
    End Select
  Else
    'We have a variant containing an array. We'll handle single and bidimensional arrays only
    'Note that each element of the array can also be an array (in a variant)
    'and that we must later reload the variant by making an array in a variant.
    'Note also, that we rely on the assertion that each element of an array contained
    'in a variant, is a variant.
    Dim avDims      As Variant
    Dim iDimCt      As Integer
    Dim i           As Long
    Dim j           As Long
    GetVarArrayBounds pvVariant, iDimCt, avDims
    'we copy the infos needed to rebuild the array, ie:
    ' - the number of dims:
    lRetSize = lRetSize + klSizeOfInt
    ' -  the dims
    For i = 1& To iDimCt * 2&
      lRetSize = lRetSize + klSizeOfLong
    Next i
    'Now we loop the whole array and we'll recurse
    For i = CLng(avDims(0)) To CLng(avDims(1))
      If iDimCt > 1& Then
        For j = CLng(avDims(2)) To CLng(avDims(3))
          lRetSize = lRetSize + CalcByteSize(pvVariant(i, j))
        Next j
      Else
        lRetSize = lRetSize + CalcByteSize(pvVariant(i))
      End If
    Next i
  End If
  CalcByteSize = lRetSize
End Function

Public Sub GetVarArrayBounds(ByRef pvVar As Variant, ByRef piRetDims As Integer, ByRef pavRetBounds As Variant)
  Dim lLB1    As Long
  Dim lUB1    As Long
  Dim lLB2    As Long
  Dim lUB2    As Long
  
  On Error Resume Next
  piRetDims = 0
  lLB1 = LBound(pvVar)
  If Err.Number = 0 Then
    piRetDims = 1
    lUB1 = UBound(pvVar)
    lLB2 = LBound(pvVar, 2)
    If Err.Number = 0& Then
      piRetDims = 2
      lUB2 = UBound(pvVar, 2)
    End If
  End If
  pavRetBounds = Array(lLB1, lUB1, lLB2, lUB2)
End Sub

