Attribute VB_Name = "MFunctions"
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'MFunctions.bas module
'This module contains general purpose functions, constants and vairable
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

' 5001 - 5010: Shared between object implementing IObjectBytes
Public Const kErrBadClassIDBytes      As Long = 5001& '%1%=class name
Public Const kErrBadClassVerBytes     As Long = 5002& '%1%=class name

Public Const klObjectErrBase          As Long = 6000&   'Which leaves us from 6000& to 29000& for our errors

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
Public Function MAKE_OBJECTERROR(ByVal plErrCode As Long, ByVal plModuleErrBase As Long) As Long
  MAKE_OBJECTERROR = vbObjectError Or (plErrCode + klObjectErrBase + plModuleErrBase)
End Function


