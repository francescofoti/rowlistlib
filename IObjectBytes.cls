VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IObjectBytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'(C) 2007-2014, Developpement Informatique Service, Francesco Foti
'          internet: http://www.devinfo.net
'          email:    info@devinfo.ch
'
'IObjectBytes.bas class interface definition module
'This interface manages object serialization as a stream of bytes;
'and as we are in VB, the stream of byte is actually an array.
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

Public Property Get ByteSize() As Long

End Property

Public Sub GetDataBytes(ByRef abRetObject() As Byte)

End Sub

Public Sub SetDataBytes(ByRef abObjectData() As Byte, Optional ByRef plByte As Long = 1&)

End Sub


