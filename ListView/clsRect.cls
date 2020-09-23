VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Copyright (C) 2006 Kristian S. Stangeland

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

' Contains the dimensions of a rectangle in a coordinate system
Public Left As Long     ' X1
Public Top As Long      ' Y1
Public Right As Long    ' X2
Public Bottom As Long   ' Y2

Friend Property Get UDT() As RECT

    ' Return the data within this class as a UDT
    With UDT
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With

End Property

Friend Property Let UDT(vNewValue As RECT)

    ' Sets the content of this class to correspond to the given UDF
    With vNewValue
        Left = .Left
        Top = .Top
        Right = .Right
        Bottom = .Bottom
    End With

End Property

Friend Function Clone() As clsRect

    ' Create a new class
    Set Clone = New clsRect
    
    ' Then, copy the content
    With Clone
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With

End Function
