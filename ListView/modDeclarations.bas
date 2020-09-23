Attribute VB_Name = "modDeclarations"
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

' Draws text within a rectangle
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

' Used to swap memory
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

' Basicly used to retrieve what window the mouse is currently over
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' API-calls needed for drawing XP style
Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal hDC As Long, prc As RECT) As Long
Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long

' Similar to the clsRect-class, but as a UDT
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Used in retrieving the mouse position
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Function CreateRect(Left, Top, Right, Bottom) As clsRect

    ' Firstly, create a new class
    Set CreateRect = New clsRect
    
    ' Set the data
    With CreateRect
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With

End Function

Public Sub SwapMemory(ByVal lpDestination As Long, ByVal lpSource As Long, ByVal Size As Long)

    Dim bTemp() As Byte
    
    ' Reallocate the temporary array
    ReDim bTemp(Size - 1)
    
    ' Copy the content of the first memory into the temp array
    CopyMemory bTemp(0), ByVal lpDestination, Size
    
    ' Then swap the memory locations
    CopyMemory ByVal lpDestination, ByVal lpSource, Size
    CopyMemory ByVal lpSource, bTemp(0), Size

End Sub
