VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

' This class is used to draw XP theme

' What class to draw
Public Class As String

' What type of the class to draw
Public PartID As Long
Public StateID As Long

Public Function IsOK(ByVal hWnd As Long, sClass As String) As Boolean
    
    On Error Resume Next
    Dim hTheme As Long
    
    ' Open the given class
    hTheme = OpenThemeData(hWnd, StrPtr(sClass))

    ' Make sure we managed to load the theme
    If hTheme <> 0 Then
        
        ' We managed to create the class - return true
        IsOK = True
        
        ' Clean up
        CloseThemeData hTheme
    
    End If

End Function

Public Function Clone() As clsTheme

    ' Create a new class
    Set Clone = New clsTheme
    
    ' Copy the content
    With Clone
        .Class = Class
        .PartID = PartID
        .StateID = StateID
    End With

End Function

Public Sub DrawTheme(ByVal hWnd As Long, ByVal hDC As Long, Dimensions As clsRect)

    Dim hTheme As Long, rRect As RECT
    
    ' Make sure the class is initialized
    If LenB(Class) = 0 Then
        Exit Sub
    End If
    
    ' Open the given class
    hTheme = OpenThemeData(hWnd, StrPtr(Class))
    
    ' Retrieve the UDT
    LSet rRect = Dimensions.UDT
    
    ' Make sure we managed to load the theme
    If (hTheme) Then
        
        ' Draw all
        DrawThemeParentBackground hWnd, hDC, rRect
        DrawThemeBackground hTheme, hDC, PartID, StateID, rRect, rRect
        
        ' Clean up
        CloseThemeData hTheme
    
    End If

End Sub
