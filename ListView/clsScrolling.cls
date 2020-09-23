VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScrolling"
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

' The different scrollbars
Public VScroll As clsScrollbar
Attribute VScroll.VB_VarUserMemId = 0
Public HScroll As clsScrollbar

' The window to use
Public ControlWnd As Long

' The amount of lines to scroll with the wheel
Public LinesWheelScroll As Long

' How many lines to scroll with the ordinary arrow buttons
Public LinesSmallChange As Long

Public Property Get Subclassed() As Boolean

    ' See if the control is already subclassed
    Subclassed = (GetProp(ControlWnd, "OldProc") <> 0)

End Property

Public Sub EnableScrolling()

    ' Make sure we haven't already subclassed this form
    If Not Subclassed Then
    
        ' Add this class to the collection over the scroll-classes
        ScrollClasses.Add Me, CStr(ObjPtr(Me))
        
        ' Retrieve and save the original wndproc
        SetProp ControlWnd, "OldProc", GetWindowLong(ControlWnd, GWL_WNDPROC)

        ' Then, subclass the form
        SetWindowLong ControlWnd, GWL_WNDPROC, AddressOf WndProc
        
        ' Initialize the scrollbars
        Set HScroll = CreateScrollbar(SB_HORZ)
        Set VScroll = CreateScrollbar(SB_VERT)
        
    End If

End Sub

Private Function CreateScrollbar(ByVal Index As ScrollbarsConst) As clsScrollbar

    ' Create a new scrollbar
    Set CreateScrollbar = New clsScrollbar

    ' Set its properties
    With CreateScrollbar
        .Index = Index
        .hWnd = ControlWnd
    End With

End Function

' Called when one wish to stop scrolling
Public Sub FinishScrolling()

    Dim oldProc As Long
    
    ' Retreve the old proc
    oldProc = GetProp(ControlWnd, "OldProc")

    ' Hide scrollbars
    ShowScrollBar ControlWnd, SB_BOTH, False

    ' Make sure there's anything to remove
    If oldProc <> 0 Then
    
        ' Reset the wndproc
        SetWindowLong ControlWnd, GWL_WNDPROC, oldProc
        
        ' Remove the properties
        RemoveProp ControlWnd, "OldProc"
    
        ' Remove this class from the following collection
        ScrollClasses.Remove CStr(ObjPtr(Me))
    
    End If
    
    ' Clear scrollbars
    Set VScroll = Nothing
    Set HScroll = Nothing

End Sub

Private Sub Class_Initialize()

    ' Retireve the amount of lines to scroll with the wheel
    SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, VarPtr(LinesWheelScroll), 0
    
    ' Set the default value, if the found is invalid
    If (LinesWheelScroll <= 0) Then
       LinesWheelScroll = 3
    End If

    ' Set other values
    LinesSmallChange = 1

End Sub
