VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScrollbar"
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

' The different ScrollbarsConst
Enum ScrollbarsConst
    SB_HORZ = 0
    SB_VERT = 1
    SB_BOTH = 3
End Enum

' The different events
Event Scrolling(ByVal Position As Long)
Event Changed(Cancel As Boolean)

' Different information related to the retrival/setting of the correct scrollbar
Public hWnd As Long
Public Index As ScrollbarsConst

' Used to set/retrieve scrollbar info
Private ScrollInfo As ScrollInfo

' Saves the state of the scrollbar
Private pVisible As Boolean

Public Function InvokeEvent(ByVal EventID As Long, Optional Paramenter As Long) As Boolean

    ' Execute the given event
    Select Case EventID
    Case 1: RaiseEvent Changed(InvokeEvent)
    Case 2: RaiseEvent Scrolling(Paramenter)
    End Select

End Function

Public Property Get Min() As Long
    
    ' Update the structure first
    UpdateInfo
    
    ' Then return the information
    Min = ScrollInfo.nMin

End Property

Public Property Let Min(ByVal vNewValue As Long)

    ' Set the new information
    UpdateInfo
    ScrollInfo.nMin = vNewValue
    
    ' Save the changes
    SetInfo

End Property

Public Property Get Max() As Long
    
    ' Update the structure first
    UpdateInfo
    
    ' Then return the information
    Max = ScrollInfo.nMax

End Property

Public Property Let Max(ByVal vNewValue As Long)

    ' Set the new information
    UpdateInfo
    ScrollInfo.nMax = vNewValue
    
    ' Save the changes
    SetInfo

End Property

Public Property Let Page(ByVal vNewValue As Long)

    ' Set the new information
    UpdateInfo
    ScrollInfo.nPage = vNewValue
    
    ' Save the changes
    SetInfo

End Property

Public Property Get Page() As Long
    
    ' Update the structure first
    UpdateInfo
    
    ' Then return the information
    Page = ScrollInfo.nPage

End Property

Public Property Let Position(ByVal vNewValue As Long)

    ' Set the new information
    UpdateInfo
    ScrollInfo.nPos = vNewValue
    
    ' Save the changes
    SetInfo

    ' Inform that the position has changed
    RaiseEvent Changed(True)

End Property

Public Property Get Position() As Long
    
    ' Update the structure first
    UpdateInfo
    
    ' Then return the information
    Position = ScrollInfo.nPos

End Property
    
Private Sub UpdateInfo()

    ' Initialize the structure
    With ScrollInfo
        .cbSize = LenB(ScrollInfo)
        .fMask = SIF_ALL
    End With

    ' Updates the scrollbar information
    GetScrollInfo hWnd, Index, ScrollInfo

End Sub

Private Sub SetInfo()

    ' Initialize the structure here too
    With ScrollInfo
        .cbSize = LenB(ScrollInfo)
        .fMask = SIF_ALL
    End With

    ' Set the given info
    SetScrollInfo hWnd, Index, ScrollInfo, True

End Sub

Public Property Get Visible() As Boolean

    ' Return the current visible state (note that this may get out of sync)
    Visible = pVisible

End Property

Public Property Let Visible(ByVal vNewValue As Boolean)

    ' Set whether or not the scrollbar is visible
    ShowScrollBar hWnd, Index, vNewValue

    ' Save the new state
    pVisible = vNewValue

End Property
