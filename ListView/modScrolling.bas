Attribute VB_Name = "modScrolling"
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

' Retrieves different system paramenters
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

' Used to enable ScrollbarsConst
Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

' Used to save information regarding a certain window
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

' Used to replace the wndproc-entry, as well as calling another window proc (the original)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Used to set/retrieve the current scroll state
Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, ByRef lpScrollInfo As ScrollInfo) As Long
Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, ByRef lpcScrollInfo As ScrollInfo, ByVal bool As Boolean) As Long

' The current scroll state
Type ScrollInfo
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

' Used in setting/retreving scroll state
Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

' Different scrollbar 'messages'
Enum ScrollBarActions
    SB_LINEUP = &H0
    SB_LINEDOWN = &H1
    SB_PAGEUP = &H2
    SB_PAGEDOWN = &H3
    SB_THUMBPOSITION = &H4
    SB_THUMBTRACK = &H5
    SB_TOP = &H6
    SB_BOTTOM = &H7
    SB_ENDSCROLL = &H8
End Enum

' The windproc-entry
Public Const GWL_WNDPROC = -4
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115

' Different window messages
Public Const WM_DESTROY As Long = &H2
Public Const WM_MOUSEWHEEL = &H20A

' Used to obtain the amount of lines to scroll with the wheel
Public Const SPI_GETWHEELSCROLLLINES = &H68

' Contains all the scroll-classes
Public ScrollClasses As New Collection

Public Function WndProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim oldProc As Long, ScrollInfo As ScrollInfo, Scrolling As clsScrolling

    ' Retrieve the different variables
    oldProc = GetProp(hWnd, "OldProc")
    
    ' Initialize the scrolling-class
    For Each Scrolling In ScrollClasses
        If Scrolling.ControlWnd = hWnd Then
            Exit For
        End If
    Next
    
    ' Set the size of the structure
    ScrollInfo.cbSize = LenB(ScrollInfo)
    ScrollInfo.fMask = SIF_ALL
     
     ' Process the messages
     Select Case iMsg
        Case WM_HSCROLL

            ' First, get the current scroll bar state
            GetScrollInfo hWnd, SB_HORZ, ScrollInfo
            
            ' Process the scrollbar message
            If ProcessMessage(LoWord(wParam), SB_HORZ, hWnd, ScrollInfo, Scrolling) Then
                SetScrollInfo hWnd, SB_HORZ, ScrollInfo, True
                ExecuteEvent SB_HORZ, Scrolling
            End If
            
         Case WM_VSCROLL

            ' First, get the current scroll bar state
            GetScrollInfo hWnd, SB_VERT, ScrollInfo
            
            ' Process the scrollbar message
            If ProcessMessage(LoWord(wParam), SB_VERT, hWnd, ScrollInfo, Scrolling) Then
                SetScrollInfo hWnd, SB_VERT, ScrollInfo, True
                ExecuteEvent SB_VERT, Scrolling
            End If
        
        ' The mouse wheel has been rotated
        Case WM_MOUSEWHEEL

            ' If the vertical scrollbar is visible, ...
            If Scrolling.VScroll.Visible Then
                ' ... scroll it!
                Scrolling.VScroll.Position = Scrolling.VScroll.Position _
                 - ((HiWord(wParam) / 120) * Scrolling.LinesWheelScroll)
            End If
            
        Case WM_DESTROY
     
            ' Indicates that we have to shut down the subclassing immediately
            If Not (Scrolling Is Nothing) Then
                Scrolling.FinishScrolling
            End If
        
     End Select

     ' Pass the message to the orignal window proc
     WndProc = CallWindowProc(oldProc, hWnd, iMsg, wParam, lParam)

End Function

Private Sub ExecuteEvent(ByVal Index As Long, Scrolling As clsScrolling)
    
    ' Inform that we've now changed the position of the scrollbar
    If Index = SB_HORZ Then
        If Scrolling.HScroll.InvokeEvent(1) Then
            Exit Sub
        End If
    Else
        If Scrolling.VScroll.InvokeEvent(1) Then
            Exit Sub
        End If
    End If

End Sub

Private Function ProcessMessage(ByVal Message As Long, ByVal Index As Long, ByVal hWnd As Long, ScrollInfo As ScrollInfo, Scrolling As clsScrolling) As Boolean

    ' Process the given message
    Select Case Message
        Case SB_PAGEUP
            ScrollInfo.nPos = ScrollInfo.nPos - ScrollInfo.nPage
        Case SB_PAGEDOWN
            ScrollInfo.nPos = ScrollInfo.nPos + ScrollInfo.nPage
        Case SB_LINEUP
            ScrollInfo.nPos = ScrollInfo.nPos - Scrolling.LinesSmallChange
        Case SB_LINEDOWN
            ScrollInfo.nPos = ScrollInfo.nPos + Scrolling.LinesSmallChange
        Case SB_THUMBTRACK
            
            ' Inform that we're now scrolling
            If Index = SB_HORZ Then
                Scrolling.HScroll.InvokeEvent 2, ScrollInfo.nTrackPos
            Else
                Scrolling.VScroll.InvokeEvent 2, ScrollInfo.nTrackPos
            End If
            
            ' No need to update anything
            Exit Function
                        
        Case SB_THUMBPOSITION
            
            ' This is the final thumb position
            ScrollInfo.nPos = ScrollInfo.nTrackPos
            
        Case Else
            
            ' Don't bother updating it
            Exit Function
            
    End Select

    ' Here we are making sure the range is correct
    If ScrollInfo.nPos < 0 Then
        ScrollInfo.nPos = 0
    End If

    ' As the above
    If ScrollInfo.nPos > ScrollInfo.nMax Then
        ScrollInfo.nPos = ScrollInfo.nMax
    End If
    
    ' Process this
    ProcessMessage = True

End Function

Private Function HiWord(ByVal lDWord As Long) As Integer

    ' Retrieves the higher-portion of a long
    HiWord = (lDWord And &HFFFF0000) \ &H10000
  
End Function

Private Function LoWord(ByVal lDWord As Long) As Integer

    ' Retrieves the lower-portion of a long
    If lDWord And &H8000& Then
        LoWord = lDWord Or &HFFFF0000
    Else
        LoWord = lDWord And &HFFFF&
    End If
    
End Function




