VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStyle"
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

' This class is utilized to draw "styled" cells

' Different text-drawing flags
Enum TextFormat
    DT_WORDBREAK = &H10
    DT_VCENTER = &H4
    DT_TOP = &H0
    DT_TABSTOP = &H80
    DT_SINGLELINE = &H20
    DT_RIGHT = &H2
    DT_NOCLIP = &H100
    DT_LEFT = &H0
    DT_EXPANDTABS = &H40
    DT_INTERNAL = &H1000
    DT_EXTERNALLEADING = &H200
    DT_CENTER = &H1
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
End Enum

Enum BorderStyleTypes
    Border_Solid = 0
    Border_Dash = 1
    Border_Dot = 2
    Border_DashDot = 3
    Border_DashDotDot = 4
    Border_Transparent = 5
    Border_InsideSolid = 6
End Enum

' The color of the background square
Public BackColor As Long

' The pictures to draw
Public Pictures As New Collection

' The color of the text
Public ForeColor As Long

' Internal spacing
Public Padding As New clsRect

' External spacing
Public Margin As New clsRect

' Border-colors (should contain a collection of rects)
Public Borders As New Collection

' The style and width of the borders
Public BorderStyle As BorderStyleTypes
Public BorderWidth As Long

' Text-specific flags
Public Format As TextFormat

' XP-theme background
Public Theme As New clsTheme

' The current font
Public Font As New StdFont

' The main drawing procedure
Public Sub DrawCell(refSurface As Object, Dimensions As clsRect, Text As String, Images As Collection)

    Dim TextArea As New clsRect, Tell As Long, Picture As clsPicture
    
    ' Initialize the area to draw the text
    With TextArea
        .Left = Dimensions.Left + Padding.Left
        .Top = Dimensions.Top + Padding.Top
        .Right = Dimensions.Right - Padding.Right
        .Bottom = Dimensions.Bottom - Padding.Bottom
    End With

    ' Draw the backcolor
    refSurface.Line (Dimensions.Left, Dimensions.Top)-(Dimensions.Right, Dimensions.Bottom), BackColor, BF
    
    ' Draw the XP-style
    Theme.DrawTheme refSurface.hWnd, refSurface.hDC, Dimensions
    
    ' Then draw all the pictures in the style, ...
    For Each Picture In Pictures
        Picture.DrawPicture refSurface, Dimensions.Left, Dimensions.Top
    Next
    
    ' ... as well as the given images.
    For Each Picture In Images
        Picture.DrawPicture refSurface, Dimensions.Left, Dimensions.Top
    Next
    
    ' Set the border-styles
    refSurface.DrawStyle = BorderStyle
    refSurface.DrawWidth = BorderWidth
    
    ' Draw the four borders
    For Tell = 0 To Borders.Count - 1
        DrawLine refSurface, Dimensions.Left + Tell, Dimensions.Top + Tell, Dimensions.Left + Tell, Dimensions.Bottom - Tell, Borders(Tell + 1).Left
        DrawLine refSurface, Dimensions.Left + Tell, Dimensions.Bottom - Tell, Dimensions.Right - Tell, Dimensions.Bottom - Tell, Borders(Tell + 1).Bottom
        DrawLine refSurface, Dimensions.Right - Tell, Dimensions.Bottom - Tell, Dimensions.Right - Tell, Dimensions.Top + Tell, Borders(Tell + 1).Right
        DrawLine refSurface, Dimensions.Right - Tell, Dimensions.Top + Tell, Dimensions.Left + Tell, Dimensions.Top + Tell, Borders(Tell + 1).Top
    Next
    
    ' Reset the border style
    refSurface.DrawStyle = 0
    refSurface.DrawWidth = 1
    
    ' Swap the font itself
    Swap refSurface.Font, Font
    
    ' Use the current text color
    refSurface.ForeColor = ForeColor
    
    ' Draw the given text
    DrawText refSurface.hDC, Text, Len(Text), TextArea.UDT, Format
    
    ' Clean up
    Swap refSurface.Font, Font

End Sub

Private Sub DrawLine(refSurface As PictureBox, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Color As Long)
    
    ' Make sure the color is valid
    If Color <> -1 Then
        refSurface.Line (X1, Y1)-(X2, Y2), Color
    End If

End Sub

Private Sub Swap(Var1 As Object, Var2 As Object)

    Dim Temp As Object
    
    ' Save the first object
    Set Temp = Var1
    
    ' Swap the two variables
    Set Var1 = Var2
    Set Var2 = Temp
    
End Sub

Public Function Clone() As clsStyle

    Dim Border As clsRect

    ' Firstly, create a new class
    Set Clone = New clsStyle
    
    ' Then, copy all the data
    With Clone
        .BackColor = Me.BackColor
        .ForeColor = Me.ForeColor
        .Font.Bold = Me.Font.Bold
        .Font.Charset = Me.Font.Charset
        .Font.Italic = Me.Font.Italic
        .Font.Name = Me.Font.Name
        .Font.Size = Me.Font.Size
        .Font.Strikethrough = Me.Font.Strikethrough
        .Font.Underline = Me.Font.Underline
        .Font.Weight = Me.Font.Weight
        .Format = Me.Format
        
        ' Copy rects
        Set .Margin = Me.Margin.Clone
        Set .Padding = Me.Padding.Clone
        
        ' Copy all the borders
        For Each Border In Me.Borders
            .Borders.Add Border.Clone
        Next
        
        ' Copy the theme
        Set .Theme = Theme.Clone
    
    End With

End Function

Private Sub Class_Initialize()

    ' Initialize the font
    Font.Name = "MS Sans Serif"
    Font.Size = 8
    
    ' Initialize border-width
    BorderWidth = 1
    
End Sub
