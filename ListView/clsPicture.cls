VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicture"
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

' Location to draw this image
Public X As Single
Public Y As Single

' The size of the resulting picture
Public Width As Long
Public Height As Long

' Where to start paining from in the source picture (clipping)
Public Clipping As New clsRect

' The opcode to use
Public Opcode As RasterOpConstants

' The image to use
Private pPicture As StdPicture

Public Sub DrawPicture(refSurface As Object, addX As Long, addY As Long)

    ' Start drawing the picture
    refSurface.PaintPicture Picture, addX + X, addY + Y, Width, Height, Clipping.Left, _
     Clipping.Top, Clipping.Right - Clipping.Left, Clipping.Bottom - Clipping.Top, Opcode

End Sub

Public Property Get Picture() As StdPicture

    ' Return the picture-object
    Set Picture = pPicture

End Property

Public Property Let Picture(ByVal vNewValue As StdPicture)

    ' Set the picture to use
    Set pPicture = vNewValue
    
    ' Resize the width and height
    Width = ConvertToPixels(pPicture.Width)
    Height = ConvertToPixels(pPicture.Height)
    Clipping.Right = Width
    Clipping.Bottom = Height

End Property

Private Function ConvertToPixels(ByVal Himetric As Double) As Single
    
    ' Convert himetric to pixels
    ConvertToPixels = CLng((Himetric / 2540 * 1440) / Screen.TwipsPerPixelX)

End Function

Private Sub Class_Initialize()

    ' Set default rasteropcode
    Opcode = vbSrcCopy

End Sub
