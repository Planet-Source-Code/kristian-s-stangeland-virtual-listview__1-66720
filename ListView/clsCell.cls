VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCell"
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

' The different available controls
Public Enum EditControls
    Edit_TextBox
    Edit_ComboBox
End Enum

' The type of edit-box to utilize
Public EditControl As EditControls

' Used in conjunction with the combobox control
Public EditData As Variant

' The text to draw
Public Text As String

' Different pictures to draw
Public Pictures As New Collection

' Different styles
Public StyleNormal As clsStyle
Public StyleSelected As clsStyle
Public StylePressed As clsStyle
Public StyleHover As clsStyle

' Whether or not the cell is currently being hovered or pressed
Public Hovering As Boolean
Public Pressing As Boolean

' The width and height (ignored when the cell is not a column header)
Public Width As Long
Public Height As Long

' Visible locations (where the cell is on the screen)
Public X As Long
Public Y As Long

' Used to relocate the object relative to the above location
Public OffsetX As Long
Public OffsetY As Long

' Extra data free to be used by the program (its preservation cannot be
' promised in non-column headers)
Public Tag As Variant

' Information regarding the positioning of this cell
Public Column As Long
Public Row As Long ' (ignored in the case of column headers, obviously)

' Whether or not to draw this cell (or column)
Public Visible As Boolean
