VERSION 5.00
Begin VB.UserControl usrListView 
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   ScaleHeight     =   6405
   ScaleWidth      =   8805
   ToolboxBitmap   =   "usrListView.ctx":0000
   Begin VB.Timer tmrHover 
      Interval        =   25
      Left            =   8160
      Top             =   5760
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   459
      TabIndex        =   0
      Top             =   120
      Width           =   6885
      Begin VB.PictureBox picContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
         Begin VB.ComboBox cmbChange 
            Height          =   315
            Left            =   -30
            TabIndex        =   3
            Top             =   -30
            Width           =   1395
         End
      End
      Begin VB.TextBox txtChange 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Line lineMove 
         BorderWidth     =   2
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   184
         X2              =   184
         Y1              =   -8
         Y2              =   240
      End
   End
End
Attribute VB_Name = "usrListView"
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

' The two different types of borders
Public Enum BorderStyleForm
    Border_None = 0
    Border_FixedSingle = 1
End Enum

' Called when the list view require updated data of a virtual cell
Public Event RetrieveListItem(Item As clsCell)
Attribute RetrieveListItem.VB_Description = "Raised when the control require information regarding a certain item."

' Called when the user clicks an item
Event ItemClick(Item As clsCell, ClickData As clsMouse, Cancel As Boolean)
Attribute ItemClick.VB_Description = "Raised when an item (a cell) has been clicked."
Event ItemReleased(Column As clsCell, ClickData As clsMouse)
Attribute ItemReleased.VB_Description = "Raised when the user is finished clicking on an item."
Event ItemChanged(Item As clsCell)
Attribute ItemChanged.VB_Description = "Raised when an item's content has been changed through EditCell."
Event ColumnClick(Column As clsCell, ClickData As clsMouse, Cancel As Boolean)
Event ColumnReleased(Column As clsCell, ClickData As clsMouse)

' Column-sizing events
Event ColumnResized(Column As clsCell, Cancel As Boolean)
Attribute ColumnResized.VB_Description = "Raised when a column is going to be resized."

' Different standard events
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Raised when the user has clicked the control."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Raised when the user has moved within the control."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Raised when the user is finishing clicking the control."
Event KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
Attribute KeyDown.VB_Description = "Raised when a key has been pressed. If cancel is set to True, moving a selection may be canceled."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Raised when a key is pressed."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Raised when a key is released."

' Different drawing-events
Event DrawingBegun(X As Long, Y As Long)

' Represents the style of which to draw each cell
Public CellStyle As New clsStyle
Attribute CellStyle.VB_VarDescription = "Sets/returns the default cell style class."

' Represents the styles to draw the heading cells with
Public ColumnStyle As New clsStyle
Attribute ColumnStyle.VB_VarDescription = "Returns/sets the default column style class."
Public ColumnHover As clsStyle
Public ColumnPress As clsStyle

' Represents the style of an selected item
Public SelectedStyle As New clsStyle
Attribute SelectedStyle.VB_VarDescription = "Returns/sets the default style of an selected cell."

' The height of each row
Public RowsHeight As Long
Attribute RowsHeight.VB_VarDescription = "Returns/sets the default height of an row."

' The height of the columns
Public ColumnHeight As Long

' Internal padding
Public Padding As New clsRect
Attribute Padding.VB_VarDescription = "Returns/sets the amount of padding to use within this control."

' Contains all the different columns (within a collection)
Public Columns As New Collection
Attribute Columns.VB_VarDescription = "Returns (do not append columns directly) the appended columns."

' Used to enable scrolling
Public Scrolling As New clsScrolling
Attribute Scrolling.VB_VarDescription = "Class used to access the scrolling capabilities of this control."

' Whether or not to ignore updates
Public IgnoreUpdates As Boolean
Attribute IgnoreUpdates.VB_VarDescription = "Whether or not to ignore paint operations."

' The current selected item
Public SelectedItem As clsCell
Attribute SelectedItem.VB_VarDescription = "Returns the current selected item."

' Whether or not to hide columns
Public HideColumns As Boolean
Attribute HideColumns.VB_VarDescription = "Whether or not to hide the columns."

' Used to access databases by SQL
Public SQL As New clsSQL

' The amount of rows
Private pRowsAmount As Long

' The different scrollbars (used to catch events)
Private WithEvents HScroll As clsScrollbar
Attribute HScroll.VB_VarHelpID = -1
Private WithEvents VScroll As clsScrollbar
Attribute VScroll.VB_VarHelpID = -1

' Local cache storage of the visible cells
Private Cells() As clsCell

' The last class we were hovering
Private CellHovering As clsCell

' The selected column in "resizing"
Private SelColumn As clsCell

Public Property Get RowsAmount() As Long

    ' Return the amount of virtual rows
    RowsAmount = pRowsAmount

End Property

Public Property Let RowsAmount(ByVal vNewValue As Long)

    ' Firstly, change the value of the private corresponding variable
    pRowsAmount = vNewValue
    
    ' Then, automatically redraw the control
    RedrawAll

End Property

Public Property Get VisibleCells(ColumnIndex As Long, VisibleRowIndex As Long) As clsCell
Attribute VisibleCells.VB_Description = "Returns a given visible cell."

    ' Firstly, see if the given indexes are valid
    If ColumnIndex > 0 And ColumnIndex <= Columns.Count Then
        If VisibleRowIndex > 0 And VisibleRowIndex <= RowsPerPage Then
    
            ' Return the wanted cell
            Set VisibleCells = Cells(ColumnIndex, VisibleRowIndex)
    
        End If
    End If

End Property

' Let the parent form insert new data into all the visible cells
' (used for instance when the data has been sorted, ect.)
Public Sub UpdateContent()
Attribute UpdateContent.VB_Description = "When called, the control issues refresh events on all visible cells."

    Dim X As Long, Y As Long, iY As Long
    
    ' Retrieve the vertical Y-position
    If Not (VScroll Is Nothing) Then
        iY = VScroll.Position
    End If
    
    ' Go through all the cells
    For Y = LBound(Cells, 2) To UBound(Cells, 2)
        For X = LBound(Cells, 1) To UBound(Cells, 1)
        
            ' Make sure the cell is valid
            If Y + iY <= RowsAmount Then
                ' Let the parent form update the cell
                RaiseEvent RetrieveListItem(Cells(X, Y))
            End If
    
        Next
    Next
    
    ' Redraw the changes
    RedrawAll True

End Sub

Public Sub RedrawAll(Optional bIgnoreScrolls As Boolean, Optional overrideX As Long, Optional overrideY As Long)
Attribute RedrawAll.VB_Description = "Refreshes and redraws the content of this control."

    Dim iX As Long, iY As Long, X As Long, Y As Long, Column As clsCell, bSelectedDrawn As Boolean
    Dim oMargin As clsRect, oCell As clsCell, pCells() As clsCell, RowY As Long, dRect As New clsRect

    ' Make sure there's anything to draw
    If Columns.Count < 1 Then
        Exit Sub
    End If

    ' Firstly, make sure the padding is correctly initialized
    FixPadding
    
    ' Ignore all redrawing updates if set to do just that
    If IgnoreUpdates Then
        Exit Sub
    End If

    ' Remove the image content of the picturebox
    picMain.Cls
    
    ' Set the default positions (0 means no override)
    iX = overrideX
    iY = overrideY
    
    ' Initialize variables
    UpdateVariables iX, iY, pCells
    
    ' Reinitialize the scrolls if not otherwise noted
    If Not bIgnoreScrolls Then
        InitializeScrolls
    End If
    
    ' Swap the newly created cells with the old
    SwapMemory VarPtrArray(Cells), VarPtrArray(pCells), 4

    ' Set the starting position
    X = -iX
    
    ' Inform that we're about to draw the list view (this is the time to draw a background, ect.)
    RaiseEvent DrawingBegun(iX, iY)
    
    ' Draw the columns
    For Each Column In Columns

        ' Retrieve the column-margin
        Set oMargin = Column.StyleNormal.Margin
        
        ' Append the left margin
        X = X + oMargin.Left
        
        ' Calculate the Y-position
        Y = oMargin.Top
        
        ' See if we should draw the columns
        If Not HideColumns Then
            
            ' Save the current location of the column
            SetLocation Column, X, Y
            
            ' Use the style-draw function to draw the cell
            DrawCell Column
            
            ' Increase the Y-position
            Y = Y + Column.Height + oMargin.Bottom

        End If
        
        ' Draw all the rows within this column
        For RowY = LBound(Cells, 2) To UBound(Cells, 2)
        
            ' Retrieve the cell
            Set oCell = Cells(Column.Column, RowY)
        
            ' Include the margin
            Y = Y + oCell.StyleNormal.Margin.Top

            ' Save the location of the cel
            SetLocation oCell, X, Y
            
            ' Save the height and width of the cell
            oCell.Width = Column.Width
            oCell.Height = RowsHeight
            
            ' Draw this cell
            If Not (SelectedItem Is Nothing) Then
                
                ' See if this cell is the selected cell
                If SelectedItem.Row = oCell.Row And SelectedItem.Column = oCell.Column Then
                
                    ' Draw the selected cell
                    DrawCell oCell, oCell.StyleSelected
        
                    ' Save it for later use
                    Set SelectedItem = oCell
                    
                    ' We have now drawn a selected cell
                    bSelectedDrawn = True
                
                Else
                    DrawCell oCell
                End If
            
            Else
                DrawCell oCell
            End If
             
            ' Increase the Y-position
            Y = Y + oCell.Height + oCell.StyleNormal.Margin.Bottom
            
        Next
        
        ' Then append the width and the right margin
        X = X + Column.Width + oMargin.Right
        
    Next
    
    ' See if we need to draw another column-element
    If (RowsWidth < picMain.ScaleWidth) And (Not HideColumns) Then
    
        With dRect
            .Top = ColumnStyle.Margin.Top
            .Left = RowsWidth
            .Bottom = .Top + ColumnHeight
            .Right = picMain.ScaleWidth
        End With
    
        ' Draw the last column-element
        ColumnStyle.DrawCell picMain, dRect, "", ColumnStyle.Pictures
    
    End If
    
    ' If no selected cells has been drawn, ...
    If Not bSelectedDrawn Then
        ' ... make it so the selected item is outside the window
        If Not (SelectedItem Is Nothing) Then
            With SelectedItem
                .X = -UserControl.ScaleWidth
                .Y = -UserControl.ScaleHeight
            End With
        End If
    End If
    
    ' Show the redrawed picturebox (as we're using autoredraw)
    picMain.Refresh

End Sub

Public Function SelectItem(ByVal Column As Long, ByVal Row As Long) As Boolean
Attribute SelectItem.VB_Description = "Used to select an given item."

    Dim lMin As Long, lMax As Long, oldIgnore As Boolean, Item As New clsCell
    
    ' First, make sure the row and column is valid
    If Column < 1 Or Column > Columns.Count Or Row < 1 Or Row > pRowsAmount Then
        Exit Function
    End If
    
    ' Calculate the min- and max positions of the row
    lMin = VScroll.Position
    lMax = lMin + RowsPerPage

    ' Ignore updates
    oldIgnore = IgnoreUpdates
    IgnoreUpdates = True

    ' Thirdly, see if we need to scroll
    If Row <= lMin Then
        VScroll.Position = Row - 1
    ElseIf Row > lMax Then
        VScroll.Position = Row - RowsPerPage
    End If

    ' Stop ignoring updates
    IgnoreUpdates = oldIgnore
    
    ' Set the properties
    With Item
        .Row = Row
        .Column = Column
    End With
    
    ' Select the correct item
    Set SelectedItem = Item
    
    ' Update all
    RedrawAll True

End Function

Public Sub EditCell(refCell As clsCell)
Attribute EditCell.VB_Description = "Used to make a cell editable."

    Dim oPadding As clsRect, aItem

    ' Firstly, hide all other controls
    HideControls
    
    ' Retrieve the padding
    Set oPadding = refCell.StyleNormal.Padding

    ' Select this cell automatically
    Set SelectedItem = refCell

    ' Then initialize the control to use
    Select Case refCell.EditControl
        Case Edit_TextBox
    
            ' Show the control to use
            With txtChange
                .Left = refCell.X + refCell.OffsetX + oPadding.Left
                .Top = refCell.Y + refCell.OffsetY + oPadding.Top
                .Width = refCell.Width - oPadding.Left - oPadding.Right
                .Height = refCell.Height - oPadding.Top - oPadding.Bottom
                .Text = refCell.Text
                .BackColor = refCell.StyleNormal.BackColor
                .Visible = True
            End With
            
            ' Select all
            txtChange.SetFocus
            txtChange.SelStart = 0
            txtChange.SelLength = Len(refCell.Text)
        
        Case Edit_ComboBox
        
            ' Show the controls to use
            With cmbChange
            
                ' Set the text
                .Clear
                .Text = refCell.Text
                .BackColor = refCell.StyleNormal.BackColor
                .Visible = True
                
                ' Add all the items in the given array
                For Each aItem In refCell.EditData
                    .AddItem aItem
                Next
                
            End With
            
            ' Show the container as well
            With picContainer
                .Left = refCell.X + refCell.OffsetX + oPadding.Left
                .Top = refCell.Y + refCell.OffsetY + oPadding.Top - 2
                .Width = refCell.Width - oPadding.Left - oPadding.Right + 4
                .Height = refCell.Height - oPadding.Top - oPadding.Bottom + 6
                .Visible = True
            End With
        
    End Select

End Sub

Public Sub DrawCell(refCell As clsCell, Optional overrideStyle As clsStyle, Optional ByVal destSurface As Object)
Attribute DrawCell.VB_Description = "Procedure for drawing a given cell on a specific location."

    Dim oStyle As clsStyle, X As Long, Y As Long

    ' If the cell isn't visible, we'll ignore drawing it
    If Not refCell.Visible Then
        Exit Sub
    End If

    ' Calculate the X- and Y-position
    X = refCell.X + refCell.OffsetX
    Y = refCell.Y + refCell.OffsetY

    ' Only draw it if actually is visible within the control
    If Not (X + refCell.Width > 0 And Y + refCell.Height > 0 And _
     X < picMain.ScaleWidth And Y < picMain.ScaleHeight) Then
        Exit Sub
    End If
        
    ' See if we need to set the default drawing surface
    If destSurface Is Nothing Then
        Set destSurface = picMain
    End If
    
    ' Retrieve the style
    If overrideStyle Is Nothing Then
        ' Use the normal style unless we're currently pressing or hovering
        If refCell.Pressing Then
            Set oStyle = refCell.StylePressed
        ElseIf refCell.Hovering Then
            Set oStyle = refCell.StyleHover
        Else
            Set oStyle = refCell.StyleNormal
        End If
    Else
        ' Use the given style
        Set oStyle = overrideStyle
    End If

    ' Draw this cell
    oStyle.DrawCell destSurface, CreateRect(X, Y, X + refCell.Width, Y + refCell.Height), _
     refCell.Text, refCell.Pictures

End Sub

Private Sub HideControls(Optional sIgnore As String = "picMain")

    On Error Resume Next
    Dim Control As Control

    ' Hides all children
    For Each Control In UserControl.Controls
        If Control.Name <> sIgnore Then
            Control.Visible = False
        End If
    Next

End Sub

Private Sub SetLocation(refCell As clsCell, ByVal X As Long, ByVal Y As Long)

    ' Save the current location
    With refCell
        .X = X
        .Y = Y
    End With

End Sub

Public Function HitTest(X As Long, Y As Long) As clsCell
Attribute HitTest.VB_Description = "Returns a cell if a specific location contains one."

    Dim Column As clsCell, iY As Long, oCell As clsCell
    
    ' Firstly, check all the column
    For Each Column In Columns
    
        ' Make sure the X-position correspond
        If Column.X + Column.OffsetX <= X And Column.X + Column.OffsetX + Column.Width >= X Then
    
            ' See whether or not we've clicked it
            If Column.Y + Column.OffsetY <= Y And Column.Y + Column.OffsetY + Column.Height >= Y Then
            
                ' Return the cell
                Set HitTest = Column
                
                ' We're done
                Exit Function
            
            End If
        
            ' If not, test all the rows that belongs to this column
            For iY = LBound(Cells, 2) To UBound(Cells, 2)
                
                ' Retrieve the cell
                Set oCell = Cells(Column.Column, iY)
                
                ' See if we have collided with this cell
                If oCell.Y + oCell.OffsetY <= Y And oCell.Y + oCell.OffsetY + oCell.Height >= Y Then
                
                    ' Return the cell
                    Set HitTest = oCell
                    
                    ' We're done
                    Exit Function
                    
                End If
            
            Next
        
        End If
    
    Next

End Function

' Creates a new column
Public Function AppendColumn(Text As String, Optional Width As Long = 80, Optional ByVal Height As Long, Optional StyleNormal As clsStyle) As clsCell
Attribute AppendColumn.VB_Description = "Appends a column and redraws if necessary."

    ' Firstly, initialize the new class
    Set AppendColumn = New clsCell
    
    ' See if we need to use the default style
    If StyleNormal Is Nothing Then
        Set StyleNormal = ColumnStyle
    End If
    
    ' See if the height is invalid
    If Height <= 0 Then
        Height = ColumnHeight
    End If

    ' Then, set the default properties and the rest
    With AppendColumn
        .Column = Columns.Count + 1
        .Width = Width
        .Height = Height
        .Text = Text
        .Visible = True
        Set .StyleNormal = StyleNormal
        Set .StyleHover = ColumnHover
        Set .StylePressed = ColumnPress
    End With
    
    ' Append the column-class
    Columns.Add AppendColumn
    
    ' Obviously, this require an update
    RedrawAll

End Function

' Copies the content of the old cache to the new (whilst refreshing its values as well)
Private Sub UpdateVariables(iX As Long, iY As Long, pCells() As clsCell)

    Dim adjacentRow As Long, X As Long, Y As Long, bUpdate As Boolean
    
    ' Save the current X- and Y-position
    If Not (HScroll Is Nothing Or VScroll Is Nothing) Then
        
        ' Only retrieve the position of not otherwise set
        If iX = 0 Then
            iX = HScroll.Position
        End If
        
        ' As the above
        If iY = 0 Then
            iY = VScroll.Position
        End If
        
    End If
    
    ' Reinitialize the cells-array
    If pRowsAmount <= RowsPerPage Then
        ReDim pCells(1 To Columns.Count, 1 To pRowsAmount)
    Else
        ReDim pCells(1 To Columns.Count, 1 To RowsPerPage)
    End If

    ' Attempt to enter the old data into this new array (if it actually is initialized)
    If Not (Not Cells) Then ' A fast way of determine whether or not an array is initialized
        For Y = 1 To UBound(Cells, 2)
            
            ' Calculate what row (visible, that is) this cell belongs to
            adjacentRow = Cells(1, Y).Row - iY
            
            ' See if this row is within the range
            If adjacentRow > 0 And adjacentRow <= UBound(pCells, 2) Then
            
                ' Simply copy these cells to this row
                For X = 1 To UBound(Cells, 1)
                    Set pCells(X, adjacentRow) = Cells(X, Y)
                Next
            
            End If
            
        Next
    End If

    ' Nextly, initialize the non-initialized cells
    For Y = 1 To UBound(pCells, 2)
        For X = 1 To UBound(pCells, 1)
        
            ' Firstly, see if we need to create a new class
            If pCells(X, Y) Is Nothing Then
            
                ' Just create a new class
                Set pCells(X, Y) = New clsCell
                
                ' Set the default style
                Set pCells(X, Y).StyleNormal = CellStyle
                Set pCells(X, Y).StyleSelected = SelectedStyle
                
                ' We'll need to update the content
                bUpdate = True
               
            Else
                
                ' No need to update the content
                bUpdate = False
            
            End If
            
            ' Then, initialize the default values
            With pCells(X, Y)
                .Column = X
                .Row = Y + iY
                .Height = RowsHeight
                .Visible = True
                .Width = Columns(X).Width
            End With
            
            ' See if we need to update the content of the cell
            If bUpdate Then
            
                ' Make sure the cell is valid
                If Y + iY <= RowsAmount Then
                    RaiseEvent RetrieveListItem(pCells(X, Y))
                End If
                
            End If
            
        Next
    Next

End Sub

Private Sub InitializeScrolls()

    ' Firstly, see if we need to initialize the scrollbars
    If Not Scrolling.Subclassed Then
        If pRowsAmount >= RowsPerPage Or RowsWidth > picMain.ScaleWidth Then
            
            ' Enable the subclassing
            Scrolling.EnableScrolling
            
            ' Set the differnet event handlers
            Set HScroll = Scrolling.HScroll
            Set VScroll = Scrolling.VScroll

            ' Initialize page data
            HScroll.Page = picMain.ScaleWidth
            VScroll.Page = RowsPerPage

        Else
            ' No need to continue
            Exit Sub
        End If
    End If
    
    ' If we've come this far, we'll set the correct values
    If pRowsAmount >= RowsPerPage Then
        VScroll.Max = pRowsAmount
        VScroll.Visible = True
    Else
        VScroll.Visible = False
    End If
    
    ' See if we need to show the horizontal scrollbar as well
    If RowsWidth > picMain.ScaleWidth Then
        HScroll.Max = RowsWidth - picMain.ScaleWidth
        HScroll.Visible = True
    Else
        HScroll.Visible = False
    End If

End Sub

Public Property Get RowsWidth() As Long
Attribute RowsWidth.VB_Description = "Returns the total width of all the columns."

    Dim Column As clsCell

    ' Return the total width of all the columns
    For Each Column In Columns
        RowsWidth = RowsWidth + Column.StyleNormal.Margin.Left + Column.Width + Column.StyleNormal.Margin.Right
    Next

End Property

Public Property Get RowsPerPage() As Long
Attribute RowsPerPage.VB_Description = "Returns the amount of rows that fits within a single page."

    Dim lngColumn As Long
    
    ' Calculate the height of a column
    If Columns.Count > 0 Then
        lngColumn = Columns(1).StyleNormal.Margin.Top + Columns(1).Height + Columns(1).StyleNormal.Margin.Bottom
    Else
        ' Just use the default size
        lngColumn = ColumnStyle.Margin.Top + RowsHeight + ColumnStyle.Margin.Bottom
    End If
    
    ' If we're hiding the columns, ...
    If HideColumns Then
        ' obviously, they'll use no space.
        lngColumn = 0
    End If

    ' Return the amount of rows per page
    RowsPerPage = RoundUp((picMain.ScaleHeight - lngColumn) _
     / (CellStyle.Margin.Top + RowsHeight + CellStyle.Margin.Bottom))

End Property

Private Function RoundUp(Value As Double) As Long

    ' Return a rounded up value
    If Value <> Fix(Value) Then
        RoundUp = Fix(Value) + 1
    Else
        RoundUp = Fix(Value)
    End If

End Function

Private Sub cmbChange_Change()

    ' Change the current selected item
    SelectedItem.Text = cmbChange.Text
    
    ' Inform about the change
    RaiseEvent ItemChanged(SelectedItem)
    
    ' Redraw the item
    DrawCell SelectedItem
    
End Sub

Private Sub cmbChange_Click()

    ' Inform about the change
    cmbChange_Change

End Sub

Private Sub cmbChange_Scroll()

    ' Might change things this as well
    cmbChange_Click

End Sub

Private Sub picContainer_Resize()

    ' Resize the internal control correctly
    cmbChange.Width = picContainer.Width + 3

End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim oItem As clsCell, bCancel As Boolean

    ' Inform that the user has pressed a key
    RaiseEvent KeyDown(KeyCode, Shift, bCancel)

    ' Retrieve the selected item
    Set oItem = SelectedItem
    
    ' Move selector if said to
    If (Not (oItem Is Nothing)) And (Not bCancel) Then

        ' Deselect editing
        HideControls
    
        ' Execute action based upon key
        Select Case KeyCode
            Case vbKeyUp
                SelectItem oItem.Column, oItem.Row - 1
            Case vbKeyDown
                SelectItem oItem.Column, oItem.Row + 1
            Case vbKeyLeft
                SelectItem oItem.Column - 1, oItem.Row
            Case vbKeyRight
                SelectItem oItem.Column + 1, oItem.Row
        End Select
    
    End If

End Sub

Private Sub picMain_KeyPress(KeyAscii As Integer)

    ' Inform that a key has been clicked
    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)

    ' Inform that a key has been released
    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim oCell As clsCell, bCancel As Boolean, ClickData As New clsMouse

    ' Hide all controls
    HideControls

    ' Inform that the mouse has been pressed
    RaiseEvent MouseDown(Button, Shift, X, Y)

    ' See if the user has pressed an item
    Set oCell = HitTest(CLng(X), CLng(Y))
    
    ' If so, ...
    If Not (oCell Is Nothing) Then
        
        ' Initialize the click-data container
        With ClickData
            .Button = Button
            .Shift = Shift
            .X = X - oCell.X
            .Y = Y - oCell.Y
        End With
        
        ' Invoke the event (depending on the cell-type)
        If oCell.Row <= 0 Then
            
            ' Inform about the click
            RaiseEvent ColumnClick(oCell, ClickData, bCancel)
            
            ' See if we should size the column
            If Not bCancel Then
            
                ' Initialize the line
                lineMove.X1 = X
                lineMove.X2 = X
                lineMove.Y2 = picMain.ScaleHeight
                
                ' See if the user has pressed the edge (to the right)
                If ClickData.X > oCell.Width - 10 Then
                    
                    ' Select the current column
                    Set SelColumn = oCell
                    
                    ' Show the line
                    lineMove.Visible = True
                
                ' As above, only here the edge to the left
                ElseIf ClickData.X < 10 And oCell.Column > 1 Then
                    
                    ' Select the previous column
                    Set SelColumn = Columns(oCell.Column - 1)
                    
                    ' Show the line
                    lineMove.Visible = True
                    
                End If
                
                ' If we didn't size the column, ...
                If Not lineMove.Visible Then
                
                    ' If there's a pressed style, ...
                    If Not (oCell.StylePressed Is Nothing) Then
                    
                        ' ... use it.
                        oCell.Pressing = True
                        
                        ' Redraw the cell
                        DrawCell oCell
                    
                    End If
                
                End If
            
            End If
            
        Else
        
            ' Execute the event
            RaiseEvent ItemClick(oCell, ClickData, bCancel)
                        
            ' Don't save this selection if said not to
            If Not bCancel Then
                
                ' Firstly, clear the original item
                If Not (SelectedItem Is Nothing) Then
                    DrawCell SelectedItem
                End If
                
                ' See if we need to reselect the object
                If Not (SelectedItem Is oCell) Then
                    
                    ' Then select the current cell
                    Set SelectedItem = oCell
                    
                    ' Finally, draw it with the selected style
                    If oCell.StylePressed Is Nothing Then
                        DrawCell oCell, oCell.StyleSelected
                    Else
                    
                        ' Press the cell insted
                        oCell.Pressing = True
                        
                        ' Draw it
                        DrawCell oCell
                        
                    End If
                    
                End If

            End If
            
        End If
    
    End If

End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim oCell As clsCell

    ' Inform that the mouse has been moved
    RaiseEvent MouseMove(Button, Shift, X, Y)

    ' See if we're currently resizing
    If lineMove.Visible Then
        
        ' Move the line
        If X - SelColumn.X > 0 Then
            lineMove.X1 = X
            lineMove.X2 = X
        End If
        
    Else
    
        ' Use a default mouse pointer
        UserControl.MousePointer = 0
    
    End If
    
    ' Retrieve what cell we're currently over
    Set oCell = HitTest(CLng(X), CLng(Y))
    
    ' Make sure we actually found a cell
    If oCell Is Nothing Then
    
        ' Remove hovering
        ResetHovering
    
    Else
    
        ' See if this is a column
        If oCell.Row <= 0 Then
        
            ' See if we'd size the columns if the user had clicked right now
            If (X - oCell.X > oCell.Width - 10) Or (X - oCell.X < 10 And oCell.Column > 1) Then
                ' Show the size arrows
                UserControl.MousePointer = 9
            End If
        
        End If
    
        ' See if we need to hover it
        If Not oCell.Hovering Then
        
            ' Make sure there's actually a hover style
            If Not (oCell.StyleHover Is Nothing) Then
            
                ' Just remove hovering
                ResetHovering True
            
                ' Set this to the hover class
                Set CellHovering = oCell
                
                ' Initialize the timer
                tmrHover.Enabled = True
                
                ' Make it hover
                oCell.Hovering = True
                
                ' Redraw all
                RedrawAll True
            
            Else
            
                ' Remove hovering and redraw
                ResetHovering
            
            End If
        
        End If
    
    End If

End Sub

Private Sub ResetHovering(Optional bIgnoreRedraw As Boolean)
    
    ' See if we need to remove the hovering in a class
    If Not (CellHovering Is Nothing) Then
        
         ' Stop checking
        tmrHover.Enabled = False
        
        ' Make it not hover
        CellHovering.Hovering = False
        
        ' Remove this class
        Set CellHovering = Nothing
        
        ' Redraw all
        If Not bIgnoreRedraw Then
            RedrawAll True
        End If
        
    End If

End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim oCell As clsCell, oRow As clsCell, bCancel As Boolean, ClickData As New clsMouse, iY As Long

    ' Inform that the user has released the mouse
    RaiseEvent MouseUp(Button, Shift, X, Y)
        
    ' Retrieve the cell
    Set oCell = HitTest(CLng(X), CLng(Y))

    ' See if we actually found anything
    If Not (oCell Is Nothing) Then
    
        ' Initialize the click-data container
        With ClickData
            .Button = Button
            .Shift = Shift
            .X = X - oCell.X
            .Y = Y - oCell.Y
        End With
    
        ' See what kind of type this cell is
        Select Case oCell.Row
            Case Is <= 0 ' Column
            
                ' Inform about the release
                RaiseEvent ColumnReleased(oCell, ClickData)
            
            Case Else ' Cell
            
                ' Here too
                RaiseEvent ItemReleased(oCell, ClickData)
            
        End Select
        
    End If
    
    ' See if there's any column that's currently pressing
    For Each oCell In Columns
        
        If oCell.Pressing Then
        
            ' Remove the pressing
            oCell.Pressing = False
            
            ' Redraw the cell
            DrawCell oCell
        
        End If
    
        ' Then see if there's any row that needs fixing
        For iY = LBound(Cells, 2) To UBound(Cells, 2)
        
            ' Save this cell
            Set oRow = Cells(oCell.Column, iY)
            
            ' Then see if this cell is pressing downwards currently
            If oRow.Pressing Then
            
                ' Stop pressing
                oRow.Pressing = False
                
                ' Redraw this cell
                If SelectedItem Is oRow Then
                    DrawCell oRow, oRow.StyleSelected
                Else
                    DrawCell oRow
                End If
            
            End If
            
        Next
    
    Next
    
    ' See if we're currently resizing
    If lineMove.Visible Then
    
        ' Hide the line
        lineMove.Visible = False
        
        ' Inform about the resizing
        RaiseEvent ColumnResized(SelColumn, bCancel)
        
        ' Resize if said to
        If Not bCancel Then
            
            ' Resize the column
            SelColumn.Width = X - SelColumn.X
        
            ' Redraw all
            RedrawAll True
      
            ' Don't register this click
            Exit Sub
            
        End If
    
    End If
    
End Sub

Private Sub tmrHover_Timer()

    Dim Cursor As POINTAPI
    
    ' Retrieve the cursor position
    GetCursorPos Cursor

    ' Here we'll see if the mouse has moved outside of the control
    If WindowFromPoint(Cursor.X, Cursor.Y) <> picMain.hWnd Then
        
        ' Remove hovering
        ResetHovering
    
    End If

End Sub

Private Sub txtChange_Change()

    ' Change the current selected item
    SelectedItem.Text = txtChange.Text
    
    ' Inform about the change
    RaiseEvent ItemChanged(SelectedItem)
    
    ' Redraw the item
    DrawCell SelectedItem

End Sub

Private Sub UserControl_Initialize()

    ' Initialize SQL-object
    Set SQL.Parent = Me

    ' Initialize the default cell style
    With CellStyle
        .Format = DT_VCENTER Or DT_SINGLELINE
        .BackColor = vbWhite
        Set .Padding = CreateRect(4, 4, 4, 4)
        Set .Margin = CreateRect(0, 1, 0, 1)
    End With

    ' Initialize the default heading style
    With ColumnStyle
        .Format = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
        .BackColor = vbButtonFace
        .Borders.Add CreateRect(vb3DHighlight, vb3DHighlight, vb3DDKShadow, vb3DDKShadow)
        .Borders.Add CreateRect(vb3DLight, vb3DLight, vb3DShadow, vb3DShadow)
        Set .Padding = CreateRect(4, 4, 4, 4)
        Set .Margin = CreateRect(0, 0, 1, 1)
    End With
    
    ' Initialize the selected style
    Set SelectedStyle = CellStyle.Clone
    
    ' Change the style a bit
    With SelectedStyle
        .BackColor = &H8000000D
        .Borders.Add CreateRect(0, 0, 0, 0)
        .BorderStyle = Border_Dot
    End With

    ' See if XP-theme is available
    If ColumnStyle.Theme.IsOK(UserControl.hWnd, "Header") Then

        ' Remove the borders
        Set ColumnStyle.Borders = New Collection
    
        ' Use XP-themes
        With ColumnStyle.Theme
            .Class = "Header"
            .PartID = 1
            .StateID = 1
        End With
        
        ' Create two more styles
        Set ColumnHover = ColumnStyle.Clone
        Set ColumnPress = ColumnHover.Clone
        
        ' Set the states
        ColumnHover.Theme.StateID = 2
        ColumnPress.Theme.StateID = 3
        
    End If

    ' Initialize scrolling
    Scrolling.ControlWnd = UserControl.hWnd

    ' Set default values
    RowsHeight = 19
    ColumnHeight = 19
    BackColor = vbWhite

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Execute the properties
    Me.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    Me.IgnoreUpdates = PropBag.ReadProperty("IgnoreUpdates", True)
    Me.RowsAmount = PropBag.ReadProperty("pRowsAmount", 0)
    Me.RowsHeight = PropBag.ReadProperty("RowsHeight", 19)
    Me.BorderStyle = PropBag.ReadProperty("BorderStyle", Border_FixedSingle)
    Me.HideColumns = PropBag.ReadProperty("HideColumns", False)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    ' Save properties
    PropBag.WriteProperty "BackColor", Me.BackColor, vbWhite
    PropBag.WriteProperty "IgnoreUpdates", Me.IgnoreUpdates, True
    PropBag.WriteProperty "pRowsAmount", Me.RowsAmount, 0
    PropBag.WriteProperty "RowsHeight", Me.RowsHeight, 19
    PropBag.WriteProperty "BorderStyle", Me.BorderStyle, Border_FixedSingle
    PropBag.WriteProperty "HideColumns", Me.HideColumns, False

End Sub

Private Sub UserControl_Terminate()

    ' Always uninitialise the scrolls
    Scrolling.FinishScrolling

End Sub

Private Sub UserControl_Resize()

    ' Firstly, fix the padding
    FixPadding
    
    ' Save the page-data
    If Not (VScroll Is Nothing) Then
        VScroll.Page = RowsPerPage
    End If
    
    ' Then, redraw all
    RedrawAll
    
End Sub

Private Sub FixPadding()

    ' Resize the main picturebox (take the padding into account)
    picMain.Left = Padding.Left
    picMain.Top = Padding.Top
    picMain.Width = UserControl.Width - picMain.Left - Padding.Right
    picMain.Height = UserControl.Height - picMain.Top - Padding.Bottom

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets/returns the backcolor of the list view."

    ' Return the backcolor
    BackColor = picMain.BackColor

End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)

    ' Set the color of both the control and the picturebox
    picMain.BackColor = vNewValue
    UserControl.BackColor = vNewValue

End Property

Public Property Get BorderStyle() As BorderStyleForm
Attribute BorderStyle.VB_Description = "Sets/returns the border style of this control."

    ' Return the type of border
    BorderStyle = UserControl.BorderStyle

End Property

Public Property Let BorderStyle(ByVal vNewValue As BorderStyleForm)

    ' Set the type of border
    UserControl.BorderStyle = vNewValue

End Property

Private Sub VScroll_Changed(Cancel As Boolean)

    ' Hide control
    HideControls

    ' Redraw everything
    RedrawAll True

End Sub

Private Sub VScroll_Scrolling(ByVal Position As Long)

    ' Hide control
    HideControls

     ' Yet again, redraw everything
    RedrawAll True, , Position

End Sub

Private Sub HScroll_Changed(Cancel As Boolean)

    ' Hide control
    HideControls

    ' As above
    RedrawAll True

End Sub

Private Sub HScroll_Scrolling(ByVal Position As Long)

    ' Hide control
    HideControls

    ' As above
    RedrawAll True, Position

End Sub
