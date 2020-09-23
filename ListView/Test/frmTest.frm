VERSION 5.00
Object = "{DEF372A3-985E-4301-8219-0350183FC961}#18.0#0"; "VirtualListView.ocx"
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   532
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   483
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   420
      Width           =   855
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "SELECT * FROM Songs"
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "D:\\Kristian\\My Music\\MediaMonkey\\MediaMonkey.mdb"
      Top             =   240
      Width           =   4455
   End
   Begin VirtualListView.usrListView usrListView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8916
      IgnoreUpdates   =   0   'False
   End
   Begin VB.Label lblCommand 
      Caption         =   "SQLCommand:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   615
      Width           =   1335
   End
   Begin VB.Label lblSQL 
      Caption         =   "SQLSource:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oSecond As clsStyle, bDirection As Boolean

Private Sub cmdConnect_Click()

    ' Initialize the list view
    With usrListView1
    
        ' Ignore redraws (to speed things up)
        .IgnoreUpdates = True
    
        ' Open the database
        .SQL.Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtPath & ";" & _
         "User Id=admin;Password=;"
        .SQL.Command = txtCommand
        .SQL.Enable = True
        
        ' Allow redrawing
        .IgnoreUpdates = False
    
        ' Update all
        .UpdateContent
        .RedrawAll
    
    End With

End Sub

Private Sub Form_Load()

    Dim oBorder As New clsRect

    ' Initialize the list view
    With usrListView1

        ' Create a border
        With oBorder
            .Right = RGB(230, 230, 230)
            .Left = -1
            .Top = -1
            .Bottom = .Right
        End With

        ' Add a border
        .CellStyle.Borders.Add oBorder

        ' Use a second style as well
        Set oSecond = .CellStyle.Clone
            
        ' Change the background color of the second style set
        oSecond.BackColor = RGB(250, 250, 250)
    
        ' Ignore redraws (to speed things up)
        .IgnoreUpdates = True
                
        ' Add default columns
        .AppendColumn "ID"
        .AppendColumn "Text"
        .AppendColumn "Date"
        
        ' Allow redrawing
        .IgnoreUpdates = False
        
        ' Set the amount of rows
        .RowsAmount = 1000000
    
        ' Redraw all
        .RedrawAll

    End With

End Sub

Private Sub usrListView1_ColumnClick(Column As VirtualListView.clsCell, ClickData As VirtualListView.clsMouse, Cancel As Boolean)

    ' Be sure this is a valid column click
    If (ClickData.X > 10 And ClickData.X < Column.Width - 10) Then
        If usrListView1.SQL.Enable Then
        
            ' Hide changes
            usrListView1.IgnoreUpdates = True
        
            ' Set SQL-command
            usrListView1.SQL.Command = txtCommand & " ORDER BY " & Left(Column.Text, _
             Len(Column.Text) - 1) & IIf(bDirection, " DESC", "")
            usrListView1.SQL.Enable = True
            
            ' Stop hiding changes
            usrListView1.IgnoreUpdates = False
            
            ' Redraw
            usrListView1.UpdateContent
        
            ' Rotate direction
            bDirection = Not bDirection
        
        End If
    End If

End Sub

Private Sub Form_Resize()

    ' Resize the controls
    usrListView1.Width = Me.ScaleWidth - 16
    usrListView1.Height = Me.ScaleHeight - usrListView1.Top - 8
    txtPath.Width = Me.ScaleWidth - txtPath.Left - cmdConnect.Width - 24
    txtCommand.Width = txtPath.Width
    cmdConnect.Left = txtCommand.Width + txtCommand.Left + 8
 
End Sub

Private Sub usrListView1_ItemClick(Item As clsCell, ClickData As clsMouse, Cancel As Boolean)

    ' Start editing the cell if this item is selected as well
    If Not (usrListView1.SelectedItem Is Nothing) Then
        If Item.Row = usrListView1.SelectedItem.Row And _
         Item.Column = usrListView1.SelectedItem.Column Then
            usrListView1.EditCell Item
        End If
    End If

End Sub

Private Sub usrListView1_RetrieveListItem(Item As clsCell)

    ' Set the styleset depending on the row
    If (Item.Row Mod 2) = 0 Then
        Set Item.StyleNormal = oSecond
    End If
    
    ' Return a default text if we're not using a SQL-source
    If Not usrListView1.SQL.Enable Then
    
        ' Check what column the item is placed within
        Select Case Item.Column
            Case 1 ' ID
                Item.Text = Item.Row
            Case 2 ' Text
                Item.Text = IIf(Rnd > 0.5, "Y", "F")
                Item.EditControl = Edit_ComboBox
                Item.EditData = Array("F", "Y")
            Case 3 ' Date
                Item.Text = Now
        End Select

    End If

End Sub
