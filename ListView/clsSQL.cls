VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Used to simplify data retrival from a database

' The connection string including the SQL command to utilize
Public Connection As String
Public Command As String

' The parent list view
Public Parent As usrListView

' The event handler to use
Private WithEvents EventHandler As usrListView
Attribute EventHandler.VB_VarHelpID = -1

' Whether or not we have enabled this module
Private bEnable As Boolean

' The different ADODB-objects
Private oConnection As ADODB.Connection
Private oRecordset As ADODB.Recordset

' The current recordset we're using
Public Property Get Table() As Object

    ' Return the current recordset
    Set Table = oRecordset

End Property

Public Property Get Enable() As Boolean
    
    ' Return the enable state
    Enable = bEnable

End Property

' Used to turn on the SQL-functionality
Public Property Let Enable(ByVal vNewValue As Boolean)

    Dim Field As Field

    ' Save this value
    bEnable = vNewValue

    ' Clean up resources
    CleanUp

    ' See if we are to enable the SQL functionality
    If vNewValue Then
        
        ' Create object
        Set oConnection = New ADODB.Connection
        
        ' If so, initialize the database
        oConnection.CursorLocation = adUseClient ' To enable AbsolutePosition
        oConnection.Open Connection
    
        ' Open recordset
        Set oRecordset = oConnection.Execute(Command)
        
        ' Clear all other columns
        Set Parent.Columns = New Collection
        
        ' Add columns to the list view
        For Each Field In oRecordset.Fields
            Parent.AppendColumn Field.Name & ":"
        Next
        
        ' Set the amount of rows in the recordset
        Parent.RowsAmount = oRecordset.RecordCount
    
        ' Initialize the event handler
        Set EventHandler = Parent
    
    End If

End Property

Private Sub CleanUp()

    ' Close classes if they're initialized
    If Not (oRecordset Is Nothing) Then
        oRecordset.Close
        oConnection.Close
    End If

    ' Clean up all
    Set oRecordset = Nothing
    Set oConnection = Nothing
    
    ' Remove references
    Set EventHandler = Nothing

End Sub

Private Sub Class_Terminate()

    ' Clean up all
    CleanUp

End Sub

Private Sub EventHandler_RetrieveListItem(Item As clsCell)

    On Error GoTo IncorrectRow

    ' See if we need to change the row-position
    If Item.Row <> oRecordset.AbsolutePosition Then
        oRecordset.AbsolutePosition = Item.Row
    End If
    
    ' Set the content of this cell
    Item.Text = oRecordset.Fields(Item.Column - 1).Value

    ' All went well
    Exit Sub
    
' In the event of an invalid row index, ...
IncorrectRow:

    ' ... just return nothing.
    Item.Text = vbNullString

End Sub
