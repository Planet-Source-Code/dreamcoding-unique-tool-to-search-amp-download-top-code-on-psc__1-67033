Attribute VB_Name = "modListview"
'Source is From: http://www.visualbasic.happycodings.com/Forms/code60.html
'Sort, Select and Return the selected items from a ListView

Option Explicit
'-----------------------------ListView API----------------------
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'-----------------------------ListView messages-----------------
Private Const LVM_FIRST = &H1000
Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Private Const LVNI_SELECTED = &H2
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Private Const MAX_PATH As Long = 260
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)
Private Const LVIF_STATE = &H8
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVIS_STATEIMAGEMASK As Long = &HF000
Private Const LVIF_TEXT = &H1
 
Private Type LV_ITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   state        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type


'Purpose     :  Returns the selected items text from a ListView
'Inputs      :  lvGetSelected                   The Listview to get the selected items from
'               asSelected                      See outputs
'               [lStartCol]                     If specified the first column in the list to return,
'                                               else starts at the 0 column (the text column).
'               [lEndCol]                       If specified the last column in the list to return,
'                                               else ends at the last column.
'Outputs     :  Returns a count of the items selected or -1 on error.
'               asSelected()                            A 2d string array. Containing the selected items in the listview.
'                                                       Format of array asSelected(lStartCol to lEndCol, 1 to ItemsSelected),
'                                                       Where 0 is the first column in the listview
'Notes       :  Requires a reference to Microsoft Windows Common Controls


Function LVSelectedItems(lvGetSelected As ListView, ByRef asSelected() As String, Optional ByVal lStartCol As Long = 0, Optional ByVal lEndCol As Long = -1) As Long
    Dim lThisItem As Long, lThisCol As Long
    Dim lLvHwnd As Long, lSelectedItemIndex As Long, lItemsSelected As Long
    Dim tListItem As LV_ITEM, lItemLen As Long
    Const clMaxItemText As Long = 200
    
    On Error GoTo ErrFailed
    lLvHwnd = lvGetSelected.hwnd
    'Clear the output array
    Erase asSelected
    'Determine the number of selected items
    lItemsSelected = SendMessage(lLvHwnd, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
    
    If lItemsSelected Then
        With lvGetSelected
            If lEndCol = -1 Then
                'Default to last column in listview
                lEndCol = .ColumnHeaders.Count - 1
            End If
            
            ReDim asSelected(lStartCol To lEndCol, 1 To lItemsSelected)
            
            With .ListItems
                lSelectedItemIndex = -1
                tListItem.cchTextMax = clMaxItemText
                tListItem.pszText = Space(clMaxItemText)
                tListItem.mask = LVIF_TEXT
                
                'Get the text from each of the selected rows
                For lThisItem = 1 To lItemsSelected
                    'Get the item index
                    lSelectedItemIndex = SendMessage(lLvHwnd, LVM_GETNEXTITEM, lSelectedItemIndex, ByVal LVNI_SELECTED)
                    
                    'Get the text from each of the columns
                    For lThisCol = lStartCol To lEndCol
                        tListItem.iSubItem = lThisCol
                        'Get the sub item
                        lItemLen = SendMessage(lvGetSelected.hwnd, LVM_GETITEMTEXT, lSelectedItemIndex, tListItem)
                        'Trim text
                        asSelected(lThisCol, lThisItem) = Left$(tListItem.pszText, lItemLen)
                    Next
                Next
            End With
        End With
    End If
    'Return the count of the selected items
    LVSelectedItems = lItemsSelected
    
    Exit Function

ErrFailed:
    'Return error code
    LVSelectedItems = -1
    On Error Resume Next
End Function

'Purpose     :  Unselects all the selected items in a listview
'Inputs      :  oListView                 The Listview to unselect all the items from.
'Outputs     :  Returns the error number if an error occurred
'Notes       :  Requires a reference to Microsoft Windows Common Controls


Function LVDeselectAll(oListView As ListView) As Boolean
    Dim sThisItem As Long, lLvHwnd As Long, lSelectedItems As Long, lItemIndex As Long
    
    On Error GoTo ErrFailed
    
    With oListView
        lLvHwnd = .hwnd
        .Visible = False             'For speed. Need to remove the line in VBA
        lSelectedItems = SendMessage(lLvHwnd, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
        lItemIndex = -1
        For sThisItem = 1 To lSelectedItems
            lItemIndex = SendMessage(lLvHwnd, LVM_GETNEXTITEM, lItemIndex, ByVal LVNI_SELECTED)
            .ListItems(lItemIndex + 1).Selected = False
        Next
        .Visible = True              'For speed. Need to remove the line in VBA
    End With
    Exit Function
    
ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
    LVDeselectAll = True
End Function


'Purpose     :  Applies a reverse sort to the selected column header.
'               Add a call to this routine in the ColumnClick event of the listview.
'               eg.
'               Private Sub lvDemo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'                   LVSortColumns lvDemo, ColumnHeader
'               End Sub
'
'Inputs      :  LVSort                          The listview to sort.
'               LVColumnHeader                  The column to sort on.
'Outputs     :  Returns the error number if an error occurred
'Notes       :  Requires a reference to Microsoft Windows Common Controls


Function LVSortColumns(LVSort As ListView, LVColumnHeader As ColumnHeader) As Long
    
    On Error GoTo ErrFailed
    With LVSort
        'HACK: Protects against an occassional 'division by zero' general protection fault when sorting an empty listview
        If .ListItems.Count > 0 Then
            .Visible = False        'For speed. Need to remove the line in VBA
            .SortKey = LVColumnHeader.Index - 1
            .SortOrder = 1 - LVSort.SortOrder
            .Sorted = True
            .Visible = True         'For speed. Need to remove the line in VBA
        End If
    End With
    
    Exit Function
    
ErrFailed:
    Debug.Assert False
    LVSortColumns = Err.Number
    On Error Resume Next
End Function


'Purpose     :  Determines if an item exists in a listview
'Inputs      :  oLv                     The listview to populate
'               sKeyName                The key of the item check exists
'Outputs     :  Returns True if the item exists else returns false
'Notes       :  Requires a reference to MSCOMCTL.OCX or COMCTL.OCX


Function LvItemExists(oLv As ListView, sKeyName As String) As Boolean
    Dim bTest As Boolean
    On Error GoTo ErrFailed
    bTest = oLv.ListItems(sKeyName).Bold
    LvItemExists = True
    Exit Function

ErrFailed:
    LvItemExists = False
    On Error GoTo 0
End Function

Function GetSelectedCount(lstView As ListView)
Dim lLvHwnd As Long
    lLvHwnd = lstView.hwnd
    GetSelectedCount = SendMessage(lLvHwnd, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
End Function
