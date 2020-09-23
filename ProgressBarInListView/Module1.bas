Attribute VB_Name = "Module1"
Option Explicit

'****************************User definded Datatypes****************
Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

' The NMHDR structure contains information about a notification message.
Public Type NMHDR
    hwndFrom    As Long        ' Window handle of control that sends the message
    idFrom      As Long        ' Identifier of control that sends the message
    code        As Long        ' Notification code
End Type

'**************************** Constants ****************************
Public Const LVM_FIRST              As Long = &H1000
Private Const LVM_GETSUBITEMRECT    As Long = (LVM_FIRST + 56)
'Retrieves information about the bounding rectangle
'for a subitem in a list view control.

Public Const LVIR_LABEL     As Long = 2
'Returns the bounding rectangle of the entire item,
'including the icon and label.


Public Const WM_NOTIFY  As Long = &H4E
Public Const WM_HSCROLL As Long = &H114
Public Const WM_VSCROLL As Long = &H115
Public Const WM_KEYDOWN As Long = &H100
'The WM_HSCROLL and WM_VSCROLL messages are sent to a window when
'a scroll event occurs in the window's standard scroll bar.
'This messages are also sent to the owner of a horizontal/vertical scroll bar control
'when a scroll event occurs in the control.

Public Const HDN_FIRST      As Long = (0 - 300)
Public Const HDN_ENDTRACK   As Long = (HDN_FIRST - 1)
'The HDN_ENDTRACK Notifies a header control's parent window
'that the user has finished dragging a divider.



'****************************API Declarations***********************
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                             ByVal wMsg As Long, _
                             ByVal wParam As Long, _
                             lParam As Any) As Long

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                         ByVal hWndNewParent As Long) As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                             Source As Any, _
                                                             ByVal Length As Long)


Public Function ListView_GetSubItemRect(ByVal hWndLV As Long, _
                                        ByVal iItem As Long, _
                                        ByVal iSubItem As Long, _
                                        ByVal code As Long, _
                                        lpRect As RECT) As Boolean

'Get the Coordinates of the ListItem specified with iITEM and iSubItem
  lpRect.Top = iSubItem
  lpRect.Left = code
  ListView_GetSubItemRect = SendMessage(hWndLV, LVM_GETSUBITEMRECT, ByVal iItem, lpRect)
End Function

    

