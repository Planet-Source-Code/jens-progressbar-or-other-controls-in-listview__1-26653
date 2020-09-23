VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "ProgressBar in ListView Demo"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1080
      Top             =   2640
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'.------------------------------------------------------------------------------
'. This is Demo project showing how to place progressbars into a listview control.
'. It should work with other controls as well e.g. CommandButtons,
'. TextBox, ComboBox, PictureBox ... (with only little modifications).
'. For subclassing I used the SoftCircuits Subclass Control
'. You can get it for free at http://www.softcircuits.com/
'.
'.------------------------------------------------------------------------------
'Author : Jens Schiefer
'Date   : 08-25-01

'Variables for the Progressbar Values set in the Timer Event
Private m_lngCount1     As Long
Private m_lngCount2     As Long

'The Column in the listview in which you want to place the Progressbar
Private Const mc_lngCol As Long = 1


Private Sub Form_Load()
    Dim lngI As Long
    'Add some ListItems to Listview1
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Add , , "Col0"
    ListView1.ColumnHeaders.Add , , "Col1"
    ListView1.ColumnHeaders.Add , , "Col2"
    ListView1.ListItems.Add , , "0"

    For lngI = 1 To 10
        ListView1.ListItems.Add , , lngI
    Next
    
    Call PutProgressBarInListview
End Sub

Private Sub PutProgressBarInListview()
    'A ProgressBar is created for each Row in ListView, resized
    'and moved to the right place
    
    Dim lngRow As Long
    Dim lngCol As Long
    'First we have to check if listview1 is in report mode
    If ListView1.View <> lvwReport Then Exit Sub
    
    For lngRow = 0 To ListView1.ListItems.Count - 1
        'Create a progressbar for each  Row in the Listview
        If lngRow > ProgressBar1.Count - 1 Then
            Load ProgressBar1(lngRow)
        End If
        'Use SetParent to make the Listview Control the Parent Window of the Progressbar
        Call SetParent(ProgressBar1(lngRow).hWnd, ListView1.hWnd)
    Next
    'Now the progressbars size must be changed and they have to be moved
    'in the right place
    Call AdjustProgressBar
    
    'Subclass the Listview Control
    Subclass1.hWnd = ListView1.hWnd
    'Intercept WM_HSCROLL and WM_VSCROLL so you can adjust the progressbars if the
    'Listview is scrolled
    Subclass1.Messages(WM_HSCROLL) = True
    Subclass1.Messages(WM_VSCROLL) = True
    'WM_KEYDOWN Is sent to listview if someone presses a key
    'Needed for Keyboard Scrolling
    Subclass1.Messages(WM_KEYDOWN) = True
    'Intercept WM_NOTIFY to adjust the progressbar if someone changes the size of
    'the columns
    Subclass1.Messages(WM_NOTIFY) = True
    
End Sub

Private Sub AdjustProgressBar()
    'This Sub is called when the size or place of the listItems change
    'It puts the progressbars in the right places again
    Dim rcPos       As RECT
    Dim lngRow      As Long
    Dim blnRect     As Boolean
    
    For lngRow = 0 To ProgressBar1.Count - 1
        blnRect = ListView_GetSubItemRect(ListView1.hWnd, lngRow, mc_lngCol, LVIR_LABEL, rcPos)
        With ProgressBar1(lngRow)
            .Left = (rcPos.Left) * Screen.TwipsPerPixelX
            .Width = (rcPos.Right - rcPos.Left) * Screen.TwipsPerPixelX
            .Height = ((rcPos.Bottom - rcPos.Top) * Screen.TwipsPerPixelY) / 2
            .Top = rcPos.Top * Screen.TwipsPerPixelY + ((rcPos.Bottom - rcPos.Top) * Screen.TwipsPerPixelY - ProgressBar1(lngRow).Height) / 2
        End With
    'If a ProgressBar disappears on top because of vertical scrolling its made
    'invisible if we don't it will apear above the Header
        If rcPos.Top <= 3 Then
            ProgressBar1(lngRow).Visible = False
        Else
            ProgressBar1(lngRow).Visible = True
        End If
    Next
End Sub

Private Sub ProgressBar1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    'If ProgressBar is pressed its row will become the selected row
    'but only If FullRowSelect is True
    ListView1.SetFocus
    If ListView1.FullRowSelect = True Then
        Set ListView1.SelectedItem = ListView1.ListItems(Index + 1)
    End If
End Sub


Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    Dim uWMNOTIFY_Message   As NMHDR
    Dim lngCode             As Long
    Dim blnAdjust           As Boolean
    
    Select Case Msg
        Case WM_HSCROLL, WM_VSCROLL
            'Someone just scrolled so we need to adjust the ProgressBars
            blnAdjust = True
        Case WM_KEYDOWN
            Select Case wParam
                Case 33 To 40
                'If someone scrolls using the keyboard
                '33 to 40 are the KeyCodes of pgDown, pgUp, Left, Right and so on
                blnAdjust = True
            End Select
        Case WM_NOTIFY
            'Some event has occured in the listview control
            'let's see if someone adjusted the size of the columns
            CopyMemory uWMNOTIFY_Message, ByVal lParam, Len(uWMNOTIFY_Message)
            lngCode = uWMNOTIFY_Message.code
            Select Case lngCode
                Case HDN_ENDTRACK
                    'HDN_ENDTRACK is sent via WM_NOTIFY when the width
                    'of the columns is changed by the user
                    blnAdjust = True
            End Select
    End Select
    Result = Subclass1.CallWndProc(Msg, wParam, lParam)
    If blnAdjust = True Then Call AdjustProgressBar
End Sub


Private Sub Timer1_Timer()
    'Sets the ProgressBar Values
    Dim lngI    As Long
    m_lngCount1 = m_lngCount1 + 5
    m_lngCount2 = m_lngCount2 + 20
    For lngI = 0 To ProgressBar1.Count - 1
        If lngI Mod 2 = 0 Then
            ProgressBar1(lngI).Value = m_lngCount1
        Else
            ProgressBar1(lngI).Value = m_lngCount2
        End If
    Next
    If m_lngCount1 = 100 Then m_lngCount1 = 0
    If m_lngCount2 = 100 Then m_lngCount2 = 0
End Sub








