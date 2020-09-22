VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListBoxHScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ITEM_MARGIN = 4
Private Const CHECKBOX_WIDTH = 14

' --- required API declarations ---

Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const WM_VSCROLL = &H115
Private Const SB_BOTTOM = 7

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" _
   (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
   lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_VSCROLL = &H200000

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXVSCROLL = 2

' --- private class variables ---
Private mvarListBox As ListBox
Private m_lMaxItemWidth As Long
Private m_hItemFont As Long
Private m_ListBoxHwnd As Long

' Call this sub before you will use other methods of this class
Public Sub Attach(ByRef pListBox As ListBox)
   Dim FontInt As IFont

   Set mvarListBox = pListBox
   m_ListBoxHwnd = mvarListBox.hwnd
   
   ' Determining the handle of the font used in the specified listbox.
   ' Using the IFont interface we can retreive the handle of the font
   ' used in the specified listbox.
   ' We'll use this handle further when we'll calculate the width of
   ' listbox items
   Set FontInt = pListBox.Font
   m_hItemFont = FontInt.hFont
End Sub

' The following routine adds a string to a specified list box
' and displays the horizontal scroll bar in it if required
Public Sub AddItem(ByRef psItemText As String, Optional ByVal pbScrollBottom As Boolean = True)
   Dim lItemWidth As Long
   
   mvarListBox.AddItem psItemText
   
   ' --- calculating the width of the currently added item ---
   lItemWidth = flGetItemWidth(psItemText)
   
   ' --- determining whether we need to display the horizontal scroll bar ---
   If lItemWidth > m_lMaxItemWidth Then
      m_lMaxItemWidth = lItemWidth
      pSetHorExtent
   End If
   
   ' --- scrolling the listbox to be sure that the user see the last item ---
   If pbScrollBottom Then
      SendMessageByLong m_ListBoxHwnd, WM_VSCROLL, SB_BOTTOM, 0
   End If
End Sub

Public Sub Clear()
   mvarListBox.Clear
   m_lMaxItemWidth = 0
   SendMessageByLong m_ListBoxHwnd, LB_SETHORIZONTALEXTENT, 0, 0
End Sub

' Use this write-only property to change the caption of a list box item
Public Property Let List(ByVal piIndex As Integer, psItemText As String)
   Dim lPrevItemWidth As Long
   Dim lNewItemWidth As Long
   
   lPrevItemWidth = flGetItemWidth(mvarListBox.List(piIndex))
   mvarListBox.List(piIndex) = psItemText
   lNewItemWidth = flGetItemWidth(psItemText)
   
   If lPrevItemWidth = m_lMaxItemWidth Then
      If lNewItemWidth > lPrevItemWidth Then
         m_lMaxItemWidth = lNewItemWidth
         pSetHorExtent
      Else
         RefreshHScroll
      End If
   Else
      If lNewItemWidth > m_lMaxItemWidth Then
         m_lMaxItemWidth = lNewItemWidth
         pSetHorExtent
      End If
   End If
End Property

Public Sub pSetHorExtent()
   SendMessageByLong m_ListBoxHwnd, LB_SETHORIZONTALEXTENT, _
      m_lMaxItemWidth + ITEM_MARGIN + IIf(mvarListBox.Style = vbListBoxCheckbox, CHECKBOX_WIDTH, 0), 0
End Sub

Public Sub RemoveItem(ByVal piIndex As Integer)
   Dim lItemWidth As Long
   
   lItemWidth = flGetItemWidth(mvarListBox.List(piIndex))
   
   mvarListBox.RemoveItem piIndex
   
   If mvarListBox.ListCount = 0 Then
      m_lMaxItemWidth = 0
      SendMessageByLong m_ListBoxHwnd, LB_SETHORIZONTALEXTENT, 0, 0
      Exit Sub
   End If
   
   If lItemWidth = m_lMaxItemWidth Then
      ' we remove the item with the maximum width
      ' and must recalculate the maximum width and the
      ' horizontal extent of our list box
      RefreshHScroll
   End If
End Sub

' recalculates the maximum width for all items and
' displays the horizontal scroll bar if required
Public Sub RefreshHScroll()
   Dim hdc As Long
   Dim rc As RECT
   Dim hOldFont As Long
   Dim i As Integer
   Dim lItemWidth As Long
   
   ' We use the same technique like in the flGetItemWidth function,
   ' but sligtly optimize it: we don't need to select/deselect the
   ' required font for each item - we can do it at once outside the cycle
   hdc = GetDC(m_ListBoxHwnd) ' retrieving HDC for the listbox
   hOldFont = SelectObject(hdc, m_hItemFont) ' selecting the required font
   
   m_lMaxItemWidth = 0
   For i = 0 To mvarListBox.ListCount - 1
      DrawText hdc, mvarListBox.List(i), -1, rc, DT_SINGLELINE + DT_CALCRECT
      lItemWidth = rc.Right - rc.Left
      If lItemWidth > m_lMaxItemWidth Then
         m_lMaxItemWidth = lItemWidth
      End If
   Next
   
   ' restoring the previous font
   Call SelectObject(hdc, hOldFont)
   ReleaseDC m_ListBoxHwnd, hdc
   
   pSetHorExtent
End Sub

Private Function flGetItemWidth(psItemText As String)
   Dim hdc As Long
   Dim rc As RECT
   Dim hOldFont As Long
   
   hdc = GetDC(m_ListBoxHwnd) ' retrieving HDC for the listbox
   hOldFont = SelectObject(hdc, m_hItemFont) ' selecting the required font
   ' if you specify the DT_CALCRECT flag,
   ' DrawText only Determines the width and height of the rectangle
   ' required to display the text:
   DrawText hdc, psItemText, -1, rc, DT_SINGLELINE + DT_CALCRECT
   flGetItemWidth = rc.Right - rc.Left
   ' restoring the previous state
   Call SelectObject(hdc, hOldFont)
   ReleaseDC m_ListBoxHwnd, hdc
End Function

' Returns the minimum width of the list box
' when the horizontal scroll bar disappears
' NOTE: the return value in pixels
Public Property Get MinWidthNoHScroll() As Long
   Dim bHasVScrBar As Boolean
   
   bHasVScrBar = GetWindowLong(m_ListBoxHwnd, GWL_STYLE) And WS_VSCROLL
   
   MinWidthNoHScroll = m_lMaxItemWidth + IIf(bHasVScrBar, GetSystemMetrics(SM_CXVSCROLL), 0) + _
      IIf(mvarListBox.Style = vbListBoxCheckbox, CHECKBOX_WIDTH, 0) + _
      ITEM_MARGIN + 4
End Property
