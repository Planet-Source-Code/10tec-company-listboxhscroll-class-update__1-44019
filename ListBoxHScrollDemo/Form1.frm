VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ListBoxHScroll Class Demo"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtItemCnt 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Text            =   "50"
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton cmdRemoveHScr 
      Caption         =   "Change the form width to remove the horizontal scroll bar"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connecting to the populated listbox"
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdChangeCap 
      Caption         =   "Change caption of the last item"
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add item"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2100
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemoveLast 
      Caption         =   "Remove the last item"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1620
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3240
      IntegralHeight  =   0   'False
      Left            =   2220
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   4395
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate the listbox"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of items:"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   240
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXFRAME = 32

Dim LBHS As CListBoxHScroll

Private Sub cmdAddItem_Click()
   LBHS.AddItem "New item: " + String(Int(Rnd * 30) + 1, "W") + "!"
End Sub

Private Sub cmdChangeCap_Click()
   LBHS.List(List1.ListCount - 1) = "New caption: " + String(Int(Rnd * 30) + 1, "W") + "!"
End Sub

Private Sub cmdClear_Click()
   LBHS.Clear
End Sub

Private Sub cmdConnect_Click()
   Dim i As Long
   
   List1.Clear
   
   For i = 1 To Val(txtItemCnt)
      List1.AddItem "Item #" & i & ": " & String(Int(Rnd * 30) + 1, "W") & "!"
   Next
   
   LBHS.RefreshHScroll
End Sub

Private Sub cmdPopulate_Click()
   Dim i As Long
   
   LBHS.Clear
   
   For i = 1 To Val(txtItemCnt)
      LBHS.AddItem "Item #" & i & ": " & String(Int(Rnd * 30) + 1, "W") & "!"
   Next
End Sub

Private Sub cmdRemoveHScr_Click()
   Me.Width = List1.Left + _
      LBHS.MinWidthNoHScroll * Screen.TwipsPerPixelX + _
      2 * GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX + _
      120
End Sub

Private Sub cmdRemoveLast_Click()
   If List1.ListCount = 0 Then
      MsgBox "Nothing to remove!", vbCritical
   Else
      LBHS.RemoveItem List1.ListCount - 1
   End If
End Sub

Private Sub Form_Load()
   With List1.Font
      .Name = "Arial"
      .Size = 12
      .Italic = True
   End With
   
   Set LBHS = New CListBoxHScroll
   LBHS.Attach List1
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   List1.Move List1.Left, List1.Top, Me.ScaleWidth - List1.Left - 120, Me.ScaleHeight - List1.Top - 120
End Sub
