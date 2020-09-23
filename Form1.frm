VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   405
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Tag             =   "string_maxlength"
      Text            =   "Form1.frx":0000
      Top             =   6300
      Width           =   1725
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   405
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Tag             =   "subclass"
      Text            =   "Form1.frx":06D5
      Top             =   5940
      Width           =   1725
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   450
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Tag             =   "Pause"
      Text            =   "Form1.frx":0AFF
      Top             =   5490
      Width           =   1725
   End
   Begin codeGold.splitter splitter1 
      Height          =   5145
      Left            =   2925
      TabIndex        =   2
      Top             =   -45
      Width           =   45
      _extentx        =   79
      _extenty        =   9075
   End
   Begin VB.TextBox Text1 
      Height          =   5010
      Left            =   3060
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   45
      Width           =   6585
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2760
   End
   Begin VB.Label lblHelp 
      Caption         =   $"Form1.frx":100F
      Height          =   375
      Index           =   2
      Left            =   2250
      TabIndex        =   8
      Tag             =   "subclass"
      Top             =   6255
      Width           =   2175
   End
   Begin VB.Label lblHelp 
      Caption         =   "basic code for subclassing"
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Tag             =   "subclass"
      Top             =   5895
      Width           =   2175
   End
   Begin VB.Label lblHelp 
      Caption         =   "Creates a control wait time but still allows execution of other code.  A cross between ""DoEvents"" and ""Sleep"""
      Height          =   420
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Tag             =   "pause"
      Top             =   5445
      Width           =   2175
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSub1 
         Caption         =   "show help &tips"
         Index           =   5
      End
      Begin VB.Menu mnuPopupSub1 
         Caption         =   "&copy code to clipboard"
         Index           =   10
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


 
Private Sub Form_Load()
  
  With splitter1
     .Width = 75
      Set .Control_Bottom_Or_Right = Text1
      Set .Control_Top_Or_Left = List1
  End With
  
  With List1
    .AddItem "Pause"
    .AddItem "Subclass"
    .AddItem "string_MaxLength"
  End With
 
  
End Sub
 
Private Sub List1_Click()
  
  With List1
     ' if an item in the listbox has been selected..
     If .ListIndex > -1 Then
       Dim l As Long
       ' loop through the entire array of textboxes
       ' that hold all the code snippets
       For l = 0 To txtCode.UBound
         ' if that textboxes tag = the listitem selected
         ' then thats the cod we want so put it in text1
         If Trim$(LCase$(txtCode(l).Tag)) = Trim$(LCase$(.List(.ListIndex))) Then
            Text1 = txtCode(l)
            
            ' if ftip is visible then show the associated help info
            If Ftip.Visible Then
               Call Ftip.DisplayHelpTip(lblHelp(l))
            End If
         
            Exit Sub
         End If
       Next l
     End If
  End With
  
End Sub

Private Sub mnuPopupSub1_Click(Index As Integer)
  
  If Index = 5 Then
     Ftip.Show vbModeless, Me
     Call List1_Click
     
  ElseIf Index = 10 Then
    ' copy whats in text1 to clipboard
    Clipboard.Clear
    Clipboard.SetText Text1
  
  End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = 1 Then PopupMenu mnuPopup
  
End Sub

