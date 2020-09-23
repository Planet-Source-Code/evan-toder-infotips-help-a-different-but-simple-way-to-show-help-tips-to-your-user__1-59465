VERSION 5.00
Begin VB.Form Ftip 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   7290
      Top             =   5175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   7470
      Top             =   4950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1140
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3645
   End
End
Attribute VB_Name = "Ftip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
  
  With Me
     .BackColor = vbInfoBackground
     .Caption = "Help~Tips"
  End With
  
  With Label1
    .AutoSize = True
    .BorderStyle = 0
    .ForeColor = vbInfoText
  End With
  
End Sub
 
Public Sub DisplayHelpTip(stext As String)
  
  ' format the label with the text and appropriate carriage returns
  Label1.Caption = string_MaxLen(stext, 50)
  
  ' since the labels [autosize=True]
  ' this code will resize the form to fit the label
  ' thus we end up with a perfect sized form every time
  With Me
    .Width = Label1.Width + 200
    .Height = (Label1.Height + titlebarHeight + 100)
  End With
  
End Sub

'-----------------------------------------------------------------
' purpose of this function is to take a string and format it
' with carriage returns so that no line exceeds the len of [maxLen]
' the return is that string
'-----------------------------------------------------------------
Private Function string_MaxLen(stext As String, maxLen As Long) As String
  
  ' error prevention
  If Len(Trim$(stext)) = 0 Then Exit Function
  
  ' break apart the string word by word
  Dim sparts() As String, stemp As String
  Dim cr As String, thispart As String
  Dim lcnt As Long
  
  sparts = Split(stext, " ")
  ' create the string with the appropriate cr so
  ' the len of any single line is never longer than [maxLen]
  For lcnt = 0 To UBound(sparts)
    thispart = sparts(lcnt)
    ' if len of stemp and this word is less than [maxLen]
    ' then add this word to stemp
    If Len(stemp & thispart) <= maxLen Then
       stemp = (stemp & thispart & " ")
    Else
       ' otherwise, add cr(maybe) and stemp goes to next line
       string_MaxLen = (string_MaxLen & _
            carriageReturnVal(string_MaxLen) & stemp & " ")
       stemp = thispart & " "
    End If
  Next lcnt
  
  string_MaxLen = (string_MaxLen & carriageReturnVal(string_MaxLen) & stemp)
  
End Function

' the purpose of this function is to return either a carriage
' return (if [strToCheck] isnt empty, otherwise we end up with
' a blank line that is merely composed of carriage return)
' or just and empty string
Private Function carriageReturnVal(strToCheck As String) As String

   If Trim$(Len(strToCheck)) > 0 Then
      carriageReturnVal = vbCrLf
   Else
      carriageReturnVal = ""
   End If
      
End Function

Private Function titlebarHeight()

  titlebarHeight = (Me.Height - Me.ScaleHeight)
  
End Function
 
