VERSION 5.00
Begin VB.UserControl splitter 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   165
   ScaleHeight     =   1665
   ScaleWidth      =   165
   Begin VB.PictureBox picCursor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4095
      ScaleHeight     =   330
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   2565
      Width           =   375
   End
End
Attribute VB_Name = "splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function InflateRect Lib "user32" (lpRect As Rect, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function UnionRect Lib "user32" (lpDestRect As Rect, lpSrc1Rect As Rect, lpSrc2Rect As Rect) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Type Rect
   Left As Long
   Top  As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
   x As Long
   y As Long
End Type


Enum enDecoration
    decRaised = 0
    decFlat = 1
End Enum

Enum enSplitterDir
    sdVertical = 0
    sdHorizontal = 1
End Enum

'Default Property Values:
Const m_def_SplitterActivateColor = &HFF8080
Const m_def_Decoration = 0
Const m_def_SplitterDirection = 0

'Property Variables:
Dim m_SplitterActivateColor As OLE_COLOR
Dim m_Control_Top_Or_Left As Object
Dim m_Control_Bottom_Or_Right As Object
Dim m_Decoration As enDecoration
Dim m_SplitterDirection As enSplitterDir


 
Dim m_moving       As Byte
Dim m_startx       As Long
Dim m_starty       As Long
Dim m_starttop     As Long
Dim m_startleft    As Long
Dim MouseDown_Rect As Rect

Private Const CTRL_BUFFER = 250
 
Event BeforeScroll()
Attribute BeforeScroll.VB_Description = "Event raised when the mouse is pressed down on the splitter bar"
Event Scrolling(direction As Long)
Attribute Scrolling.VB_Description = "Event raised when the splitter is being moved"
Event AfterScroll()
Attribute AfterScroll.VB_Description = "Event raised when the mouse raised off of the splitter bar after the splitter bar has been moved"
Event Error(sErrDescription As String)
Attribute Error.VB_Description = "Event raised in the event of an important control error"

 
' ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___  ___
'|___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___|
'|___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___||___|
'    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '
'    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '    '









 
'     .    .    .    .    .    .      .     .    .    .    .
'_   _  _ _' ___  _ _  ___  ___  _ __  _____  _ _  ___ _
' | | |/ __|/ _ \| '_\/ __|/ _ \| '_ ` _   _|| '_\/ _ \| |
' |_| |\__ \  __/| |   (__  (_) | | | | | |  | |   (_) | |_
'\__,_||___/\___||_|  \___|\___/|_| |_| |_|  |_|  \___/|___|
'    .     . _   .    .
' _ _'_   _ | |__  _ _'
'/ __| | | ||  _ \/ __|
'\__ \ |_| || |_) \__ \
'|___/\__,_||_.__/|___/


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pt As POINTAPI
On Error GoTo local_error:

  m_moving = 1 '-- tag for the mouse_move
  '-- defines the boundries of the scroll allowed
  Call AdjustRect
  '-- store the mousedown point and the current postition
  '   of this control so we can move in {see usercontrol_mousemove}
  GetCursorPos pt
  m_startx = pt.x
  m_starty = pt.y
  m_starttop = Extender.Top
  m_startleft = Extender.Left
  '-- draw highlight box around the control
  UserControl.Line (0, 0)-(Width - 20, Height - 20), m_SplitterActivateColor, B
  '-- event raised when splitter is mouseed down on
  RaiseEvent BeforeScroll
  
Exit Sub
local_error:
   If Err.Number <> 0 Then
       Debug.Print "splitter.ocx.UserControl_MouseDown: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pt           As POINTAPI
Dim sliderL      As Long
Dim sliderT      As Long
Dim sliderWid    As Long
Dim sliderHei    As Long
On Error GoTo local_error:
 
 If m_moving = 1 Then '-- tag set in the mouse_down
    '-- confine the cursor to within the allowed scrollable area
    '   which is within 250 twips of the outer edges of
    '   [Control_Top_Or_Left] and [Control_Bottom_Or_Right]
   Call ClipCursor(MouseDown_Rect)
   '-- we move the splitter to where the mouse is
   GetCursorPos pt
  
   '-- move the splitter bar
   If m_SplitterDirection = sdHorizontal Then
       Extender.Top = m_starttop + ((pt.y - m_starty) * Screen.TwipsPerPixelY)

       If pt.y > m_starty Then
          RaiseEvent Scrolling(1)
       ElseIf pt.y < m_starty Then
          RaiseEvent Scrolling(0)
       End If
    Else
       Extender.Left = m_startleft + ((pt.x - m_startx) * Screen.TwipsPerPixelX)
       
       If pt.x > m_startx Then
          RaiseEvent Scrolling(1)
       ElseIf pt.x < m_startx Then
          RaiseEvent Scrolling(0)
       End If
   End If
   
 
 End If
 
Exit Sub
local_error:
   If Err.Number <> 0 Then
      If Err.Number = 91 Then
         If Erl() = 2 Then
            RaiseEvent Error("A container object has not been assigned (set)" & _
                             "for property [SplitterContainerControl]")
         End If
      End If
      
      Debug.Print "splitter.ocx." & "UserControl_MouseMove: " & _
                    Err.Number & "." & Err.Description
   End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim O1          As Object
Dim O2          As Object
Dim scrX        As Long
Dim scrY        As Long
Dim Rl          As Long
Dim Rt          As Long
Dim Rr          As Long
Dim Rb          As Long
On Error GoTo local_error:
  
  '-- free up the cursors boundried
  ClipCursor ByVal 0&
  Cls
  '-- clear the hilight effect of mouse_down
  Call UserControl_Paint
  
  scrX = Screen.TwipsPerPixelX
  scrY = Screen.TwipsPerPixelY
  Set O1 = m_Control_Top_Or_Left
  Set O2 = m_Control_Bottom_Or_Right
  
 
  If m_SplitterDirection = sdVertical Then
     '-- set control1 width to where the splitter is
     O1.Width = (Extender.Left - O1.Left)
     '-- coods for where the bottom or right control will be moved to
     Rl = ((Extender.Left + 50) / scrX)
     Rt = (O2.Top / scrY)
     Rr = (((O2.Left + O2.Width) - (Extender.Left + Extender.Width)) / scrX)
     Rb = (O2.Height / scrY)
  Else
      '-- set control1 height to where the splitter is
     O1.Height = (Extender.Top - O1.Top)
     '-- coods for where the bottom or right control will be moved to
     Rl = (O2.Left / scrX)
     Rt = ((Extender.Top + 50) / scrY)
     Rr = (O2.Width / scrX)
     Rb = (((O2.Top + O2.Height) - (Extender.Top + Extender.Height)) / Screen.TwipsPerPixelY)
  End If
 
  '-- resize control #2
  '   we need to use movewindow and not ctl.move because,
  '   for whatever reason..ctl.move is not very precise
  '   and the right/bottom control doesnt maintain proper
  '   positioning
  MoveWindow O2.hWnd, Rl, Rt, Rr, Rb, True
   
  If m_moving = 1 Then
     m_moving = 0
     RaiseEvent AfterScroll
  End If
  
  Set O1 = Nothing
  Set O2 = Nothing
  
Exit Sub
local_error:
   If Err.Number <> 0 Then
       Debug.Print "splitter.ocx.UserControl_MouseUp: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Sub

Private Sub UserControl_Paint()
  
  If m_Decoration = decRaised Then
     UserControl.Line (0, 0)-(Width, Height), RGB(255, 255, 255), B
     UserControl.Line (-50, -50)-(Width - 20, Height - 20), RGB(180, 180, 190), B
  End If
End Sub
Private Sub Usercontrol_Resize()
On Error GoTo local_error:

  '-- maintain the splitters thickness
  If m_SplitterDirection = sdHorizontal Then
     Height = 50
  Else
     Width = 50
  End If
  '-- if the decoration is raised then we need to
  '   paint hilight and shadow effects
  If m_Decoration = decRaised Then
     Cls
     Call UserControl_Paint
  End If

Exit Sub
local_error:
   If Err.Number <> 0 Then
       Debug.Print "Form1.AddChildNodes: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Sub

Private Sub UserControl_Show()
On Error GoTo local_error:

'-- move control1 and 2 so that they line up right next to the splitter
If Ambient.UserMode Then
  '-- set and position control1 and control2 to where the splitter is
  UserControl_MouseUp 0, 0, 0, 0
End If

local_error:
   If Err.Number <> 0 Then
       Debug.Print "Form1.AddChildNodes: " & Err.Number & "." & Err.Description
       Err.Clear
       Resume Next
   End If
End Sub














' _ __     . _      .     .     .    .  '    .     . _   .    .
'| '_ \ _ _ (_)_   __ __ _ _____  ___   ' _ _'_   _ | |__  _ _'
'| |_) | '_\| | \ / // _` |_   _|/ _ \  '/ __| | | ||  _ \/ __|
'| .__/| |  | |\ V /  (_| | | |    __/  '\__ \ |_| || |_) \__ \
'|_|   |_|  |_| \_/  \__,_| |_|  \___|  '|___/\__,_||_.__/|___/

 
Private Sub AdjustRect()
Dim Rect1        As Rect
Dim Rect2        As Rect

 '-- get the first controls rect and the second controls rect
 GetWindowRect m_Control_Top_Or_Left.hWnd, Rect1
 GetWindowRect m_Control_Bottom_Or_Right.hWnd, Rect2
 '-- join them...this ends up being our rect area splitter can move in
 UnionRect MouseDown_Rect, Rect1, Rect2
 '-- create a small buffer that will be the min height or width of the control
 If m_SplitterDirection = sdHorizontal Then
    InflateRect MouseDown_Rect, 0, -(CTRL_BUFFER / Screen.TwipsPerPixelY)
 Else
    InflateRect MouseDown_Rect, -(CTRL_BUFFER / Screen.TwipsPerPixelX), 0
 End If
 
End Sub
















' _ __     .    . _ __     .    .     . _     .    .
'| '_ \ _ _  ___ | '_ \ ___  _ _ _____ (_) ___  _ _'
'| |_) | '_\/ _ \| |_) / _ \| '_\_   _|| |/ _ \/ __|
'| .__/| |   (_) | .__/  __/| |   | |  | |  __/\__ \
'|_|   |_|  \___/|_|   \___||_|   |_|  |_|\___||___/


'Backcolor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Color of the splitter bar"
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call UserControl_Paint
End Property
'Control_Top_Or_Left
Public Property Get Control_Top_Or_Left() As Object
Attribute Control_Top_Or_Left.VB_Description = "The name of the control the is to the left of the splitter bar (if [SplitterDirection] = Vertical)  or to the top of the splitter bar (if [SplitterDirection]=Horizontal)"
    Set Control_Top_Or_Left = m_Control_Top_Or_Left
End Property
Public Property Set Control_Top_Or_Left(ByVal New_Control_Top_Or_Left As Object)
    Set m_Control_Top_Or_Left = New_Control_Top_Or_Left
    PropertyChanged "Control_Top_Or_Left"
End Property
'Control_Bottom_Or_Right
Public Property Get Control_Bottom_Or_Right() As Object
Attribute Control_Bottom_Or_Right.VB_Description = "The name of the control the is to the right of the splitter bar (if [SplitterDirection] = Vertical)  or to the bottom of the splitter bar (if [SplitterDirection]=Horizontal)"
    Set Control_Bottom_Or_Right = m_Control_Bottom_Or_Right
End Property
Public Property Set Control_Bottom_Or_Right(ByVal New_Control_Bottom_Or_Right As Object)
    Set m_Control_Bottom_Or_Right = New_Control_Bottom_Or_Right
    PropertyChanged "Control_Bottom_Or_Right"
End Property
'CustomMousePointer
Public Property Get CustomMousePointer() As Picture
Attribute CustomMousePointer.VB_Description = "If this property is not specified the the default cursor for EW (if [SplitterDirection] = Vertical  or NS (if [SplitterDirection]= Horizontal.  If specifed, this is the mouse icon when the mouse hovers over the splitter"
    Set CustomMousePointer = picCursor.Picture
End Property
Public Property Set CustomMousePointer(ByVal New_CustomMousePointer As Picture)
    Set picCursor.Picture = New_CustomMousePointer
    PropertyChanged "CustomMousePointer"
End Property
'Decoration
Public Property Get Decoration() As enDecoration
Attribute Decoration.VB_Description = "The appearance of the splitter being either  [Flat]  or  [Raised]"
    Decoration = m_Decoration
End Property
Public Property Let Decoration(ByVal New_Decoration As enDecoration)
    m_Decoration = New_Decoration
    PropertyChanged "Decoration"
    Cls
    Call UserControl_Paint
End Property
'SplitterActivateColor
Public Property Get SplitterActivateColor() As OLE_COLOR
Attribute SplitterActivateColor.VB_Description = "When the Splitter control is being moved it displays a highlight color providing visual feedback that its activated"
    SplitterActivateColor = m_SplitterActivateColor
End Property
Public Property Let SplitterActivateColor(ByVal New_SplitterActivateColor As OLE_COLOR)
    m_SplitterActivateColor = New_SplitterActivateColor
    PropertyChanged "SplitterActivateColor"
End Property
Public Property Get SplitterDirection() As enSplitterDir
Attribute SplitterDirection.VB_Description = "The orientation of the the splitter..either [Vertical] or [Horizontal]"
    SplitterDirection = m_SplitterDirection
End Property
Public Property Let SplitterDirection(ByVal New_SplitterDirection As enSplitterDir)
    m_SplitterDirection = New_SplitterDirection
    PropertyChanged "SplitterDirection"
    
    If m_SplitterDirection = sdVertical Then
        Width = 50:  Height = 2500
    Else
        Width = 2500: Height = 50
    End If
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SplitterDirection = m_def_SplitterDirection
    m_Decoration = m_def_Decoration
    m_SplitterActivateColor = m_def_SplitterActivateColor
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_SplitterDirection = PropBag.ReadProperty("SplitterDirection", m_def_SplitterDirection)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Decoration = PropBag.ReadProperty("Decoration", m_def_Decoration)
    Set m_Control_Top_Or_Left = PropBag.ReadProperty("Control_Top_Or_Left", Nothing)
    Set m_Control_Bottom_Or_Right = PropBag.ReadProperty("Control_Bottom_Or_Right", Nothing)
    m_SplitterActivateColor = PropBag.ReadProperty("SplitterActivateColor", m_def_SplitterActivateColor)
    Set picCursor.Picture = PropBag.ReadProperty("CustomMousePointer", Nothing)
     
    If picCursor.Picture = 0 Then
      If m_SplitterDirection = sdVertical Then
           MousePointer = 9 '-- north south
      Else
           MousePointer = 7 '-- north south
      End If
    Else
      MousePointer = 99
      MouseIcon = picCursor.Picture
    End If
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SplitterDirection", m_SplitterDirection, m_def_SplitterDirection)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Decoration", m_Decoration, m_def_Decoration)
    Call PropBag.WriteProperty("Control_Top_Or_Left", m_Control_Top_Or_Left, Nothing)
    Call PropBag.WriteProperty("Control_Bottom_Or_Right", m_Control_Bottom_Or_Right, Nothing)
    Call PropBag.WriteProperty("SplitterActivateColor", m_SplitterActivateColor, m_def_SplitterActivateColor)
    Call PropBag.WriteProperty("CustomMousePointer", picCursor.Picture, Nothing)
End Sub

  

