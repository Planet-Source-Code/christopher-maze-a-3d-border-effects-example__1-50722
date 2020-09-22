Attribute VB_Name = "mod3D"
Option Explicit

'3D effect constants
Global Const BORDER_INSET = 0
Global Const BORDER_RAISED = 1

'Color Constants
Global Const DARK_GRAY = &H808080
Global Const WHITE = &HFFFFFF

Public Sub Make3D(pic As Form, ctl As Control, ByVal borderstyle As Integer)
Dim AdjustX As Integer, AdjustY As Integer
Dim RightSide As Single
Dim BW As Integer, BorderWidth As Integer
Dim LeftTopColor As Long, RightBottomColor As Long
Dim i As Integer

'   If ctl.Visible = False Then Exit Sub
   
   'Variable for distance to move each line if the border width is greater than 1
   AdjustX = Screen.TwipsPerPixelX
   AdjustY = Screen.TwipsPerPixelY
   
   'The width of the 3D border around the objects
   BorderWidth = 2
   
   'Check the style of border to add (inset or raised) to set the appropriate corner colors
   Select Case borderstyle
   Case BORDER_INSET: 'Inset
      LeftTopColor = DARK_GRAY
      RightBottomColor = WHITE
   Case BORDER_RAISED: 'Raised
      LeftTopColor = WHITE
      RightBottomColor = DARK_GRAY
   End Select
   
   'Draw the border around the control
   For BW = 1 To BorderWidth
      'top
      pic.CurrentX = ctl.Left - (AdjustX * BW)
      pic.CurrentY = ctl.Top - (AdjustY * BW)
      pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top - (AdjustY * BW)), LeftTopColor
      pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
      pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
      pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top - (AdjustY * BW)), LeftTopColor
   Next
End Sub
