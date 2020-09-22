VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Effects Example"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExample2 
      Caption         =   "Example 2"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Tag             =   "/3DUP/"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdExample1 
      Caption         =   "Example 1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Tag             =   "/3D/"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtExample2 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "/3DUP/"
      Text            =   "Text Example 2"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtExample1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "/3D/"
      Text            =   " Text Example 1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblExample2 
      AutoSize        =   -1  'True
      Caption         =   "Example Label 2"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Tag             =   "/3DUP/"
      Top             =   600
      Width           =   1170
   End
   Begin VB.Label lblExample1 
      AutoSize        =   -1  'True
      Caption         =   "Example Label 1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Tag             =   "/3D/"
      Top             =   120
      Width           =   1170
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'YOU MUST MAKE SURE THAT THE AUTOREDRAW PROPERTY OF THE FORM IS SET TO TRUE FOR THIS TO WORK
'YOU MUST ALSO MAKE SURE THE SCALEMODE IS SET TO 1 (TWIP), OTHERWISE YOU WILL HAVE TO
'   ADJUST THE CODE TO FIND THE PROPER RATION OF TWIPS TO PIXELS TO USE PIXELS

Private Sub Form_Load()
'Variable for looping through controls
Dim a As Integer

'Loop through the controls in the form
For a = 0 To Me.Controls.Count - 1
    'See if the control has a 3D tag
   If InStr(UCase$(Me.Controls(a).Tag), "/3D/") Then
      'Control has an "inset" tag so draw the inset border around it
      Make3D Me, Me.Controls(a), BORDER_INSET
   ElseIf InStr(UCase$(Me.Controls(a).Tag), "/3DUP/") Then
      'Ctonrol has a "raised" tag so draw the raised border around it
      Make3D Me, Me.Controls(a), BORDER_RAISED
   End If
Next a
End Sub
