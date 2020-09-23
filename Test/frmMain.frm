VERSION 5.00
Object = "{E8599063-D892-47F9-BA41-EDD5EAC4446F}#9.0#0"; "acSuperViewPort2.ocx"
Begin VB.Form frmMain 
   Caption         =   "SuperViewPort 2.0 - Test"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin acSuperViewPort.SuperViewPort SuperViewPort1 
      Align           =   1  'Align Top
      Height          =   4995
      Left            =   0
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8811
      BorderStyle     =   0
      ScrollBars      =   3
      ViewPortWidth   =   300
      ViewPortHeight  =   250
      GradientEnabled =   -1  'True
      MouseTrack      =   -1  'True
      Begin acSuperViewPort.SuperViewPort SuperViewPort2 
         Height          =   1905
         Left            =   1230
         Top             =   420
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3360
         ForeColor       =   8438015
         BackColor       =   8388608
         GradientEnabled =   -1  'True
         GradientAngle   =   135
         Begin VB.CommandButton Command1 
            Caption         =   "Dummy 1"
            Height          =   435
            Left            =   870
            TabIndex        =   0
            Top             =   780
            Width           =   855
         End
      End
      Begin acSuperViewPort.SuperViewPort SuperViewPort3 
         Height          =   1905
         Left            =   3810
         Top             =   420
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3360
         ForeColor       =   8438015
         BackColor       =   8388608
         ScrollBars      =   2
         GradientEnabled =   -1  'True
         GradientAngle   =   45
         Begin VB.CommandButton Command2 
            Caption         =   "Dummy 2"
            Height          =   435
            Left            =   780
            TabIndex        =   1
            Top             =   780
            Width           =   855
         End
      End
      Begin acSuperViewPort.SuperViewPort SuperViewPort4 
         Height          =   1905
         Left            =   1230
         Top             =   2340
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3360
         ForeColor       =   8438015
         BackColor       =   8388608
         ScrollBars      =   1
         GradientEnabled =   -1  'True
         GradientAngle   =   225
         Begin VB.CommandButton Command4 
            Caption         =   "Dummy 3"
            Height          =   435
            Left            =   870
            TabIndex        =   3
            Top             =   600
            Width           =   855
         End
      End
      Begin acSuperViewPort.SuperViewPort SuperViewPort5 
         Height          =   1905
         Left            =   3810
         Top             =   2340
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   3360
         ForeColor       =   8438015
         BackColor       =   8388608
         ScrollBars      =   3
         GradientEnabled =   -1  'True
         GradientAngle   =   315
         Begin VB.CommandButton Command3 
            Caption         =   "Dummy 4"
            Height          =   435
            Left            =   780
            TabIndex        =   2
            Top             =   600
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Sub Form_Resize()
    SuperViewPort1.Height = ScaleHeight
End Sub
