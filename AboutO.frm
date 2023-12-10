VERSION 5.00
Begin VB.Form AboutO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Video Club"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AboutO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "WEB:    http://users.yubc.net/~hronos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "AboutO.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   2400
      Picture         =   "AboutO.frx":0316
      Top             =   120
      Width           =   3060
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail : hronos@yubc.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      MouseIcon       =   "AboutO.frx":2BAC
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2001 by Aleksandar Todorovic"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "VERZIJA 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VIDEO CLUB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "AboutO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
Set AboutO = Nothing
End Sub

Private Sub Label4_Click()
Shell "C:\Program Files\Outlook Express\MSIMN.EXE", vbNormalFocus


End Sub
