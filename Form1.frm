VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Three Color Gradient User Control"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin Project1.TCG TCG1 
      Height          =   2130
      Left            =   15
      TabIndex        =   0
      Top             =   60
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3757
      ctopr           =   155
      cmidr           =   125
      cmidg           =   125
      cmidb           =   125
      cbotb           =   155
      Caption         =   "United States of America"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Fontsize        =   12
      Fontbold        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
'TCG1.Width = Form1.Width
'TCG2.Width = Form1.Width
'TCG1.Height = Form1.Height - 300
'TCG2.Height = Form1.Height \ 2
'TCG2.Top = Form1.Height \ 2 - 30
End Sub

