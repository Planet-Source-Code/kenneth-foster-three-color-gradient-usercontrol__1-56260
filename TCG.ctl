VERSION 5.00
Begin VB.UserControl TCG 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   990
      Left            =   30
      TabIndex        =   0
      Top             =   1365
      Width           =   4755
   End
End
Attribute VB_Name = "TCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'*         Three Color Gradient
'*           By Ken Foster
'*           Version 1.0.0
'*             Sept 2004
'*   freeware-- use as you please -- enjoy
'*   You can even claim it as your own
'*******************************************************
'  this code based on MasterGradient by Joshua Foster
'=======================================================
Option Explicit

'Constants
Const m_def_tr = 255
Const m_def_tg = 0
Const m_def_tb = 0
Const m_def_mr = 100
Const m_def_mg = 0
Const m_def_mb = 100
Const m_def_br = 0
Const m_def_bg = 0
Const m_def_bb = 255
Const m_def_Caption = ""

'Declares
Dim m_tr As Double
Dim m_tg As Double
Dim m_tb As Double
Dim m_mr As Double
Dim m_mg As Double
Dim m_mb As Double
Dim m_br As Double
Dim m_bg As Double
Dim m_bb As Double
Dim m_Caption As String

Public Sub DrawGrad(tpR, tpG, tpB, mdR, mdG, mdB, btR, btG, btB)
   
   Dim R1 As Double
   Dim G1 As Double
   Dim B1 As Double
   Dim R2 As Double
   Dim G2 As Double
   Dim B2 As Double
   Dim R0 As Double
   Dim G0 As Double
   Dim B0 As Double
   Dim i As Integer
   With UserControl
      .AutoRedraw = True
      .ScaleMode = 0
      .ScaleHeight = 500
      .ScaleWidth = 1
   End With
   
   R0 = tpR
   G0 = tpG
   B0 = tpB
   
   R1 = (mdR - R0) / 250
   G1 = (mdG - G0) / 250
   B1 = (mdB - B0) / 250
   R2 = (btR - mdR) / 250
   G2 = (btG - mdG) / 250
   B2 = (btB - mdB) / 250
   
   For i = 0 To 500
      UserControl.Line (0, i)-(1, i), RGB(R0 * 2.55, G0 * 2.55, B0 * 2.55)
      If i < 250 Then
         R0 = R0 + R1
         G0 = G0 + G1
         B0 = B0 + B1
      Else
         R0 = R0 + R2
         G0 = G0 + G2
         B0 = B0 + B2
      End If
   Next i
   UserControl.ScaleMode = 1
End Sub

Private Sub UserControl_Initialize()
   ctopr = m_tr
   ctopg = m_tg
   ctopb = m_tb
   cmidr = m_mr
   cmidg = m_mg
   cmidb = m_mb
   cbotr = m_br
   cbotg = m_bg
   cbotb = m_bb
   Caption = m_Caption
   Label1.Caption = Caption
   UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
   m_tr = m_def_tr
   m_tg = m_def_tg
   m_tb = m_def_tb
   m_mr = m_def_mr
   m_mg = m_def_mg
   m_mb = m_def_mb
   m_br = m_def_br
   m_bg = m_def_bg
   m_bb = m_def_bb
   m_Caption = m_def_Caption
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_tr = PropBag.ReadProperty("ctopr", m_def_tr)
   m_tg = PropBag.ReadProperty("ctopg", m_def_tg)
   m_tb = PropBag.ReadProperty("ctopb", m_def_tb)
   m_mr = PropBag.ReadProperty("cmidr", m_def_mr)
   m_mg = PropBag.ReadProperty("cmidg", m_def_mg)
   m_mb = PropBag.ReadProperty("cmidb", m_def_mb)
   m_br = PropBag.ReadProperty("cbotr", m_def_br)
   m_bg = PropBag.ReadProperty("cbotg", m_def_bg)
   m_bb = PropBag.ReadProperty("cbotb", m_def_bb)
   m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
   Set Font = PropBag.ReadProperty("Font", Ambient.Font)
   Label1.Fontsize = PropBag.ReadProperty("Fontsize", 10)
   Label1.Fontbold = PropBag.ReadProperty("Fontbold", False)
   Label1.Caption = Caption
   UserControl_Resize
End Sub

Private Sub UserControl_Resize()
   DoEvents
   Label1.Top = (UserControl.Height \ 2) - (Label1.Height \ 6)
   Label1.Width = UserControl.Width
   DrawGrad ctopr, ctopg, ctopb, cmidr, cmidg, cmidb, cbotr, cbotg, cbotb
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "ctopr", m_tr, m_def_tr
   PropBag.WriteProperty "ctopg", m_tg, m_def_tg
   PropBag.WriteProperty "ctopb", m_tb, m_def_tb
   PropBag.WriteProperty "cmidr", m_mr, m_def_mr
   PropBag.WriteProperty "cmidg", m_mg, m_def_mg
   PropBag.WriteProperty "cmidb", m_mb, m_def_mb
   PropBag.WriteProperty "cbotr", m_br, m_def_br
   PropBag.WriteProperty "cbotg", m_bg, m_def_bg
   PropBag.WriteProperty "cbotb", m_bb, m_def_bb
   PropBag.WriteProperty "Caption", m_Caption, m_def_Caption
   PropBag.WriteProperty "Font", Font, Ambient.Font
   PropBag.WriteProperty "Fontsize", Label1.Fontsize, 10
   PropBag.WriteProperty "Fontbold", Label1.Fontbold, False
   
End Sub
Public Property Get ctopr() As Double
   ctopr = m_tr
End Property
Public Property Let ctopr(ByVal new_tr As Double)
   m_tr = new_tr
   UserControl_Resize
   PropertyChanged "ctopr"
End Property

Public Property Get ctopg() As Double
   ctopg = m_tg
End Property
Public Property Let ctopg(ByVal new_tg As Double)
   m_tg = new_tg
   UserControl_Resize
   PropertyChanged "ctopg"
End Property
Public Property Get ctopb() As Double
   ctopb = m_tb
End Property
Public Property Let ctopb(ByVal new_tb As Double)
   m_tb = new_tb
   UserControl_Resize
   PropertyChanged "ctopb"
End Property
Public Property Get cmidr() As Double
   cmidr = m_mr
End Property
Public Property Let cmidr(ByVal new_mr As Double)
   m_mr = new_mr
   UserControl_Resize
   PropertyChanged "cmidr"
End Property
Public Property Get cmidg() As Double
   cmidg = m_mg
End Property
Public Property Let cmidg(ByVal new_mg As Double)
   m_mg = new_mg
   UserControl_Resize
   PropertyChanged "cmidg"
End Property
Public Property Get cmidb() As Double
   cmidb = m_mb
End Property
Public Property Let cmidb(ByVal new_mb As Double)
   m_mb = new_mb
   UserControl_Resize
   PropertyChanged "cmidb"
End Property
Public Property Get cbotr() As Double
   cbotr = m_br
End Property
Public Property Let cbotr(ByVal new_br As Double)
   m_br = new_br
   UserControl_Resize
   PropertyChanged "cbotr"
End Property
Public Property Get cbotg() As Double
   cbotg = m_bg
End Property
Public Property Let cbotg(ByVal new_bg As Double)
   m_bg = new_bg
   UserControl_Resize
   PropertyChanged "cbotg"
End Property
Public Property Get cbotb() As Double
   cbotb = m_bb
End Property
Public Property Let cbotb(ByVal new_bb As Double)
   m_bb = new_bb
   UserControl_Resize
   PropertyChanged "cbotb"
End Property
Public Property Get Caption() As String
   Caption = m_Caption
End Property
Public Property Let Caption(ByVal new_Caption As String)
   m_Caption = new_Caption
   Label1.Caption = Caption
   UserControl_Resize
   PropertyChanged "Caption"
End Property
Public Property Get Font() As Font
   Set Font = Label1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
   Set Label1.Font = New_Font
   PropertyChanged "Font"
End Property
Public Property Get Fontsize() As Single
   Fontsize = Label1.Fontsize
End Property
Public Property Let Fontsize(ByVal New_Fontsize As Single)
   Label1.Fontsize = New_Fontsize
   PropertyChanged "Fontsize"
End Property
Public Property Get Fontbold() As Boolean
   Fontbold = Label1.Fontbold
End Property
Public Property Let Fontbold(ByVal New_Fontbold As Boolean)
   Label1.Fontbold = New_Fontbold
   PropertyChanged "Fontbold"
End Property
