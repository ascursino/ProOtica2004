VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o Sistema"
   ClientHeight    =   6000
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "frmSistema.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6195.219
   ScaleMode       =   0  'User
   ScaleWidth      =   7293.953
   Begin VB.Frame Frame1 
      Height          =   5880
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   6510
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         Picture         =   "frmSistema.frx":0CCA
         ScaleHeight     =   975
         ScaleWidth      =   3855
         TabIndex        =   10
         Top             =   240
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   3960
         OleObjectBlob   =   "frmSistema.frx":4C1C
         Top             =   120
      End
      Begin VB.Frame Frame3 
         Height          =   60
         Left            =   90
         TabIndex        =   2
         Top             =   4560
         Width           =   6360
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   150
         TabIndex        =   1
         Top             =   1200
         Width           =   6255
      End
      Begin RichTextLib.RichTextBox rtbLicenca 
         Height          =   2325
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Contrato de licença de software"
         Top             =   1440
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4101
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         FileName        =   "C:\Infodigital\Sistemas\ProOtica2004\_licenca.dll"
         TextRTF         =   $"frmSistema.frx":4E50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "frmSistema.frx":A89D
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "frmSistema.frx":A903
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "frmSistema.frx":A96F
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblDtCriacao 
         Height          =   255
         Left            =   5400
         OleObjectBlob   =   "frmSistema.frx":A9E3
         TabIndex        =   7
         ToolTipText     =   "data de criação"
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUltimaAtualizacao 
         Height          =   255
         Left            =   5400
         OleObjectBlob   =   "frmSistema.frx":AA4F
         TabIndex        =   8
         ToolTipText     =   "data da última atualização"
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblVersao 
         Height          =   255
         Left            =   5400
         OleObjectBlob   =   "frmSistema.frx":AABB
         TabIndex        =   9
         ToolTipText     =   "versão"
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "frmSistema.frx":AB1D
         TabIndex        =   11
         ToolTipText     =   "Informações sobre compra, suporte e atualizações do produto"
         Top             =   3960
         Width           =   6255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   975
         Left            =   120
         OleObjectBlob   =   "frmSistema.frx":AC6B
         TabIndex        =   12
         ToolTipText     =   "Informações de copyright"
         Top             =   4800
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    MDIPrincipal.WindowState = 2

    LblVersao.Caption = App.Major & "." & App.Minor & "." & App.Revision
    
    rtbLicenca.FileName = App.Path & "\_licenca.dll"
    
    Height = 6480
    Width = 6795
    Left = 3690
    Top = 1260
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Resize()
  frmSistema.Left = (MDIPrincipal.Width / 2) - (frmSistema.Width / 2)
  frmSistema.Top = (MDIPrincipal.Height / 3) - (frmSistema.Height / 3)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

