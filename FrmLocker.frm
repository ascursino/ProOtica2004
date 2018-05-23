VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{9C152BB9-D77B-11D7-A6B5-00D009F8C11B}#3.0#0"; "shlock.ocx"
Begin VB.Form FrmLocker 
   Caption         =   "Pró Ótica 2004 - Registro"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   Icon            =   "FrmLocker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   15
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   5055
   End
   Begin VB.CommandButton CmdContinuar 
      Caption         =   "Continuar"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      ToolTipText     =   "Continuar utilizando o software sem registrar"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   840
      Picture         =   "FrmLocker.frx":0CCA
      ScaleHeight     =   900
      ScaleWidth      =   3750
      TabIndex        =   3
      ToolTipText     =   "Pró Ótica 2004 é um software desenvolvido por Infodigital Soluções em Desenvolvimento"
      Top             =   120
      Width           =   3750
   End
   Begin VB.CommandButton CmdRegistrar 
      Caption         =   "Registrar"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Registrar software"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox txtChave 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Chave gerada pela Infodigital para efetuar o registro do software"
      Top             =   3600
      Width           =   2715
   End
   Begin VB.TextBox txtSerial 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Text            =   "txtSerial"
      ToolTipText     =   "Serial gerado para registro do software"
      Top             =   2520
      Width           =   2600
   End
   Begin SHLock.SHLocker SHLocker1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   1032
      _ExtentY        =   979
      SenhaProg       =   "ProOtica2004"
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FrmLocker.frx":4C1C
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "FrmLocker.frx":4E50
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "FrmLocker.frx":4EC6
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   615
      Left            =   360
      OleObjectBlob   =   "FrmLocker.frx":4F38
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   0
      OleObjectBlob   =   "FrmLocker.frx":4FE2
      TabIndex        =   7
      Top             =   1200
      Width           =   5295
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblDiasQueUsei 
      Height          =   255
      Left            =   2160
      OleObjectBlob   =   "FrmLocker.frx":5062
      TabIndex        =   9
      ToolTipText     =   "Dias utilizados do software"
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "FrmLocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdContinuar_Click()
    Unload Me
    MDIPrincipal.Show
End Sub

Private Sub CmdRegistrar_Click()
    'se retornar true é porque a chave inserida está correta.
    If SHLocker1.Liberar(txtChave.Text) Then
        Unload Me
        MDIPrincipal.Show
    Else
        VGStrBox = MsgBox("Chave inválida!", vbCritical, "Pró Ótica 2004")
    End If
End Sub

'Private Sub Form_Resize()
'  FrmLocker.Left = (MDIPrincipal.Width / 2) - (FrmLocker.Width / 2)
'  FrmLocker.Top = (MDIPrincipal.Height / 3) - (FrmLocker.Height / 3)
'End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbNormal
    
    Left = 4815
    Top = 2865

    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    If VGStrLocker = "sim" Then
        CmdContinuar.Enabled = False
    Else
        CmdContinuar.Enabled = True
    End If

    'quantos dias a pessoa utilizou
    LblDiasQueUsei.Caption = CStr(SHLocker1.DiasQueUsei)
    
    'número de serial a ser informado para o sistema SHUnloker gerar a chave para liberar
    txtSerial.Text = SHLocker1.Serial
    
    'esta senha serve para gerar diferentes seriais e
    'permitir que um micro tenha vários programas protegidos pelo SHLocker
    ''txtSenhaProg.Text = SHLocker1.SenhaProg

End Sub

Private Sub Form_Terminate()
    Unload MDIPrincipal
    Screen.MousePointer = vbHourglass
End Sub
