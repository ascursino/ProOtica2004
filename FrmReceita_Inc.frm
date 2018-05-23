VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmReceita_Inc 
   Caption         =   "Inclus�o de Receita"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmReceita_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   7680
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
      Height          =   6495
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton CmdCalcAdicao 
         Caption         =   "Calcula adi��o"
         Height          =   495
         Left            =   5640
         TabIndex        =   85
         ToolTipText     =   "Calcular adi��o"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox TxtObsRec 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Observa��o sobre a receita"
         Top             =   5640
         Width           =   7095
      End
      Begin VB.ComboBox CboMedico 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Nome do m�dico"
         Top             =   4920
         Width           =   5655
      End
      Begin VB.CommandButton CmdIncluirMed 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   18
         ToolTipText     =   "Adicionar m�dico"
         Top             =   4920
         Width           =   375
      End
      Begin VB.Frame Frame15 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   3720
         TabIndex        =   62
         Top             =   480
         Width           =   3615
         Begin VB.Frame Frame16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   63
            Top             =   120
            Width           =   2895
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":0CCA
               TabIndex        =   64
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   65
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Inc.frx":0D2C
               TabIndex        =   66
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            TabIndex        =   67
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":0D8C
               TabIndex        =   68
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   69
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Inc.frx":0DEC
               TabIndex        =   70
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   71
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtPDEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   6
               ToolTipText     =   "Perto grau esf�rico olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame21 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   72
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtPDCil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   7
               ToolTipText     =   "Perto grau cil�ndrico olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            TabIndex        =   73
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtPDEixo 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   8
               ToolTipText     =   "Perto eixo olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame26 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   74
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtPECil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   10
               ToolTipText     =   "Perto grau cil�ndrico olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            TabIndex        =   75
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtPEEixo 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   11
               ToolTipText     =   "Perto eixo olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   76
            Top             =   1080
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":0E4C
               TabIndex        =   77
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   78
            Top             =   1560
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":0EA8
               TabIndex        =   79
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame25 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   80
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtPEEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   9
               ToolTipText     =   "Perto grau esf�rico olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   3615
         Begin VB.Frame Frame14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   44
            Top             =   120
            Width           =   2895
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":0F04
               TabIndex        =   45
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   46
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Inc.frx":0F66
               TabIndex        =   47
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame31 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            TabIndex        =   48
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":0FC6
               TabIndex        =   49
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   50
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Inc.frx":1026
               TabIndex        =   51
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   52
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtLDEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   0
               ToolTipText     =   "Longe grau esf�rico olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame34 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   53
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtLDCil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   1
               ToolTipText     =   "Longe grau cil�ndrico olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame35 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            TabIndex        =   54
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtLDEixo 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   2
               ToolTipText     =   "Longe eixo olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame36 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   55
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtLECil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   4
               ToolTipText     =   "Longe grau cil�ndrico olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame37 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2520
            TabIndex        =   56
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtLEEixo 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   5
               ToolTipText     =   "Longe eixo olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":1086
               TabIndex        =   58
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame39 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":10E2
               TabIndex        =   60
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame40 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   61
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtLEEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   3
               ToolTipText     =   "Longe grau esf�rico olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame41 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   2655
         Begin VB.Frame Frame45 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   33
            Top             =   120
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Inc.frx":113E
               TabIndex        =   34
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame43 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   31
            Top             =   120
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Inc.frx":119E
               TabIndex        =   32
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame46 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   35
            Top             =   600
            Width           =   975
            Begin VB.TextBox TxtDNPD 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   12
               ToolTipText     =   "DNP olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame47 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   36
            Top             =   600
            Width           =   975
            Begin VB.TextBox TxtAltD 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   13
               ToolTipText     =   "Altura olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame49 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   37
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtAltE 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   15
               ToolTipText     =   "Altura olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame51 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":11FC
               TabIndex        =   39
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame52 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":1258
               TabIndex        =   41
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame53 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   42
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtDNPE 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   14
               ToolTipText     =   "DNP olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   3720
         TabIndex        =   24
         Top             =   2760
         Width           =   1695
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   1455
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":12B4
               TabIndex        =   26
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   1455
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Inc.frx":1318
               TabIndex        =   28
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame48 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   1455
            Begin VB.TextBox TxtAdicAO 
               Height          =   285
               Left            =   360
               MaxLength       =   6
               TabIndex        =   16
               ToolTipText     =   "Adi��o ambos os olhos"
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReceita_Inc.frx":1374
         TabIndex        =   81
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNomeCli 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmReceita_Inc.frx":13DC
         TabIndex        =   82
         Top             =   240
         Width           =   5415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReceita_Inc.frx":1456
         TabIndex        =   83
         Top             =   4920
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReceita_Inc.frx":14BC
         TabIndex        =   84
         Top             =   5400
         Width           =   1215
      End
   End
   Begin VB.Frame FraBotaoCli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   6600
      Width           =   7455
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1200
         OleObjectBlob   =   "FrmReceita_Inc.frx":152A
         Top             =   120
      End
      Begin VB.CommandButton CmdFechar 
         Caption         =   "&Fechar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   21
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         ToolTipText     =   "Efetuar inclus�o"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmReceita_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdCalcAdicao_Click()
    'Call TxtAdicD_GotFocus
    'Call TxtAdicE_GotFocus
    Call TxtAdicAO_GotFocus
End Sub

Private Sub CmdFechar_Click()
    VPStrResponse = MsgBox("Deseja incluir uma venda?", vbYesNo, "Pr� �tica 2004 - Informa��o")
    
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
    
    If VPStrResponse = vbYes Then
        FrmVenda_Inc.Show
    Else
        VGIntCodCli = 0
        VGStrNomeCli = ""
        VGStrForm = ""
    End If
    
End Sub

Private Sub CmdIncluirMed_Click()
    VGStrForm = "receita"
    FrmMedico_Inc.Show
End Sub

Private Sub CmdOK_Click()
    Conecta
    
    Dim RecRec As New ADODB.Recordset
    Dim VLIntCodMed As Long
    
    If CboMedico.Text = "" Then
        VLIntCodMed = 0
    Else
        VLIntCodMed = Mid(CboMedico.Text, Len(CboMedico.Text) - 20)
    End If
    
    StrSql = "SELECT * FROM tb_Receita"
    RecRec.Open StrSql, vgCon, 1, 3
        
    RecRec.AddNew
    RecRec("CodCli") = VGIntCodCli
    RecRec("CodMed") = VLIntCodMed
    RecRec("DtRec") = FormataData(Date)
    RecRec("LODEsf") = TxtLDEsf.Text
    RecRec("LODCil") = TxtLDCil.Text
    RecRec("LODEixo") = TxtLDEixo.Text
    RecRec("LOEEsf") = TxtLEEsf.Text
    RecRec("LOECil") = TxtLECil.Text
    RecRec("LOEEixo") = TxtLEEixo.Text
    RecRec("PODEsf") = TxtPDEsf.Text
    RecRec("PODCil") = TxtPDCil.Text
    RecRec("PODEixo") = TxtPDEixo.Text
    RecRec("POEEsf") = TxtPEEsf.Text
    RecRec("POECil") = TxtPECil.Text
    RecRec("POEEixo") = TxtPEEixo.Text
    RecRec("ODDNP") = TxtDNPD.Text
    RecRec("OEDNP") = TxtDNPE.Text
    RecRec("ODAlt") = TxtAltD.Text
    RecRec("OEAlt") = TxtAltE.Text
    RecRec("ODAdicao") = TxtAdicD.Text
    RecRec("OEAdicao") = TxtAdicE.Text
    RecRec("AOAdicao") = TxtAdicAO.Text
    RecRec("Obs") = TxtObsRec.Text
    RecRec.Update
        
    RecRec.Close
    
    Desconecta
    
    VPStrBox = MsgBox("Receita cadastrada.", vbInformation, "Pr� �tica 2004 - Informa��o")
    
    Unload Me
    
    FrmVenda_Inc.Show
End Sub

Private Sub Form_Resize()
  FrmReceita_Inc.Left = (MDIPrincipal.Width / 2) - (FrmReceita_Inc.Width / 2)
  FrmReceita_Inc.Top = (MDIPrincipal.Height / 3) - (FrmReceita_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 7950
    Width = 7800
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    LblNomeCli.Caption = VGStrNomeCli
    
    Call MontaCboMedico
    
End Sub

Sub MontaCboMedico()
    Dim RecCbo As New ADODB.Recordset
    
    CboMedico.Clear
    
    Conecta
    
    StrSql = "Select CodMed,Nome from tb_Medico"
    RecCbo.Open StrSql, vgCon, 1, 3
    
    CboMedico.AddItem ("")
    
    Do While Not RecCbo.EOF
        CboMedico.AddItem (RecCbo.Fields.Item(1).Value & "                                                                                                 " & RecCbo.Fields.Item(0).Value)
        RecCbo.MoveNext
    Loop
    
    Desconecta
End Sub

Private Sub TxtAdicAO_GotFocus()
    'If TxtAdicD.Text = "" Or TxtAdicE.Text = "" Then
    '    VPStrBox = MsgBox("Preencha os campos esf�ricos do OD e OE.", vbCritical, "Pr� �tica 2004 - Erro")
    '    TxtAdicAO.Text = ""
    'Else
    '    TxtAdicAO.Text = FormataNumDecRec(Val(TxtAdicD.Text) + Val(TxtAdicE.Text))
    'End If
    
    Dim POD As Currency
    Dim LOD As Currency
    Dim OD As Currency
    Dim POE As Currency
    Dim LOE As Currency
    Dim OE As Currency
    
    If (TxtLDEsf.Text = "" Or TxtPDEsf.Text = "") And (TxtLEEsf.Text = "" Or TxtPEEsf.Text = "") Then
        VPStrBox = MsgBox("Preencha os campos esf�ricos do OD e OE.", vbCritical, "Pr� �tica 2004 - Erro")
        TxtAdicAO.Text = ""
    Else
        POD = Val(Replace(TxtPDEsf.Text, "-", ""))
        LOD = Val(Replace(TxtLDEsf.Text, "-", ""))
        OD = POD - LOD
    
        POE = Val(Replace(TxtPEEsf.Text, "-", ""))
        LOE = Val(Replace(TxtLEEsf.Text, "-", ""))
        OE = POE - LOE
        
        If OD <> OE Then
            VPStrBox = MsgBox("Problemas na gera��o da adi��o." & Chr(13) & "� poss�vel que os dados estejam incorretos.", vbCritical, "Pr� �tica 2004 - Erro")
            
            TxtLDEsf.SetFocus
        Else
            TxtAdicAO.Text = FormataNumDecRec(OD)
        End If
    End If
End Sub

Private Sub TxtAdicAO_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

''Private Sub TxtAdicD_GotFocus()
''    Dim POD As Integer
''    Dim LOD As Integer
''
''    If TxtLDEsf.Text = "" Or TxtPDEsf.Text = "" Then
''        VPStrBox = MsgBox("Preencha os campos esf�ricos do OD.", vbCritical, "Pr� �tica 2004 - Erro")
''        TxtAdicD.Text = ""
''    Else
''        POD = Val(Replace(TxtPDEsf.Text, "-", ""))
''        LOD = Val(Replace(TxtLDEsf.Text, "-", ""))
''
''        'TxtAdicD.Text = FormataNumDecRec(Val(TxtLDEsf.Text) + Val(TxtPDEsf.Text))
''        TxtAdicD.Text = FormataNumDecRec(POD - LOD)
''    End If
''End Sub

''Private Sub TxtAdicD_KeyPress(KeyAscii As Integer)
''    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
''    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
''        KeyAscii = 0
''    End If
''End Sub

''Private Sub TxtAdicE_GotFocus()
''    Dim POE As Integer
''    Dim LOE As Integer
''
''    If TxtLEEsf.Text = "" Or TxtPEEsf.Text = "" Then
''        VPStrBox = MsgBox("Preencha os campos esf�ricos do OE.", vbCritical, "Pr� �tica 2004 - Erro")
''        TxtAdicE.Text = ""
''    Else
''        POE = Val(Replace(TxtPEEsf.Text, "-", ""))
''        LOE = Val(Replace(TxtLEEsf.Text, "-", ""))
''
''        TxtAdicE.Text = FormataNumDecRec(POE - LOE)
''        'TxtAdicE.Text = FormataNumDecRec(Val(TxtLEEsf.Text) + Val(TxtPEEsf.Text))
''    End If
''End Sub

''Private Sub TxtAdicE_KeyPress(KeyAscii As Integer)
''    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
''    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
''        KeyAscii = 0
''    End If
''End Sub

Private Sub TxtAltD_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAltE_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDNPD_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDNPD_LostFocus()
    TxtDNPD.Text = FormataNumDec(TxtDNPD.Text)
End Sub

Private Sub TxtDNPE_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDNPE_LostFocus()
    TxtDNPE.Text = FormataNumDec(TxtDNPE.Text)
End Sub

Private Sub TxtLDCil_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLDEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLDEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLDEsf_LostFocus()
    TxtLDEsf.Text = FormataNumDecRec(TxtLDEsf.Text)
End Sub

Private Sub TxtLDCil_LostFocus()
    TxtLDCil.Text = FormataNumDecRec(TxtLDCil.Text)
End Sub

Private Sub TxtLDEixo_LostFocus()
    TxtLDEixo.Text = FormataEixo(TxtLDEixo.Text)
End Sub

Private Sub TxtLECil_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLEEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLEEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLEEsf_LostFocus()
    TxtLEEsf.Text = FormataNumDecRec(TxtLEEsf.Text)
End Sub

Private Sub TxtLECil_LostFocus()
    TxtLECil.Text = FormataNumDecRec(TxtLECil.Text)
End Sub

Private Sub TxtLEEixo_LostFocus()
    TxtLEEixo.Text = FormataEixo(TxtLEEixo.Text)
End Sub

Private Sub TxtPDCil_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPDEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPDEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPDEsf_LostFocus()
    TxtPDEsf.Text = FormataNumDecRec(TxtPDEsf.Text)
End Sub

Private Sub TxtPDCil_LostFocus()
    TxtPDCil.Text = FormataNumDecRec(TxtPDCil.Text)
End Sub

Private Sub TxtPDEixo_LostFocus()
    TxtPDEixo.Text = FormataEixo(TxtPDEixo.Text)
End Sub

Private Sub TxtPECil_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPEEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPEEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita n�meros, ponto, v�rgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPEEsf_LostFocus()
    TxtPEEsf.Text = FormataNumDecRec(TxtPEEsf.Text)
End Sub

Private Sub TxtPECil_LostFocus()
    TxtPECil.Text = FormataNumDecRec(TxtPECil.Text)
End Sub

Private Sub TxtPEEixo_LostFocus()
    TxtPEEixo.Text = FormataEixo(TxtPEEixo.Text)
End Sub

