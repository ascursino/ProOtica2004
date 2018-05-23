VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmReceita_Alt 
   Caption         =   "Alteração de Receita"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
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
   Icon            =   "FrmReceita_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   7665
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
      TabIndex        =   25
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton CmdCalcAdicao 
         Caption         =   "Calcula adição"
         Height          =   495
         Left            =   5880
         TabIndex        =   93
         ToolTipText     =   "Calcular adição"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   2760
         TabIndex        =   77
         Top             =   2760
         Width           =   3135
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   78
            Top             =   120
            Width           =   2895
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":0CCA
               TabIndex        =   79
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   80
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":0D2E
               TabIndex        =   81
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2040
            TabIndex        =   82
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":0D8A
               TabIndex        =   83
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame28 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1080
            TabIndex        =   84
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":0DE6
               TabIndex        =   85
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame42 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   86
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtAdicD 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   16
               ToolTipText     =   "Adição olho direito"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame44 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1080
            TabIndex        =   87
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtAdicE 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   17
               ToolTipText     =   "Adição olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame48 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2040
            TabIndex        =   88
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtAdicAO 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   18
               ToolTipText     =   "Adição ambos os olhos"
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame41 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         TabIndex        =   64
         Top             =   2760
         Width           =   2655
         Begin VB.Frame Frame43 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   67
            Top             =   120
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":0E42
               TabIndex        =   68
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame45 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   1560
            TabIndex        =   65
            Top             =   120
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":0EA0
               TabIndex        =   66
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
            TabIndex        =   69
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
            TabIndex        =   70
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
            TabIndex        =   71
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
            TabIndex        =   72
            Top             =   600
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":0F00
               TabIndex        =   73
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
            TabIndex        =   74
            Top             =   1080
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":0F5C
               TabIndex        =   75
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
            TabIndex        =   76
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
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   3615
         Begin VB.Frame Frame14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   46
            Top             =   120
            Width           =   2895
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":0FB8
               TabIndex        =   47
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
            TabIndex        =   48
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":101A
               TabIndex        =   49
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
            TabIndex        =   50
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":107A
               TabIndex        =   51
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
            TabIndex        =   52
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":10DA
               TabIndex        =   53
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
            TabIndex        =   54
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtLDEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   0
               ToolTipText     =   "Longe grau esférico olho direito"
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
            TabIndex        =   55
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtLDCil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   1
               ToolTipText     =   "Longe grau cilíndrico olho direito"
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
            TabIndex        =   56
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
            TabIndex        =   57
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtLECil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   4
               ToolTipText     =   "Longe grau cilíndrico olho esquerdo"
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
            TabIndex        =   58
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
            TabIndex        =   59
            Top             =   1080
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":113A
               TabIndex        =   60
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
            TabIndex        =   61
            Top             =   1560
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":1196
               TabIndex        =   62
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
            TabIndex        =   63
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtLEEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   3
               ToolTipText     =   "Longe grau esférico olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame15 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   3720
         TabIndex        =   26
         Top             =   480
         Width           =   3615
         Begin VB.Frame Frame16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   600
            TabIndex        =   27
            Top             =   120
            Width           =   2895
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":11F2
               TabIndex        =   28
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
            TabIndex        =   29
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":1254
               TabIndex        =   30
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
            TabIndex        =   31
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":12B4
               TabIndex        =   32
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
            TabIndex        =   33
            Top             =   600
            Width           =   975
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "FrmReceita_Alt.frx":1314
               TabIndex        =   34
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
            TabIndex        =   35
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtPDEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   6
               ToolTipText     =   "Perto grau esférico olho direito"
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
            TabIndex        =   36
            Top             =   1080
            Width           =   975
            Begin VB.TextBox TxtPDCil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   7
               ToolTipText     =   "Perto grau cilíndrico olho direito"
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
            TabIndex        =   37
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
            TabIndex        =   38
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtPECil 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   10
               ToolTipText     =   "Perto grau cilíndrico olho esquerdo"
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
            TabIndex        =   39
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
            TabIndex        =   40
            Top             =   1080
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":1374
               TabIndex        =   41
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
            TabIndex        =   42
            Top             =   1560
            Width           =   495
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmReceita_Alt.frx":13D0
               TabIndex        =   43
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
            TabIndex        =   44
            Top             =   1560
            Width           =   975
            Begin VB.TextBox TxtPEEsf 
               Height          =   285
               Left            =   120
               MaxLength       =   6
               TabIndex        =   9
               ToolTipText     =   "Perto grau esférico olho esquerdo"
               Top             =   240
               Width           =   735
            End
         End
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
         TabIndex        =   20
         ToolTipText     =   "Adicionar médico"
         Top             =   4920
         Width           =   375
      End
      Begin VB.ComboBox CboMedico 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Nome do médico"
         Top             =   4920
         Width           =   5655
      End
      Begin VB.TextBox TxtObsRec 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Observação sobre a receita"
         Top             =   5640
         Width           =   7215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReceita_Alt.frx":142C
         TabIndex        =   89
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNomeCli 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmReceita_Alt.frx":1494
         TabIndex        =   90
         Top             =   240
         Width           =   5415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReceita_Alt.frx":150E
         TabIndex        =   91
         Top             =   4920
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmReceita_Alt.frx":1574
         TabIndex        =   92
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
      TabIndex        =   24
      Top             =   6600
      Width           =   7455
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1560
         OleObjectBlob   =   "FrmReceita_Alt.frx":15E2
         Top             =   240
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
         TabIndex        =   23
         ToolTipText     =   "Fechar janela"
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
         TabIndex        =   22
         ToolTipText     =   "Efetuar alteração"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmReceita_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdCalcAdicao_Click()
    Call TxtAdicD_GotFocus
    Call TxtAdicE_GotFocus
    Call TxtAdicAO_GotFocus
End Sub

Private Sub CmdFechar_Click()
    VGIntCodCli = 0
    VGStrNomeCli = ""
    VGStrForm = ""
    
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
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
    
    StrSql = "SELECT * FROM tb_Receita where CodRec=" & VGIntCodRec
    RecRec.Open StrSql, vgCon, 1, 3
        
    RecRec("CodMed") = VLIntCodMed
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
    
    VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Ótica 2004 - Informação")
    
    FrmPrincipal.CmdPesqRec.Value = True
    
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2

End Sub

Private Sub Form_Resize()
  FrmReceita_Alt.Left = (MDIPrincipal.Width / 2) - (FrmReceita_Alt.Width / 2)
  FrmReceita_Alt.Top = (MDIPrincipal.Height / 3) - (FrmReceita_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 7950
    Width = 7785
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    LblNomeCli.Caption = VGStrNomeCli

    Call MontaCboMedico
    
    Conecta
    
    Dim RecRec As New ADODB.Recordset
        
    StrSql = "SELECT LODEsf,LODCil,LODEixo,LOEEsf," & _
    "LOECil,LOEEixo,PODEsf,PODCil,PODEixo,POEEsf,POECil,POEEixo,ODDNP,OEDNP,ODAlt," & _
    "OEAlt,ODAdicao,OEAdicao,AOAdicao,R.Obs,M.Nome,R.CodMed " & _
    "FROM tb_Receita AS R,tb_Medico As M " & _
    "WHERE R.CodMed=M.CodMed and CodRec=" & VGIntCodRec
    RecRec.Open StrSql, vgCon, 1, 3
    
    TxtLDEsf.Text = VerificaNulo(RecRec.Fields.Item(0).Value)
    TxtLDCil.Text = VerificaNulo(RecRec.Fields.Item(1).Value)
    TxtLDEixo.Text = VerificaNulo(RecRec.Fields.Item(2).Value)
    TxtLEEsf.Text = VerificaNulo(RecRec.Fields.Item(3).Value)
    TxtLECil.Text = VerificaNulo(RecRec.Fields.Item(4).Value)
    TxtLEEixo.Text = VerificaNulo(RecRec.Fields.Item(5).Value)
    TxtPDEsf.Text = VerificaNulo(RecRec.Fields.Item(6).Value)
    TxtPDCil.Text = VerificaNulo(RecRec.Fields.Item(7).Value)
    TxtPDEixo.Text = VerificaNulo(RecRec.Fields.Item(8).Value)
    TxtPEEsf.Text = VerificaNulo(RecRec.Fields.Item(9).Value)
    TxtPECil.Text = VerificaNulo(RecRec.Fields.Item(10).Value)
    TxtPEEixo.Text = VerificaNulo(RecRec.Fields.Item(11).Value)
    TxtDNPD.Text = VerificaNulo(RecRec.Fields.Item(12).Value)
    TxtDNPE.Text = VerificaNulo(RecRec.Fields.Item(13).Value)
    TxtAltD.Text = VerificaNulo(RecRec.Fields.Item(14).Value)
    TxtAltE.Text = VerificaNulo(RecRec.Fields.Item(15).Value)
    TxtAdicD.Text = VerificaNulo(RecRec.Fields.Item(16).Value)
    TxtAdicE.Text = VerificaNulo(RecRec.Fields.Item(17).Value)
    TxtAdicAO.Text = VerificaNulo(RecRec.Fields.Item(18).Value)
    TxtObsRec.Text = VerificaNulo(RecRec.Fields.Item(19).Value)
    CboMedico.Text = RecRec.Fields.Item(20).Value & "                                                                                                 " & RecRec.Fields.Item(21).Value
    
    Desconecta
    
End Sub

Sub MontaCboMedico()
    Dim RecCbo As New ADODB.Recordset
    
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
    If TxtAdicD.Text = "" Or TxtAdicE.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos esféricos do OD e OE.", vbCritical, "Pró Ótica 2004 - Erro")
        TxtAdicAO.Text = ""
    Else
        TxtAdicAO.Text = FormataNumDecRec(Val(TxtAdicD.Text) + Val(TxtAdicE.Text))
    End If
End Sub

Private Sub TxtAdicD_GotFocus()
    If TxtLDEsf.Text = "" Or TxtPDEsf.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos esféricos do OD.", vbCritical, "Pró Ótica 2004 - Erro")
        TxtAdicD.Text = ""
    Else
        TxtAdicD.Text = FormataNumDecRec(Val(TxtLDEsf.Text) + Val(TxtPDEsf.Text))
    End If
End Sub

Private Sub TxtAdicE_GotFocus()
    If TxtLEEsf.Text = "" Or TxtPEEsf.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos esféricos do OE.", vbCritical, "Pró Ótica 2004 - Erro")
        TxtAdicE.Text = ""
    Else
        TxtAdicE.Text = FormataNumDecRec(Val(TxtLEEsf.Text) + Val(TxtPEEsf.Text))
    End If
End Sub

Private Sub TxtDNPD_LostFocus()
    TxtDNPD.Text = FormataNumDec(TxtDNPD.Text)
End Sub

Private Sub TxtDNPE_LostFocus()
    TxtDNPE.Text = FormataNumDec(TxtDNPE.Text)
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

Private Sub TxtLEEsf_LostFocus()
    TxtLEEsf.Text = FormataNumDecRec(TxtLEEsf.Text)
End Sub

Private Sub TxtLECil_LostFocus()
    TxtLECil.Text = FormataNumDecRec(TxtLECil.Text)
End Sub

Private Sub TxtLEEixo_LostFocus()
    TxtLEEixo.Text = FormataEixo(TxtLEEixo.Text)
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

Private Sub TxtPEEsf_LostFocus()
    TxtPEEsf.Text = FormataNumDecRec(TxtPEEsf.Text)
End Sub

Private Sub TxtPECil_LostFocus()
    TxtPECil.Text = FormataNumDecRec(TxtPECil.Text)
End Sub

Private Sub TxtPEEixo_LostFocus()
    TxtPEEixo.Text = FormataEixo(TxtPEEixo.Text)
End Sub

Private Sub TxtAdicAO_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAdicD_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAdicE_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAltD_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtAltE_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDNPD_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDNPE_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLDCil_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLDEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita números, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLDEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLECil_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLEEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita números, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtLEEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPDCil_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPDEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita números, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPDEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPECil_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPEEixo_KeyPress(KeyAscii As Integer)
    '=== Aceita números, enter, backspace, o sobrescrito ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 186 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPEEsf_KeyPress(KeyAscii As Integer)
    '=== Aceita números, ponto, vírgula, enter, backspace, sinal de - e + ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And (KeyAscii < 43 Or KeyAscii > 46) Then
        KeyAscii = 0
    End If
End Sub
