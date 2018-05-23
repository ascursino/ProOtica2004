VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FraVenda 
      Caption         =   "Consulta de Vendas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   10320
      TabIndex        =   154
      Top             =   7680
      Width           =   10935
      Begin FPSpread.vaSpread GridVenda 
         Height          =   3615
         Left            =   240
         TabIndex        =   139
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   8
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   6
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":0CCA
      End
      Begin VB.TextBox TxtDtVenda2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   135
         ToolTipText     =   "Maior data da venda"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenda1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   134
         ToolTipText     =   "Menor data da venda"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox CboTipoVenda 
         Height          =   315
         ItemData        =   "FrmPrincipal.frx":12A0
         Left            =   1560
         List            =   "FrmPrincipal.frx":12A2
         Style           =   2  'Dropdown List
         TabIndex        =   136
         ToolTipText     =   "Tipo da venda"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox TxtVendedor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   137
         ToolTipText     =   "Nome do vendedor"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox TxtCliVend 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   133
         ToolTipText     =   "Nome do cliente"
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton CmdPesqVenda 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   138
         ToolTipText     =   "Pesquisar vendas"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame FraBotaoVenda 
         Height          =   735
         Left            =   120
         TabIndex        =   155
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdCarne 
            Caption         =   "Imprimir &Carnê"
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
            Left            =   2640
            TabIndex        =   141
            ToolTipText     =   "Imprimir carnê"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton CmdIncluirVenda 
            Caption         =   "&Incluir"
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
            Left            =   5400
            TabIndex        =   142
            ToolTipText     =   "Incluir venda"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdDetVenda 
            Caption         =   "&Detalhe da venda"
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
            Left            =   240
            TabIndex        =   140
            ToolTipText     =   "Visualizar detalhe da venda"
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton CmdVendedorVenda 
            Caption         =   "&Vendedor"
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
            Left            =   9360
            TabIndex        =   146
            ToolTipText     =   "Dados dos vendedores"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdImprimirVenda 
            Caption         =   "I&mprimir"
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
            Left            =   8040
            TabIndex        =   144
            ToolTipText     =   "Imprimir consulta de venda"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirVenda 
            Caption         =   "&Excluir"
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
            Left            =   6720
            TabIndex        =   143
            ToolTipText     =   "Excluir venda"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":12A4
         TabIndex        =   157
         Top             =   600
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":130C
         TabIndex        =   158
         Top             =   1080
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmPrincipal.frx":137A
         TabIndex        =   159
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmPrincipal.frx":13E8
         TabIndex        =   160
         Top             =   1080
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalVend 
         Height          =   255
         Left            =   7560
         OleObjectBlob   =   "FrmPrincipal.frx":1452
         TabIndex        =   242
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7080
         TabIndex        =   161
         Top             =   600
         Width           =   105
      End
   End
   Begin VB.Frame FraCliente 
      Caption         =   "Consulta de Clientes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   9360
      TabIndex        =   227
      Top             =   7680
      Width           =   10935
      Begin VB.CommandButton CmdPesqCli 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   127
         ToolTipText     =   "Pesquisar clientes"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtBairroCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         MaxLength       =   60
         TabIndex        =   123
         ToolTipText     =   "Bairro do cliente"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox TxtTelCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   125
         ToolTipText     =   "Telefone do cliente"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtNomeCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   122
         ToolTipText     =   "Nome do cliente"
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox CboSexoCli 
         Height          =   315
         ItemData        =   "FrmPrincipal.frx":14DC
         Left            =   1080
         List            =   "FrmPrincipal.frx":14DE
         Style           =   2  'Dropdown List
         TabIndex        =   126
         ToolTipText     =   "Sexo do cliente"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TxtCpfCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   124
         ToolTipText     =   "Cpf do cliente"
         Top             =   840
         Width           =   2415
      End
      Begin VB.Frame FraBotaoCli 
         Height          =   735
         Left            =   120
         TabIndex        =   228
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdIncluirCli 
            Caption         =   "&Incluir"
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
            Left            =   5400
            TabIndex        =   129
            ToolTipText     =   "Incluir cliente"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarCli 
            Caption         =   "&Alterar"
            Enabled         =   0   'False
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
            Left            =   6720
            TabIndex        =   130
            ToolTipText     =   "Alterar cliente"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirCli 
            Caption         =   "&Excluir"
            Enabled         =   0   'False
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
            Left            =   8040
            TabIndex        =   131
            ToolTipText     =   "Excluir cliente"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdImprimirCli 
            Caption         =   "I&mprimir"
            Enabled         =   0   'False
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
            Left            =   9360
            TabIndex        =   132
            ToolTipText     =   "Imprimir consulta de clientes"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":14E0
         TabIndex        =   229
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmPrincipal.frx":1542
         TabIndex        =   230
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":15AC
         TabIndex        =   231
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmPrincipal.frx":160C
         TabIndex        =   232
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":1672
         TabIndex        =   233
         Top             =   1320
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalCli 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "FrmPrincipal.frx":16D4
         TabIndex        =   234
         Top             =   1680
         Width           =   2655
      End
      Begin FPSpread.vaSpread GridCliente 
         Height          =   3615
         Left            =   240
         TabIndex        =   128
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   16
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":1760
      End
   End
   Begin VB.Frame FraReceita 
      Caption         =   "Consulta de Receitas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   8640
      TabIndex        =   221
      Top             =   7680
      Width           =   10935
      Begin VB.CommandButton CmdPesqRec 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   116
         ToolTipText     =   "Pesquisar receitas"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtRecCliente 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   112
         ToolTipText     =   "Nome do cliente"
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox TxtRecMedico 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   113
         ToolTipText     =   "Nome do médico"
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox TxtDtRec2 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   115
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data da receita"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TxtDtRec1 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   114
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data da receita"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Frame FraBotaoRec 
         Height          =   735
         Left            =   120
         TabIndex        =   222
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdIncluirRec 
            Caption         =   "&Incluir"
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
            Left            =   5400
            TabIndex        =   118
            ToolTipText     =   "Incluir receita"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarRec 
            Caption         =   "&Alterar"
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
            Left            =   6720
            TabIndex        =   119
            ToolTipText     =   "Alterar receita"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirRec 
            Caption         =   "&Excluir"
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
            Left            =   8040
            TabIndex        =   120
            ToolTipText     =   "Excluir receita"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdImprimirRec 
            Caption         =   "I&mprimir"
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
            Left            =   9360
            TabIndex        =   121
            ToolTipText     =   "Imprimir consulta de receitas"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":1FAE
         TabIndex        =   223
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":2016
         TabIndex        =   224
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":207C
         TabIndex        =   225
         Top             =   1320
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalRec 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "FrmPrincipal.frx":2126
         TabIndex        =   226
         Top             =   1680
         Width           =   2655
      End
      Begin FPSpread.vaSpread GridReceita 
         Height          =   3615
         Left            =   240
         TabIndex        =   117
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   24
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":21B4
      End
   End
   Begin VB.Frame FraMedico 
      Caption         =   "Consulta de Médicos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   8040
      TabIndex        =   211
      Top             =   7680
      Width           =   10935
      Begin VB.CommandButton CmdPesqMed 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   106
         ToolTipText     =   "Pesquisar médicos"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtBairroMed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   60
         TabIndex        =   103
         ToolTipText     =   "Bairro do médico"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox TxtTelMed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   8
         TabIndex        =   105
         ToolTipText     =   "Telefone do médico"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtNomeMed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   100
         ToolTipText     =   "Nome do médico"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox TxtCrmMed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   102
         ToolTipText     =   "CRM do médico"
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox TxtCliConsMed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   200
         TabIndex        =   101
         ToolTipText     =   "Clínica/Consultório do médico"
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox TxtCpfMed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   104
         ToolTipText     =   "CPF do médico"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Frame FraBotaoMed 
         Height          =   735
         Left            =   120
         TabIndex        =   212
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdIncluirMed 
            Caption         =   "&Incluir"
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
            Left            =   5400
            TabIndex        =   108
            ToolTipText     =   "Incluir médico"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarMed 
            Caption         =   "&Alterar"
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
            Left            =   6720
            TabIndex        =   109
            ToolTipText     =   "Alterar médico"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirMed 
            Caption         =   "&Excluir"
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
            Left            =   8040
            TabIndex        =   110
            ToolTipText     =   "Excluir médico"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdImprimirMed 
            Caption         =   "I&mprimir"
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
            Left            =   9360
            TabIndex        =   111
            ToolTipText     =   "Imprimir consulta de médicos"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":2CA6
         TabIndex        =   214
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":2D08
         TabIndex        =   215
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":2D68
         TabIndex        =   216
         Top             =   1320
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmPrincipal.frx":2DC8
         TabIndex        =   217
         Top             =   360
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmPrincipal.frx":2E48
         TabIndex        =   218
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmPrincipal.frx":2EAE
         TabIndex        =   219
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalMed 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "FrmPrincipal.frx":2F18
         TabIndex        =   220
         Top             =   1680
         Width           =   2655
      End
      Begin FPSpread.vaSpread GridMedico 
         Height          =   3615
         Left            =   240
         TabIndex        =   107
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   16
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":2FA2
      End
   End
   Begin VB.Frame FraFornecedor 
      Caption         =   "Consulta de Fornecedores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   7440
      TabIndex        =   204
      Top             =   7680
      Width           =   10935
      Begin VB.TextBox TxtNomeForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   90
         ToolTipText     =   "Nome do fornecedor"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox TxtTelForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         TabIndex        =   91
         ToolTipText     =   "Telefone do fornecedor"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton CmdPesqForn 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   94
         ToolTipText     =   "Pesquisar fornecedores"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtTipoForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5880
         TabIndex        =   93
         ToolTipText     =   "Tipo de fornecimento"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TxtCnpjForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   92
         ToolTipText     =   "CNPJ do fornecedor"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Frame FraBotaoForn 
         Height          =   735
         Left            =   120
         TabIndex        =   205
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdImprimirForn 
            Caption         =   "I&mprimir"
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
            Left            =   9360
            TabIndex        =   99
            ToolTipText     =   "Imprimir consulta de fornecedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirForn 
            Caption         =   "&Incluir"
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
            Left            =   5400
            TabIndex        =   96
            ToolTipText     =   "Incluir fornecedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarForn 
            Caption         =   "&Alterar"
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
            Left            =   6720
            TabIndex        =   97
            ToolTipText     =   "Alterar fornecedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirForn 
            Caption         =   "&Excluir"
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
            Left            =   8040
            TabIndex        =   98
            ToolTipText     =   "Excluir fornecedor"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel60 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":37C9
         TabIndex        =   206
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel61 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmPrincipal.frx":3837
         TabIndex        =   207
         Top             =   1080
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel62 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":3899
         TabIndex        =   208
         Top             =   1080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel63 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmPrincipal.frx":38FB
         TabIndex        =   209
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalForn 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "FrmPrincipal.frx":3965
         TabIndex        =   210
         Top             =   1680
         Width           =   2895
      End
      Begin FPSpread.vaSpread GridFornecedor 
         Height          =   3615
         Left            =   240
         TabIndex        =   95
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   14
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":39F7
      End
   End
   Begin VB.Frame FraEstoque 
      Caption         =   "Consulta de Estoque"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   6720
      TabIndex        =   147
      Top             =   7680
      Width           =   10935
      Begin VB.CommandButton CmdPesqEst 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   84
         ToolTipText     =   "Pesquisar estoque"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtCorEst 
         Height          =   285
         Left            =   840
         MaxLength       =   20
         TabIndex        =   77
         ToolTipText     =   "Cor da armação"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox CboProdEst 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   75
         ToolTipText     =   "Tipo de produto"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox TxtQtdeMinEst 
         Height          =   285
         Left            =   7200
         MaxLength       =   9
         TabIndex        =   76
         ToolTipText     =   "Quantidade mínima em estoque"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtPteEst 
         Height          =   285
         Left            =   7200
         MaxLength       =   20
         TabIndex        =   81
         ToolTipText     =   "Tamanho da ponte da armação"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtAroEst 
         Height          =   285
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   80
         ToolTipText     =   "Tamanho do aro da armação"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtModEst 
         Height          =   285
         Left            =   4080
         MaxLength       =   80
         TabIndex        =   79
         ToolTipText     =   "Modelo da armação"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtNumEst 
         Height          =   285
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   78
         ToolTipText     =   "Número da armação"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtTipoEst 
         Height          =   285
         Left            =   840
         MaxLength       =   100
         TabIndex        =   82
         ToolTipText     =   "Tipo de lente"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxtChaEst 
         Height          =   285
         Left            =   2640
         MaxLength       =   200
         TabIndex        =   83
         ToolTipText     =   "Chave da lente"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame FraBotaoEst 
         Height          =   735
         Left            =   120
         TabIndex        =   148
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdImprimirEst 
            Caption         =   "I&mprimir"
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
            Left            =   9360
            TabIndex        =   89
            ToolTipText     =   "Imprimir consulta de estoque"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirAlterarEst 
            Caption         =   "&Incluir/Alterar"
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
            Left            =   6000
            TabIndex        =   87
            ToolTipText     =   "Incluir/Alterar estoque"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton CmdExcluirEst 
            Caption         =   "&Excluir"
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
            Left            =   8040
            TabIndex        =   88
            ToolTipText     =   "Excluir estoque"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox ChkDesatAlerta 
            Caption         =   "Desativar alerta"
            Height          =   195
            Left            =   240
            TabIndex        =   86
            ToolTipText     =   "Desativar alerta de estoque"
            Top             =   360
            Width           =   1575
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel57 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":4197
         TabIndex        =   149
         Top             =   960
         Width           =   6735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel58 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":42B3
         TabIndex        =   150
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel59 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "FrmPrincipal.frx":4325
         TabIndex        =   151
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalEst 
         Height          =   255
         Left            =   7440
         OleObjectBlob   =   "FrmPrincipal.frx":4395
         TabIndex        =   152
         Top             =   1680
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":4429
         TabIndex        =   153
         Top             =   1440
         Width           =   2175
      End
      Begin FPSpread.vaSpread GridEstoque 
         Height          =   3615
         Left            =   240
         TabIndex        =   85
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   8
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":44BB
      End
   End
   Begin VB.Frame FraProduto 
      Caption         =   "Consulta de Produto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   6120
      TabIndex        =   198
      Top             =   7680
      Width           =   10935
      Begin VB.ComboBox CboLenteProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   68
         ToolTipText     =   "Tipo de lente"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton CmdPesqProd 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   69
         ToolTipText     =   "Pesquisar produtos"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox CboFornProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   65
         ToolTipText     =   "Nome do fornecedor"
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox CboGriffeProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   67
         ToolTipText     =   "Nome da griffe"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox CboTipoProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   66
         ToolTipText     =   "Tipo de produto"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   199
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdImprimirProd 
            Caption         =   "I&mprimir"
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
            Left            =   9360
            TabIndex        =   74
            ToolTipText     =   "Imprimir consulta de produto"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirProd 
            Caption         =   "&Incluir"
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
            Left            =   5400
            TabIndex        =   71
            ToolTipText     =   "Incluir produto"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarProd 
            Caption         =   "&Alterar"
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
            Left            =   6720
            TabIndex        =   72
            ToolTipText     =   "Alterar produto"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirProd 
            Caption         =   "&Excluir"
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
            Left            =   8040
            TabIndex        =   73
            ToolTipText     =   "Excluir produto"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel54 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmPrincipal.frx":4AEF
         TabIndex        =   200
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel55 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmPrincipal.frx":4B5D
         TabIndex        =   201
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel56 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmPrincipal.frx":4BC3
         TabIndex        =   202
         Top             =   720
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalProd 
         Height          =   255
         Left            =   7440
         OleObjectBlob   =   "FrmPrincipal.frx":4C35
         TabIndex        =   203
         Top             =   1680
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmPrincipal.frx":4CC1
         TabIndex        =   235
         Top             =   1200
         Width           =   1215
      End
      Begin FPSpread.vaSpread GridProduto 
         Height          =   3615
         Left            =   240
         TabIndex        =   70
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   11
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":4D35
      End
   End
   Begin VB.Frame FraCrediario 
      Caption         =   "Consulta de Crediário"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      TabIndex        =   188
      Top             =   840
      Width           =   10935
      Begin VB.TextBox TxtCodParcCred 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   54
         ToolTipText     =   "Código da parcela do crediário"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVencCred2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   52
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data do vencimento"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVencCred1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   51
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data do vencimento"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtDtCred2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   49
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data do crediário"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtDtCred1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   48
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data do crediário"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox CboTipoCred 
         Height          =   315
         ItemData        =   "FrmPrincipal.frx":542A
         Left            =   1680
         List            =   "FrmPrincipal.frx":542C
         Style           =   2  'Dropdown List
         TabIndex        =   53
         ToolTipText     =   "Tipo de crediário"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton CmdPesqCred 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   55
         ToolTipText     =   "Pesquisar crediários"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtCredstaCred 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   50
         ToolTipText     =   "Nome do crediarista"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox TxtCliCred 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   47
         ToolTipText     =   "Nome do cliente"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Frame FraBotaoCred 
         Height          =   735
         Left            =   120
         TabIndex        =   189
         Top             =   5760
         Width           =   10695
         Begin VB.Frame FraCrediarista 
            Caption         =   "Crediarista"
            Height          =   735
            Left            =   0
            TabIndex        =   237
            Top             =   0
            Width           =   5415
            Begin VB.CommandButton CmdImprimirCredsta 
               Caption         =   "I&mprimir"
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
               Left            =   4080
               TabIndex        =   60
               ToolTipText     =   "Imprimir consulta de crediarista"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton CmdAlterarCredsta 
               Caption         =   "&Alterar"
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
               Left            =   1440
               TabIndex        =   58
               ToolTipText     =   "Alterar crediarista"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton CmdExcluirCredsta 
               Caption         =   "&Excluir"
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
               Left            =   2760
               TabIndex        =   59
               ToolTipText     =   "Excluir crediarista"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton CmdIncluirCredsta 
               Caption         =   "&Incluir"
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
               Left            =   120
               OLEDropMode     =   1  'Manual
               TabIndex        =   57
               ToolTipText     =   "Incluir crediarista"
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Crediário/Parcela"
            Height          =   735
            Left            =   5400
            TabIndex        =   190
            Top             =   0
            Width           =   5295
            Begin VB.CommandButton CmdImprimirCred 
               Caption         =   "I&mprimir"
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
               Left            =   2760
               TabIndex        =   63
               ToolTipText     =   "Imprimir consulta crediário"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton CmdAlterarCred 
               Caption         =   "&Alterar"
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
               Left            =   120
               TabIndex        =   61
               ToolTipText     =   "Alterar crediário"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton CmdExcluirCred 
               Caption         =   "&Excluir"
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
               Left            =   1440
               TabIndex        =   62
               ToolTipText     =   "Excluir crediário"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton CmdQuitarCred 
               Caption         =   "&Quitar"
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
               Left            =   4080
               TabIndex        =   64
               ToolTipText     =   "Quitar parcela do crediário"
               Top             =   240
               Width           =   1095
            End
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":542E
         TabIndex        =   191
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel48 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":5496
         TabIndex        =   192
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel49 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmPrincipal.frx":5506
         TabIndex        =   193
         Top             =   1200
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel50 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmPrincipal.frx":557A
         TabIndex        =   194
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmPrincipal.frx":55E8
         TabIndex        =   195
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel52 
         Height          =   255
         Left            =   6960
         OleObjectBlob   =   "FrmPrincipal.frx":5656
         TabIndex        =   196
         Top             =   480
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel53 
         Height          =   255
         Left            =   6960
         OleObjectBlob   =   "FrmPrincipal.frx":56B0
         TabIndex        =   197
         Top             =   840
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalCred 
         Height          =   255
         Left            =   7560
         OleObjectBlob   =   "FrmPrincipal.frx":570A
         TabIndex        =   236
         Top             =   1680
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmPrincipal.frx":579A
         TabIndex        =   156
         Top             =   1200
         Width           =   975
      End
      Begin FPSpread.vaSpread GridCrediario 
         Height          =   3615
         Left            =   240
         TabIndex        =   56
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   16
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":5806
      End
   End
   Begin VB.Frame FraCaixa 
      Caption         =   "Consulta de Movimento de Caixa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      TabIndex        =   185
      Top             =   840
      Width           =   10935
      Begin VB.ComboBox CboTipoPagtoCx 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   38
         ToolTipText     =   "Tipo de pagamento"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox TxtDtMovCx2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         MaxLength       =   200
         TabIndex        =   37
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data do movimento de caixa"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtDtMovCx1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         MaxLength       =   200
         TabIndex        =   36
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data do movimento de caixa"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CmdPesqCx 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   39
         ToolTipText     =   "Pesquisar movimento de caixa"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame FraBotaoCx 
         Height          =   735
         Left            =   120
         TabIndex        =   186
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdImprimirCx 
            Caption         =   "I&mprimir"
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
            Left            =   6720
            TabIndex        =   44
            ToolTipText     =   "Imprimir consulta de movimento de caixa"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirCx 
            Caption         =   "&Incluir"
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
            Left            =   2760
            TabIndex        =   41
            ToolTipText     =   "Incluir movimento de caixa"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarCx 
            Caption         =   "&Alterar"
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
            Left            =   4080
            TabIndex        =   42
            ToolTipText     =   "Alterar movimento de caixa"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirCx 
            Caption         =   "&Excluir"
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
            Left            =   5400
            TabIndex        =   43
            ToolTipText     =   "Excluir movimento de caixa"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdPagar 
            Caption         =   "A &pagar"
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
            Left            =   8040
            TabIndex        =   45
            ToolTipText     =   "Dados de contas a pagar"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdReceber 
            Caption         =   "A &receber"
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
            Left            =   9360
            TabIndex        =   46
            ToolTipText     =   "Dados de contas a receber"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "FrmPrincipal.frx":6088
         TabIndex        =   187
         Top             =   600
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "FrmPrincipal.frx":6104
         TabIndex        =   238
         Top             =   600
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel64 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "FrmPrincipal.frx":6162
         TabIndex        =   239
         Top             =   1080
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalCx 
         Height          =   255
         Left            =   6960
         OleObjectBlob   =   "FrmPrincipal.frx":61DE
         TabIndex        =   240
         Top             =   1680
         Width           =   3735
      End
      Begin FPSpread.vaSpread GridCaixa 
         Height          =   3615
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   7
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":6280
      End
   End
   Begin VB.Frame FraExtra 
      Caption         =   "Consulta de Opções Extras de Relatório"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      TabIndex        =   169
      Top             =   840
      Width           =   10935
      Begin VB.Frame FraExplic 
         Caption         =   "Folhetos explicativos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         TabIndex        =   179
         Top             =   2160
         Visible         =   0   'False
         Width           =   8055
         Begin VB.ComboBox CboFolheto 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmPrincipal.frx":6858
            Left            =   1800
            List            =   "FrmPrincipal.frx":686E
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Tipo de folheto explicativo"
            Top             =   1080
            Width           =   5415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "FrmPrincipal.frx":68E3
            TabIndex        =   180
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Frame FraMala 
         Caption         =   "Mala direta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         TabIndex        =   171
         Top             =   2160
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TxtDtNiverCli1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   32
            Text            =   "__/__/____"
            ToolTipText     =   "Menor data de vencimento"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox TxtDtNiverCli2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4680
            TabIndex        =   33
            Text            =   "__/__/____"
            ToolTipText     =   "Maior data de vencimento"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox TxtCliente 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   30
            ToolTipText     =   "Nome do cliente"
            Top             =   600
            Width           =   3135
         End
         Begin VB.ComboBox CboSexo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmPrincipal.frx":694B
            Left            =   3120
            List            =   "FrmPrincipal.frx":6955
            Style           =   2  'Dropdown List
            TabIndex        =   31
            ToolTipText     =   "Sexo do cliente"
            Top             =   1200
            Width           =   3135
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "FrmPrincipal.frx":696E
            TabIndex        =   172
            Top             =   600
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "FrmPrincipal.frx":69D6
            TabIndex        =   173
            Top             =   1200
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Index           =   1
            Left            =   2040
            OleObjectBlob   =   "FrmPrincipal.frx":6A38
            TabIndex        =   243
            Top             =   1800
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Index           =   1
            Left            =   4440
            OleObjectBlob   =   "FrmPrincipal.frx":6AA6
            TabIndex        =   244
            Top             =   1800
            Width           =   255
         End
      End
      Begin VB.Frame FraCob 
         Caption         =   "Cartas de cobrança"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         TabIndex        =   174
         Top             =   2160
         Visible         =   0   'False
         Width           =   8055
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":6B00
            TabIndex        =   253
            Top             =   840
            Width           =   7575
         End
         Begin VB.TextBox TxtClienteCob 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            ToolTipText     =   "Nome do cliente"
            Top             =   1320
            Width           =   6375
         End
         Begin VB.TextBox TxtDtVenc2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            TabIndex        =   29
            Text            =   "__/__/____"
            ToolTipText     =   "Maior data de vencimento"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TxtDtVenc1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   28
            Text            =   "__/__/____"
            ToolTipText     =   "Menor data de vencimento"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox CboTipoCarta 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmPrincipal.frx":6C06
            Left            =   1440
            List            =   "FrmPrincipal.frx":6C13
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Tipo de carta de cobrança"
            Top             =   480
            Width           =   6375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":6C53
            TabIndex        =   175
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":6CB5
            TabIndex        =   176
            Top             =   1320
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":6D1D
            TabIndex        =   177
            Top             =   1920
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "FrmPrincipal.frx":6D8B
            TabIndex        =   178
            Top             =   1920
            Width           =   255
         End
      End
      Begin VB.Frame FraNiver 
         Caption         =   "Aniversariantes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         TabIndex        =   181
         Top             =   2160
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TxtMes1 
            Height          =   285
            Left            =   2640
            TabIndex        =   24
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox TxtMes2 
            Height          =   285
            Left            =   3480
            TabIndex        =   25
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox TxtDia1 
            Height          =   285
            Left            =   2640
            TabIndex        =   22
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox TxtDia2 
            Height          =   285
            Left            =   3480
            TabIndex        =   23
            Top             =   720
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Index           =   0
            Left            =   2160
            OleObjectBlob   =   "FrmPrincipal.frx":6DE5
            TabIndex        =   182
            Top             =   720
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Index           =   0
            Left            =   3960
            OleObjectBlob   =   "FrmPrincipal.frx":6E45
            TabIndex        =   213
            Top             =   1440
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Index           =   2
            Left            =   3240
            OleObjectBlob   =   "FrmPrincipal.frx":6E9D
            TabIndex        =   245
            Top             =   720
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Index           =   2
            Left            =   2160
            OleObjectBlob   =   "FrmPrincipal.frx":6EF7
            TabIndex        =   246
            Top             =   1440
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Index           =   3
            Left            =   3240
            OleObjectBlob   =   "FrmPrincipal.frx":6F57
            TabIndex        =   247
            Top             =   1440
            Width           =   255
         End
      End
      Begin VB.Frame FraPropCred 
         Caption         =   "Proposta de crédito"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         TabIndex        =   248
         Top             =   2160
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TxtCredstaProp 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   19
            ToolTipText     =   "Nome do cliente"
            Top             =   1080
            Width           =   5895
         End
         Begin VB.TextBox TxtDtVendaProp01 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   20
            Text            =   "__/__/____"
            ToolTipText     =   "Menor data de vencimento"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox TxtDtVendaProp02 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3480
            TabIndex        =   21
            Text            =   "__/__/____"
            ToolTipText     =   "Maior data de vencimento"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox TxtClienteProp 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   18
            ToolTipText     =   "Nome do cliente"
            Top             =   480
            Width           =   5895
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":6FB1
            TabIndex        =   249
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":7019
            TabIndex        =   250
            Top             =   1680
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "FrmPrincipal.frx":7095
            TabIndex        =   251
            Top             =   1680
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":70EF
            TabIndex        =   252
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.OptionButton OptPropCred 
         Caption         =   "Proposta de crédito"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         ToolTipText     =   "Consulta de mala direta"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton OptCob 
         Caption         =   "Cartas de cobrança"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   13
         ToolTipText     =   "Consulta de cartas de cobrança à clientes"
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton OptEtiqArm 
         Caption         =   "Etiquetas para armações"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Consulta de etiquetas para armações"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton OptNiver 
         Caption         =   "Aniversariantes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         ToolTipText     =   "Consulta de aniversariantes"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton OptMala 
         Caption         =   "Mala direta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         ToolTipText     =   "Consulta de mala direta"
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton OptExplic 
         Caption         =   "Folhetos explicativos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         ToolTipText     =   "Consulta de folhetos explicativos"
         Top             =   840
         Width           =   2175
      End
      Begin VB.Frame FraEtiqArm 
         Caption         =   "Etiquetas para armações"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1560
         TabIndex        =   183
         Top             =   2160
         Visible         =   0   'False
         Width           =   8055
         Begin VB.ComboBox CboGriffe 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Nome da griffe"
            Top             =   1080
            Width           =   5895
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Left            =   720
            OleObjectBlob   =   "FrmPrincipal.frx":715F
            TabIndex        =   184
            Top             =   1080
            Width           =   615
         End
      End
      Begin VB.Frame FraBotaoExt 
         Height          =   735
         Left            =   120
         TabIndex        =   170
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdImprimirExt 
            Caption         =   "&Imprimir"
            Enabled         =   0   'False
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
            Left            =   9360
            TabIndex        =   35
            ToolTipText     =   "Imprimir consulta"
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Frame FraOrcamento 
      Caption         =   "Consulta de Orçamentos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      TabIndex        =   162
      Top             =   840
      Width           =   10935
      Begin VB.TextBox TxtDtOrc2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   4
         ToolTipText     =   "Maior data de cadastro"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtDtOrc1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   3
         ToolTipText     =   "Menor data de cadastro"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtVendOrc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   1
         ToolTipText     =   "Nome do vendedor"
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox TxtTelOrc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Telefone do cliente"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdPesqOrc 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   5
         ToolTipText     =   "Pesquisar orçamentos"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TxtCliOrc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Nome do cliente"
         Top             =   600
         Width           =   2775
      End
      Begin VB.Frame FraBotaoOrc 
         Height          =   735
         Left            =   120
         TabIndex        =   163
         Top             =   5760
         Width           =   10695
         Begin VB.CommandButton CmdImprimirOrc 
            Caption         =   "I&mprimir"
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
            Left            =   9360
            TabIndex        =   10
            ToolTipText     =   "Imprimir consulta de orçamento"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirOrc 
            Caption         =   "&Incluir"
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
            Left            =   5400
            TabIndex        =   7
            ToolTipText     =   "Incluir orçamento"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarOrc 
            Caption         =   "&Alterar"
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
            Left            =   6720
            TabIndex        =   8
            ToolTipText     =   "Alterar orçamento"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirOrc 
            Caption         =   "&Excluir"
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
            Left            =   8040
            TabIndex        =   9
            ToolTipText     =   "Excluir orçamento"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":71C5
         TabIndex        =   164
         Top             =   600
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmPrincipal.frx":722D
         TabIndex        =   165
         Top             =   1080
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":728F
         TabIndex        =   166
         Top             =   1080
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "FrmPrincipal.frx":72F9
         TabIndex        =   167
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalOrc 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "FrmPrincipal.frx":7363
         TabIndex        =   241
         Top             =   1680
         Width           =   3015
      End
      Begin FPSpread.vaSpread GridOrcamento 
         Height          =   3615
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6376
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   20
         MaxRows         =   1
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":73F3
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6960
         TabIndex        =   168
         Top             =   1080
         Width           =   105
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "FrmPrincipal.frx":7DAA
      Top             =   7680
   End
   Begin MSComctlLib.TabStrip TabPrincipal 
      Height          =   7335
      Left            =   240
      TabIndex        =   145
      Top             =   240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12938
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   11
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "VENDA"
            Object.Tag             =   "Controle de vendas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CLIENTE"
            Object.Tag             =   "Controle de clientes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "RECEITA"
            Object.Tag             =   "Controle de receitas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MÉDICO"
            Object.Tag             =   "Controle de médicos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "FORNECEDOR"
            Object.Tag             =   "Controle de fornecedores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ESTOQUE"
            Object.Tag             =   "Controle de estoque"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PRODUTO"
            Object.Tag             =   "Controle de produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CREDIÁRIO"
            Object.Tag             =   "Controle de crediário"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CAIXA"
            Object.Tag             =   "Controle do movimento de caixa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "EXTRA"
            Object.Tag             =   "Controle de opções extras"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ORÇAMENTO"
            Object.Tag             =   "Controle de orçamentos"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrCrediarioQuitado As String
Public RecPesq As New ADODB.Recordset

Private Sub ChkDesatAlerta_Click()
    Screen.MousePointer = vbHourglass
    Conecta

    If ChkDesatAlerta.Value = 1 Then
        'desativar o alerta
        vgCon.Execute ("Update tb_Alerta Set Ativado='não'")
    Else
        'ativar o alerta
        vgCon.Execute ("Update tb_Alerta Set Ativado='sim'")
    End If
    
    Desconecta
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdAlterarCli_Click()
    FrmCliente_Alt.Show
End Sub

Private Sub CmdAlterarCred_Click()
    FrmCrediario_Alt.Show
End Sub

Private Sub CmdAlterarCredsta_Click()
    If VGIntCodCredsta = 0 Then
        FrmCrediarista_Lista.Show
    Else
        FrmCrediarista_Alt.Show
    End If
End Sub

Private Sub CmdAlterarCx_Click()
    FrmCaixa_Alt.Show
End Sub

Private Sub CmdAlterarExt_Click()
    If OptExplic.Value = True Then
        frmExtra_folheto_alt.Show
        
    ElseIf OptCob.Value = True Then
        frmExtra_cartacob_alt.Show
        
    End If
End Sub

Private Sub CmdAlterarForn_Click()
    FrmFornecedor_Alt.Show
End Sub

Private Sub CmdAlterarMed_Click()
    FrmMedico_Alt.Show
End Sub

Private Sub CmdAlterarOrc_Click()
    FrmOrcamento_Alt.Show
End Sub

Private Sub CmdAlterarProd_Click()
    FrmProduto_Alt.Show
End Sub

Private Sub CmdAlterarRec_Click()
    FrmReceita_Alt.Show
End Sub

Private Sub CmdCarne_Click()
    
    Dim RecCar As New ADODB.Recordset
    Dim codparc As String
    Dim datacred As String
    Dim valortotal As String
    Dim parcela As String
    Dim vencimento As String
    Dim valor As String
    
    Conecta
    
    StrSql = "Select CR.DtCred,CR.Parcela,CR.ValorTotal,P.CodParc,P.NumParc,P.Vencimento,P.Valor,C.Nome " & _
             "From tb_Crediario as CR,tb_Crediario_Parcela as P,tb_Cliente as C,tb_Venda as V " & _
             "Where P.CodCred=CR.CodCred and C.CodCli=CR.CodCli and CR.CodCred=V.CodCred and V.CodVenda=" & VGIntCodVenda & " order by P.NumParc"
    RecPesq.Open StrSql, vgCon, 1, 3
    
    Do While Not RecPesq.EOF
        codpar = FormataNum(RecPesq.Fields.Item(3).Value)
        datacred = FormataData(RecPesq.Fields.Item(0).Value)
        valortotal = FormataMoeda(RecPesq.Fields.Item(2).Value)
        parcela = FormataNum(RecPesq.Fields.Item(4).Value) & "/" & FormataNum(RecPesq.Fields.Item(1).Value)
        vencimento = FormataData(RecPesq.Fields.Item(5).Value)
        valor = FormataMoeda(RecPesq.Fields.Item(6).Value)
        VGStrClienteRel = RecPesq.Fields.Item(7).Value
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06) " & _
        "VALUES ('" & codpar & "','" & datacred & "','" & valortotal & "','" & parcela & "','" & vencimento & "','" & valor & " ')"
         
        RecPesq.MoveNext
    Loop
    
    Desconecta
    
    rptCarne.Show
End Sub

Private Sub CmdDetVenda_Click()
    FrmVenda_Detalhe.Show
End Sub

Private Sub CmdExcluirCli_Click()
    VPStrResponse = MsgBox("Deseja excluir este cliente e suas receitas?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Cliente WHERE CodCli=" & VGIntCodCli)
        vgCon.Execute ("DELETE FROM tb_Receita WHERE CodCli=" & VGIntCodCli)
        Desconecta
        
        FrmPrincipal.CmdPesqCli.Value = True
    End If
End Sub

Private Sub CmdExcluirRec_Click()
    VPStrResponse = MsgBox("Deseja excluir esta receita?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Receita WHERE CodRec=" & VGIntCodRec)
        Desconecta
        FrmPrincipal.CmdPesqRec.Value = True
    End If
End Sub

Private Sub CmdExcluirCred_Click()
    
    VPStrResponse = MsgBox("Esta ação apagará todos os dados referente a este crediário." & Chr(13) & "Deseja continuar?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        
        Conecta
        
        vgCon.Execute ("DELETE FROM tb_Crediario_Parcela_Quitacao WHERE CodCred=" & VGIntCodCred)
        
        vgCon.Execute ("DELETE FROM tb_Crediario_Parcela WHERE CodCred=" & VGIntCodCred)
        
        vgCon.Execute ("DELETE FROM tb_Crediario WHERE CodCred=" & VGIntCodCred)
        
        Desconecta
        
        FrmPrincipal.CmdPesqCred.Value = True
    End If

End Sub

Private Sub CmdExcluirCredsta_Click()
    
    Dim RecCredsta As New ADODB.Recordset
    
    VPStrResponse = MsgBox("Deseja excluir este crediarista?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
    
        Conecta
        
        StrSql = "Select CodCredsta from tb_Crediario where CodCredsta=" & VGIntCodCredsta
        RecCredsta.Open StrSql, vgCon, 1, 3
        
        If Not RecCredsta.EOF Then
            Desconecta
            VPStrBox = MsgBox("Existe crediário aberto em nome deste crediarista." & Chr(13) & "Crediarista não poderá ser excluído.", vbInformation, "Pró Ótica 2004 - Informação")
        Else
            vgCon.Execute ("DELETE FROM tb_Crediarista WHERE CodCredsta=" & VGIntCodCredsta)
            Desconecta
            
            FrmPrincipal.CmdPesqCred.Value = True
        End If
    End If
End Sub

Private Sub CmdExcluirCx_Click()
    VPStrResponse = MsgBox("Deseja excluir este movimento do caixa?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
    
        Conecta
        vgCon.Execute ("DELETE FROM tb_Caixa WHERE CodCx=" & VGIntCodCx)
        Desconecta
        
        FrmPrincipal.CmdPesqCx.Value = True
        
    End If
End Sub

Private Sub CmdExcluirEst_Click()
    VPStrResponse = MsgBox("Deseja excluir este produto do estoque?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Estoque WHERE CodEst=" & VGIntCodEst)
        Desconecta
        
        FrmPrincipal.CmdPesqEst.Value = True
    End If
End Sub

Private Sub CmdExcluirForn_Click()
    VPStrResponse = MsgBox("Deseja excluir este fornecedor?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Fornecedor WHERE CodForn=" & VGIntCodForn)
        Desconecta
        
        FrmPrincipal.CmdPesqForn.Value = True
    End If
End Sub

Private Sub CmdExcluirMed_Click()
    
    VPStrResponse = MsgBox("Deseja excluir este médico?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        
        Dim RecMed As New ADODB.Recordset
        StrSql = "Select CodRec from tb_Receita where CodMed=" & VGIntCodMed
        RecMed.Open StrSql, vgCon, 1, 3
        
        If Not RecMed.EOF Then
            If RecMed.RecordCount = 1 Then
                VPStrBox = MsgBox("Existe " & FormataNum(RecMed.RecordCount) & " receita vinculada a este médico." & Chr(13) & "Este médico não poderá ser excluído.", vbInformation, "Pró Ótica 2004 - Informação")
            ElseIf RecMed.RecordCount > 1 Then
                VPStrBox = MsgBox("Existem " & FormataNum(RecMed.RecordCount) & " receitas vinculadas a este médico." & Chr(13) & "Este médico não poderá ser excluído.", vbInformation, "Pró Ótica 2004 - Informação")
            End If
            
            Desconecta
        
        Else
            vgCon.Execute ("DELETE FROM tb_Medico WHERE CodMed=" & VGIntCodMed)
            
            Desconecta
            FrmPrincipal.CmdPesqMed.Value = True
        
        End If
    End If
    
End Sub

Private Sub CmdExcluirOrc_Click()
    VPStrResponse = MsgBox("Deseja excluir este orçamento?", vbYesNo, "Pró Ótica 2004 - Informação")
    If VPStrResponse = vbYes Then
    
        Conecta
        vgCon.Execute ("DELETE FROM tb_Orcamento WHERE CodOrc=" & VGIntCodOrc)
        Desconecta
        
        FrmPrincipal.CmdPesqOrc.Value = True
    End If
End Sub

Private Sub CmdExcluirProd_Click()
    VPStrResponse = MsgBox("Deseja excluir este produto e seus" & Chr(13) & "lançamentos de estoque?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Produto WHERE CodProd=" & VGIntCodProd)
        vgCon.Execute ("DELETE FROM tb_Estoque WHERE CodProd=" & VGIntCodProd)
        Desconecta
        
        FrmPrincipal.CmdPesqProd.Value = True
        
    End If
End Sub

Private Sub CmdExcluirVenda_Click()
    Dim RecVenda As New ADODB.Recordset
    Dim RecEst As New ADODB.Recordset
    
    VPStrResponse = MsgBox("Deseja excluir esta venda?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        VPStrResponse = MsgBox("Antes de excluir essa venda, deseja" & Chr(13) & "retornar o(s) produto(s) ao estoque?", vbYesNo, "Pró Ótica 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            
            StrSql = "SELECT CodProd01,CodProd02,CodProd03,CodProd04,CodProd05,CodProd06," & _
                     "Qtde01,Qtde02,Qtde03,Qtde04,Qtde05,Qtde06 " & _
                     "FROM tb_Venda where CodVenda=" & VGIntCodVenda
            RecVenda.Open StrSql, vgCon, 1, 3
            
            '===== Retorna estoque do produto 01 =============
            If RecVenda!CodProd01 <> "" And RecVenda!CodProd01 <> 0 And IsNull(RecVenda!CodProd01) = False Then
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & RecVenda!CodProd01
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    RecEst("QtdeProd") = Val(RecEst!qtdeprod) + Val(RecVenda!Qtde01)
                    RecEst.Update
                End If
                RecEst.Close
            End If
          
            '===== Retorna estoque do produto 02 =============
            If RecVenda!CodProd02 <> "" And RecVenda!CodProd02 <> 0 And IsNull(RecVenda!CodProd02) = False Then
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & RecVenda!CodProd02
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    RecEst("QtdeProd") = Val(RecEst!qtdeprod) + Val(RecVenda!Qtde02)
                    RecEst.Update
                End If
                RecEst.Close
            End If
          
            '===== Retorna estoque do produto 03 =============
            If RecVenda!CodProd03 <> "" And RecVenda!CodProd03 <> 0 And IsNull(RecVenda!CodProd03) = False Then
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & RecVenda!CodProd03
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    RecEst("QtdeProd") = Val(RecEst!qtdeprod) + Val(RecVenda!Qtde03)
                    RecEst.Update
                End If
                RecEst.Close
            End If
          
            '===== Retorna estoque do produto 04 =============
            If RecVenda!CodProd04 <> "" And RecVenda!CodProd04 <> 0 And IsNull(RecVenda!CodProd04) = False Then
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & RecVenda!CodProd04
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    RecEst("QtdeProd") = Val(RecEst!qtdeprod) + Val(RecVenda!Qtde04)
                    RecEst.Update
                End If
                RecEst.Close
            End If
          
            '===== Retorna estoque do produto 05 =============
            If RecVenda!CodProd05 <> "" And RecVenda!CodProd05 <> 0 And IsNull(RecVenda!CodProd05) = False Then
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & RecVenda!CodProd05
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    RecEst("QtdeProd") = Val(RecEst!qtdeprod) + Val(RecVenda!Qtde05)
                    RecEst.Update
                End If
                RecEst.Close
            End If
          
            '===== Retorna estoque do produto 06 =============
            If RecVenda!CodProd06 <> "" And RecVenda!CodProd06 <> 0 And IsNull(RecVenda!CodProd06) = False Then
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & RecVenda!CodProd06
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    RecEst("QtdeProd") = Val(RecEst!qtdeprod) + Val(RecVenda!Qtde06)
                    RecEst.Update
                End If
                RecEst.Close
            End If
            
            Desconecta
        End If
        
        Conecta
        vgCon.Execute ("DELETE FROM tb_Venda WHERE CodVenda=" & VGIntCodVenda)
        Desconecta
                
        VPStrBox = MsgBox("Venda excluída!", vbInformation, "Pró Ótica 2004 - Informação")
        
        FrmPrincipal.CmdPesqVenda.Value = True
    End If
End Sub

Private Sub CmdImprimirCli_Click()
    Screen.MousePointer = vbHourglass
    
    Dim datacad As String
    Dim nome As String
    Dim sexo As String
    Dim endereco As String
    Dim bairro As String
    Dim cep As String
    Dim cidest As String
    Dim datanasc As String
    Dim tel As String
    Dim cel As String
    Dim cpf As String
    Dim email As String
    Dim obs As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridCliente.MaxRows
        
        GridCliente.Col = 1
        GridCliente.Row = VLStrLinha
        nome = GridCliente.Text
        
        GridCliente.Col = 2
        GridCliente.Row = VLStrLinha
        datacad = GridCliente.Text
        
        GridCliente.Col = 3
        GridCliente.Row = VLStrLinha
        sexo = GridCliente.Text
        
        GridCliente.Col = 4
        GridCliente.Row = VLStrLinha
        endereco = GridCliente.Text
        
        GridCliente.Col = 5
        GridCliente.Row = VLStrLinha
        bairro = GridCliente.Text
        
        GridCliente.Col = 6
        GridCliente.Row = VLStrLinha
        cep = GridCliente.Text
        
        GridCliente.Col = 7
        GridCliente.Row = VLStrLinha
        cidest = GridCliente.Text
        
        GridCliente.Col = 8
        GridCliente.Row = VLStrLinha
        cidest = cidest & "/" & GridCliente.Text
        
        GridCliente.Col = 9
        GridCliente.Row = VLStrLinha
        datanasc = GridCliente.Text
        
        GridCliente.Col = 10
        GridCliente.Row = VLStrLinha
        tel = GridCliente.Text
        
        GridCliente.Col = 11
        GridCliente.Row = VLStrLinha
        cel = GridCliente.Text
        
        GridCliente.Col = 12
        GridCliente.Row = VLStrLinha
        cpf = GridCliente.Text
        
        GridCliente.Col = 13
        GridCliente.Row = VLStrLinha
        email = GridCliente.Text
        
        GridCliente.Col = 14
        GridCliente.Row = VLStrLinha
        obs = GridCliente.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13) " & _
        "VALUES ('" & datacad & "','" & nome & "','" & sexo & "','" & endereco & "','" & bairro & "','" & cep & "','" & cidest & "','" & datanasc & "','" & tel & "','" & cel & "','" & cpf & "','" & email & "','" & obs & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCliente.Show

End Sub

Private Sub CmdImprimirCred_Click()
    Screen.MousePointer = vbHourglass
    
    Dim cliente As String
    Dim datacred As String
    Dim tipocred As String
    Dim valorvenda As String
    Dim juros As String
    Dim valortotal As String
    Dim tipoentr As String
    Dim valorentr As String
    Dim parc As String
    Dim venc As String
    Dim valor As String
    Dim quitado As String
    
    Dim VLStrLinha As String
    Dim clientetemp As String
    Dim clientegrid As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridCrediario.MaxRows
        
        GridCrediario.Col = 1
        GridCrediario.Row = VLStrLinha
        clientegrid = GridCrediario.Text
        
        If clientetemp = clientegrid Then
            cliente = GridCrediario.Text
            
            GridCrediario.Col = 3
            GridCrediario.Row = VLStrLinha
            datacred = GridCrediario.Text
            
            GridCrediario.Col = 4
            GridCrediario.Row = VLStrLinha
            tipocred = GridCrediario.Text
            
            GridCrediario.Col = 5
            GridCrediario.Row = VLStrLinha
            valorvenda = GridCrediario.Text
            
            GridCrediario.Col = 6
            GridCrediario.Row = VLStrLinha
            juros = GridCrediario.Text
            
            GridCrediario.Col = 7
            GridCrediario.Row = VLStrLinha
            valortotal = GridCrediario.Text
            
            GridCrediario.Col = 8
            GridCrediario.Row = VLStrLinha
            tipoentr = GridCrediario.Text
            
            GridCrediario.Col = 9
            GridCrediario.Row = VLStrLinha
            valorentr = GridCrediario.Text
            
            GridCrediario.Col = 10
            GridCrediario.Row = VLStrLinha
            parc = GridCrediario.Text
            
            GridCrediario.Col = 11
            GridCrediario.Row = VLStrLinha
            venc = GridCrediario.Text
            
            GridCrediario.Col = 12
            GridCrediario.Row = VLStrLinha
            valor = GridCrediario.Text
            
            GridCrediario.Col = 13
            GridCrediario.Row = VLStrLinha
            quitado = GridCrediario.Text
            
            VLStrLinha = VLStrLinha + 1
        Else
            cliente = ""
            datacred = ""
            tipocred = ""
            valorvenda = ""
            juros = ""
            valortotal = ""
            tipoentr = ""
            valorentr = ""
            parc = ""
            venc = ""
            valor = ""
            quitado = ""
        End If
        
        If clientetemp <> "" Then
            vgCon.Execute "INSERT INTO tb_Auxiliar " & _
            "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12) " & _
            "VALUES ('" & cliente & "','" & datacred & "','" & tipocred & "','" & valorvenda & "','" & juros & "','" & valortotal & "','" & tipoentr & "','" & valorentr & "','" & parc & "','" & venc & "','" & valor & "','" & quitado & "')"
        End If
        
        clientetemp = clientegrid

    Loop
    
    Desconecta
        
    rptCrediario.Show

End Sub

Private Sub CmdImprimirCredsta_Click()
    Screen.MousePointer = vbHourglass
    
    Dim nome As String
    Dim endereco As String
    Dim bairro As String
    Dim cep As String
    Dim cidest As String
    Dim datanasc As String
    Dim tel As String
    Dim cel As String
    Dim cpf As String
    Dim email As String
    Dim obs As String
    Dim nomecli As String
    Dim datacred As String
    Dim tipocred As String
    Dim valorvenda As String
    Dim parcela As String
        
    Dim VLStrLinha As String
    Dim cliente As String
    Dim clientetemp As String
    Dim TotalCredTemp As Long
    
    Dim RecCredsta As New ADODB.Recordset
    Dim RecCred As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    
    Conecta
    
    StrSql = "Select * from tb_Crediarista where CodCredsta=" & VGIntCodCredsta
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    If RecCredsta.EOF Then
        VPStrBox = MsgBox("Selecione algum crediarista no resultado da pesquisa.", vbInformation, "Pró Ótica 2004 - Informação")
        Desconecta
        Screen.MousePointer = vbNormal
    Else
        StrSql = "Select * from tb_Crediario where CodCredsta=" & VGIntCodCredsta
        RecCred.Open StrSql, vgCon, 1, 3
                
        If Not RecCred.EOF Then
            VGIntTotalCred = RecCred.RecordCount
            TotalCredTemp = VGIntTotalCred
        Else
            VGIntTotalCred = 0
        End If
        
        VLStrLinha = 1
        
        Do While TotalCredTemp <> 0
            
            If Not RecCred.EOF Then
                StrSql = "Select Nome from tb_Cliente where CodCli=" & RecCred!CodCli
                RecCli.Open StrSql, vgCon, 1, 3
                cliente = RecCli.Fields.Item(0).Value
            Else
                cliente = ""
            End If
            
            If clientetemp <> cliente Then
            
                nome = RecCredsta.Fields.Item(1).Value
                cpf = RecCredsta.Fields.Item(10).Value
                endereco = RecCredsta.Fields.Item(2).Value
                bairro = RecCredsta.Fields.Item(3).Value
                cep = RecCredsta.Fields.Item(4).Value
                cidest = RecCredsta.Fields.Item(5).Value & "/" & RecCredsta.Fields.Item(6).Value
                datanasc = FormataData(RecCredsta.Fields.Item(7).Value)
                tel = RecCredsta.Fields.Item(8).Value
                cel = RecCredsta.Fields.Item(9).Value
                email = RecCredsta.Fields.Item(11).Value
                If IsNull(RecCredsta.Fields.Item(12).Value) = False Then
                    obs = RecCredsta.Fields.Item(12).Value
                Else
                    obs = ""
                End If
                nomecli = cliente
                datacred = FormataData(RecCred.Fields.Item(3).Value)
                tipocred = RecCred.Fields.Item(4).Value
                valorvenda = FormataMoeda(RecCred.Fields.Item(5).Value)
                parcela = FormataNum(RecCred.Fields.Item(6).Value)
                
                vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16) " & _
                "VALUES ('" & nome & "','" & cpf & "','" & endereco & "','" & bairro & "','" & cep & "','" & cidest & "','" & datanasc & "','" & tel & "','" & cel & "','" & email & "','" & obs & "','" & nomecli & "','" & datacred & "','" & tipocred & "','" & valorvenda & "','" & parcela & "')"
                
                TotalCredTemp = TotalCredTemp - 1
                clientetemp = cliente
            End If
            
            VLStrLinha = VLStrLinha + 1
            'RecCli.Close
            RecCred.MoveNext
        Loop
        
        Desconecta
            
        rptCrediarista.Show
    End If
End Sub

Private Sub CmdImprimirCx_Click()
    Screen.MousePointer = vbHourglass
    
    Dim desc As String
    Dim datamov As String
    Dim tipomov As String
    Dim cred As String
    Dim deb As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridCaixa.MaxRows
        
        GridCaixa.Col = 2
        GridCaixa.Row = VLStrLinha
        desc = GridCaixa.Text
        
        GridCaixa.Col = 3
        GridCaixa.Row = VLStrLinha
        datamov = GridCaixa.Text
        
        GridCaixa.Col = 4
        GridCaixa.Row = VLStrLinha
        tipomov = GridCaixa.Text
        
        GridCaixa.Col = 5
        GridCaixa.Row = VLStrLinha
        cred = GridCaixa.Text
        
        GridCaixa.Col = 6
        GridCaixa.Row = VLStrLinha
        deb = GridCaixa.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & desc & "','" & datamov & "','" & tipomov & "','" & cred & "','" & deb & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCaixa.Show

End Sub

Private Sub CmdImprimirEst_Click()
    Screen.MousePointer = vbHourglass
    
    Dim tipoprod As String
    Dim prod As String
    Dim qtdemin As String
    Dim qtdeest As String
    Dim precovenda As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridEstoque.MaxRows
        
        GridEstoque.Col = 1
        GridEstoque.Row = VLStrLinha
        tipoprod = GridEstoque.Text
        
        GridEstoque.Col = 2
        GridEstoque.Row = VLStrLinha
        prod = GridEstoque.Text
        
        GridEstoque.Col = 3
        GridEstoque.Row = VLStrLinha
        qtdemin = GridEstoque.Text
        
        GridEstoque.Col = 4
        GridEstoque.Row = VLStrLinha
        qtdeest = GridEstoque.Text
        
        GridEstoque.Col = 7
        GridEstoque.Row = VLStrLinha
        precovenda = GridEstoque.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & tipoprod & "','" & prod & "','" & qtdemin & "','" & qtdeest & "','" & precovenda & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptEstoque.Show

End Sub

'Private Sub CmdImprimirExt_Click()
'    Dim RecPesq As New ADODB.Recordset
'    Dim RecProd As New ADODB.Recordset
'    Dim CodProdTemp As Integer
'    Dim VLIntCont As Integer
'    Dim VLStrGravar As String
'    Dim VLStrCampo01 As String
'    Dim VLStrCampo02 As String
'    Dim VLStrCampo03 As String
'    Dim VLStrCampo04 As String
'    Dim VLStrCampo05 As String
'    Dim VLStrCampo06 As String
'    Dim VLStrCampo07 As String
'    Dim VLStrCampo08 As String
'    Dim VLStrCampo09 As String
'    Dim VLStrCampo10 As String
'    Dim VLStrCampo11 As String
'    Dim VLStrCampo12 As String
'    Dim VLStrCampo13 As String
'    Dim VLStrCampo14 As String
'    Dim VLStrCampo15 As String
'    Dim VLStrCampo16 As String
'    Dim VLStrCampo17 As String
'    Dim VLStrCampo18 As String
'    Dim VLStrCampo19 As String
'    Dim VLStrCampo20 As String
'    Dim VLStrCampo21 As String
'    Dim VLStrCampo22 As String
'    Dim VLStrCampo23 As String
'    Dim VLStrCampo24 As String
'    Dim VLStrCampo25 As String
'    Dim VLStrCampo26 As String
'    Dim VLStrCampo27 As String
'    Dim VLStrCampo28 As String
'    Dim VLStrCampo29 As String
'    Dim VLStrCampo30 As String
'    Dim VLStrCampo31 As String
'    Dim VLStrCampo32 As String
'    Dim VLStrCampo33 As String
'    Dim VLStrCampo34 As String
'    Dim VLStrCampo35 As String
'    Dim VLStrCampo36 As String
'    Dim VLStrCampo37 As String
'    Dim VLIntCodCredTemp As Integer
'
'    '============ Mala direta ============
'    If OptMala.Value = True Then
'        Conecta
'        StrSql = "Select * from tb_Cliente where 0=0"
'
'        '====== PESQUISAR POR CLIENTE ==========
'        If TxtCliente.Text <> "" Then
'            StrSql = StrSql + " and Nome like '%" & TxtCliente.Text & "%'"
'        End If
'
'        '====== PESQUISAR POR SEXO ==========
'        If CboSexo.Text <> "" Then
'            StrSql = StrSql + " and Sexo='" & CboSexo.Text & "'"
'        End If
'
'        StrSql = StrSql + " order by Nome"
'        RecPesq.Open StrSql, vgCon, 1, 3
'
'        If RecPesq.EOF Then
'            VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
'            TxtCliente.SetFocus
'        Else
'            VLIntCont = 1
'            Do While Not RecPesq.EOF
'                If VLIntCont = 1 Then
'                    VLStrCampo01 = RecPesq.Fields.Item(2).Value
'                    VLStrCampo02 = RecPesq.Fields.Item(4).Value
'                    VLStrCampo03 = RecPesq.Fields.Item(5).Value
'                    VLStrCampo04 = RecPesq.Fields.Item(6).Value
'                    VLStrCampo05 = RecPesq.Fields.Item(7).Value & "/" & RecPesq.Fields.Item(8).Value
'                    VLIntCont = 2
'
'                ElseIf VLIntCont = 2 Then
'                    VLStrCampo06 = RecPesq.Fields.Item(2).Value
'                    VLStrCampo07 = RecPesq.Fields.Item(4).Value
'                    VLStrCampo08 = RecPesq.Fields.Item(5).Value
'                    VLStrCampo09 = RecPesq.Fields.Item(6).Value
'                    VLStrCampo10 = RecPesq.Fields.Item(7).Value & "/" & RecPesq.Fields.Item(8).Value
'                    VLIntCont = 1
'
'                    VLStrGravar = "sim"
'
'                End If
'
'                RecPesq.MoveNext
'
'                If RecPesq.EOF = True Or VLStrGravar = "sim" Then
'                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
'                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10) " & _
'                    "VALUES ('" & VLStrCampo01 & "','" & VLStrCampo02 & "','" & VLStrCampo03 & "','" & VLStrCampo04 & "','" & VLStrCampo05 & "','" & VLStrCampo06 & "','" & VLStrCampo07 & "','" & VLStrCampo08 & "','" & VLStrCampo09 & "','" & VLStrCampo10 & "')"
'
'                    VLStrGravar = ""
'                    VLStrCampo01 = ""
'                    VLStrCampo02 = ""
'                    VLStrCampo03 = ""
'                    VLStrCampo04 = ""
'                    VLStrCampo05 = ""
'                    VLStrCampo06 = ""
'                    VLStrCampo07 = ""
'                    VLStrCampo08 = ""
'                    VLStrCampo09 = ""
'                    VLStrCampo10 = ""
'                End If
'            Loop
'        End If
'        Desconecta
'        rptExtra_Mala.Show
'
'    '============ Cartas de cobrança ============
'    ElseIf OptCob.Value = True Then
'        'If TxtClienteCob.Text = "" And TxtDtVenc1.Text = "" And TxtDtVenc2.Text = "" Then
'        '    VPStrBox = MsgBox("Preencha pelo menos mais um campo", vbInformation, "Pró Ótica 2004 - Informação")
'        '    CboTipoCarta.SetFocus
'        'Else
'            Conecta
''            StrSql = "Select CR.Parcela,CR.DtCred,CR.CodCred,C.Nome " & _
''                     "From tb_Crediario as CR, tb_Crediario_Parcela as P, tb_Cliente as C " & _
''                     "Where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and P.Quitado='não'"
'
'            StrSql = "Select CR.Parcela,CR.CodCred,P.NumParc,P.Vencimento,P.Valor,C.Nome " & _
'                     "From tb_Crediario as CR, tb_Crediario_Parcela as P, tb_Cliente as C " & _
'                     "Where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and P.Quitado='não'"
'
'            '====== PESQUISAR POR CLIENTE ==========
'            If TxtClienteCob.Text <> "" Then
'                StrSql = StrSql + " and C.Nome like '%" & TxtClienteCob.Text & "%'"
'            End If
'
'            '====== PESQUISAR POR DATA DO VENCIMENTO ==========
'            If (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
'                StrSql = StrSql + " and P.Vencimento >=#" & FormataDataUS(TxtDtVenc1.Text) & "# and P.Vencimento <= #" & FormataDataUS(TxtDtVenc2.Text) & "#"
'
'            ElseIf (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text = "" Or TxtDtVenc2.Text = "__/__/____") Then
'                StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc1.Text) & "#"
'
'            ElseIf (TxtDtVenc1.Text = "" Or TxtDtVenc1.Text = "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
'                StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc2.Text) & "#"
'            End If
'
'            StrSql = StrSql + " order by C.Nome,P.Vencimento desc"
'            RecPesq.Open StrSql, vgCon, 1, 3
'
'            If RecPesq.EOF Then
'                VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
'                TxtClienteCob.SetFocus
'            Else
'                Do While Not RecPesq.EOF
'                    VLIntCont = 1
'                    VLIntCodCredTemp = RecPesq!CodCred
'
'                    VLStrCampo01 = RecPesq!nome
'
'                    Do While (RecPesq!CodCred = VLIntCodCredTemp) And (RecPesq.EOF = False)
'                        If VLIntCont = 1 Then
'                            VLStrCampo02 = RecPesq!vencimento
'                            VLStrCampo03 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo04 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 2 Then
'                            VLStrCampo05 = RecPesq!vencimento
'                            VLStrCampo06 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo07 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 3 Then
'                            VLStrCampo08 = RecPesq!vencimento
'                            VLStrCampo09 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo10 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 4 Then
'                            VLStrCampo11 = RecPesq!vencimento
'                            VLStrCampo12 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo13 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 5 Then
'                            VLStrCampo14 = RecPesq!vencimento
'                            VLStrCampo15 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo16 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 6 Then
'                            VLStrCampo17 = RecPesq!vencimento
'                            VLStrCampo18 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo19 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 7 Then
'                            VLStrCampo20 = RecPesq!vencimento
'                            VLStrCampo21 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo22 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 8 Then
'                            VLStrCampo23 = RecPesq!vencimento
'                            VLStrCampo24 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo25 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 9 Then
'                            VLStrCampo26 = RecPesq!vencimento
'                            VLStrCampo27 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo28 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 10 Then
'                            VLStrCampo29 = RecPesq!vencimento
'                            VLStrCampo30 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo31 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 11 Then
'                            VLStrCampo32 = RecPesq!vencimento
'                            VLStrCampo33 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo34 = FormataMoeda(RecPesq!valor)
'                        ElseIf VLIntCont = 12 Then
'                            VLStrCampo35 = RecPesq!vencimento
'                            VLStrCampo36 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
'                            VLStrCampo37 = FormataMoeda(RecPesq!valor)
'                        End If
'
'                        VLIntCont = VLIntCont + 1
'
'                        RecPesq.MoveNext
'
'                        If RecPesq.EOF = True Then
'                            Exit Do
'                        End If
'                    Loop
'
'                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
'                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18,campo19,campo20,campo21,campo22,campo23,campo24,campo25,campo26,campo27,campo28,campo29,campo30,campo31,campo32,campo33,campo34,campo35,campo36,campo37) " & _
'                    "VALUES ('" & VLStrCampo01 & "','" & VLStrCampo02 & "','" & VLStrCampo03 & "','" & VLStrCampo04 & "','" & VLStrCampo05 & "','" & VLStrCampo06 & "','" & VLStrCampo07 & "','" & VLStrCampo08 & "','" & VLStrCampo09 & "','" & VLStrCampo10 & "','" & VLStrCampo11 & "','" & VLStrCampo12 & "','" & VLStrCampo13 & "','" & VLStrCampo14 & "','" & VLStrCampo15 & "','" & VLStrCampo16 & "','" & VLStrCampo17 & "','" & VLStrCampo18 & "','" & VLStrCampo19 & "','" & VLStrCampo20 & "','" & VLStrCampo21 & "','" & VLStrCampo22 & "','" & VLStrCampo23 & "','" & VLStrCampo24 & "','" & VLStrCampo25 & "','" & VLStrCampo26 & "','" & VLStrCampo27 & "','" & VLStrCampo28 & "','" & VLStrCampo29 & "','" & VLStrCampo30 & "','" & VLStrCampo31 & "','" & VLStrCampo32 & "','" & VLStrCampo33 & "','" & VLStrCampo34 & "','" & VLStrCampo35 & "','" & VLStrCampo36 & "','" & VLStrCampo37 & "')"
'
'                Loop
'            End If
'            Desconecta
'
'            If InStr(CboTipoCarta.Text, "simples") <> 0 Then
'                rptExtra_CobrancaSimples.Show
'
'            ElseIf InStr(CboTipoCarta.Text, "amigável") <> 0 Then
'                rptExtra_CobrancaAmigavel.Show
'
'            ElseIf InStr(CboTipoCarta.Text, "último") <> 0 Then
'                rptExtra_CobrancaUltimoAviso.Show
'
'            End If
'
'        'End If
'
'    '============ Folhetos explicativos ============
'    ElseIf OptExplic.Value = True Then
'        If CboFolheto.Text = "" Then
'            VPStrBox = MsgBox("Escolha o folheto", vbInformation, "Pró Ótica 2004 - Informação")
'            CboFolheto.SetFocus
'        Else
'            If InStr(CboFolheto.Text, "catarata") <> 0 Then
'                rptExtra_Folheto_Catarata.Show
'
'            ElseIf InStr(CboFolheto.Text, "óculos") <> 0 Then
'                rptExtra_Folheto_Oculos.Show
'
'            ElseIf InStr(CboFolheto.Text, "lentes") <> 0 Then
'                rptExtra_Folheto_Lentes.Show
'
'            ElseIf InStr(CboFolheto.Text, "glaucoma") <> 0 Then
'                rptExtra_Folheto_Glaucoma.Show
'
'            End If
'        End If
'
'    '============ Aniversariantes ============
'    ElseIf OptNiver.Value = True Then
'        If TxtDtNiver1.Text = "" And TxtDtNiver2.Text = "" Then
'            VPStrBox = MsgBox("Preencha pelo menos um dos campos", vbInformation, "Pró Ótica 2004 - Informação")
'            TxtDtNiver1.SetFocus
'        Else
'            Conecta
'            StrSql = "Select * from tb_Cliente where 0=0"
'
'            '====== PESQUISAR POR DATA DE NASCIMENTO ==========
'            If (TxtDtNiver1.Text <> "" And TxtDtNiver1.Text <> "__/__/____") And (TxtDtNiver2.Text <> "" And TxtDtNiver2.Text <> "__/__/____") Then
'                StrSql = StrSql + " and DtNasc >=#" & FormataDataUS(TxtDtNiver1.Text) & "# and DtNasc <= #" & FormataDataUS(TxtDtNiver2.Text) & "#"
'
'            ElseIf (TxtDtNiver1.Text <> "" And TxtDtNiver1.Text <> "__/__/____") And (TxtDtNiver2.Text = "" Or TxtDtNiver2.Text = "__/__/____") Then
'                StrSql = StrSql + " and DtNasc =#" & FormataDataUS(TxtDtNiver1.Text) & "#"
'
'            ElseIf (TxtDtNiver1.Text = "" Or TxtDtNiver1.Text = "__/__/____") And (TxtDtNiver2.Text <> "" And TxtDtNiver2.Text <> "__/__/____") Then
'                StrSql = StrSql + " DtNasc =#" & FormataDataUS(TxtDtNiver2.Text) & "#"
'            End If
'
'            StrSql = StrSql + " order by Nome,DtNasc"
'            RecPesq.Open StrSql, vgCon, 1, 3
'
'            If RecPesq.EOF Then
'                VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
'                TxtDtNiver1.SetFocus
'            Else
'                Do While Not RecPesq.EOF
'                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
'                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08) " & _
'                    "VALUES ('" & RecPesq.Fields.Item(2).Value & "','" & FormataData(RecPesq.Fields.Item(9).Value) & "','" & RecPesq.Fields.Item(4).Value & "','" & RecPesq.Fields.Item(5).Value & "','" & RecPesq.Fields.Item(6).Value & "','" & RecPesq.Fields.Item(7).Value & "/" & RecPesq.Fields.Item(8).Value & "','" & RecPesq.Fields.Item(10).Value & "','" & RecPesq.Fields.Item(13).Value & "')"
'
'                    RecPesq.MoveNext
'                Loop
'            End If
'            Desconecta
'
'            rptExtra_Niver.Show
'        End If
'
'    '============ Etiquetas para armação ============
'    ElseIf OptEtiqArm.Value = True Then
'        If CboGriffe.Text = "" Then
'            VPStrBox = MsgBox("Selecione a griffe", vbInformation, "Pró Ótica 2004 - Informação")
'            CboGriffe.SetFocus
'        Else
'            Dim VLIntCodGriffe As Long
'            VLIntCodGriffe = Mid(CboGriffe.Text, Len(CboGriffe.Text) - 10)
'
'            Conecta
'            StrSql = "Select CodProd,Cor,Numero,Modelo,TamAro,TamPonte from tb_Produto where CodGriffe=" & VLIntCodGriffe
'            RecPesq.Open StrSql, vgCon, 1, 3
'
'            If RecPesq.EOF Then
'                VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
'                CboGriffe.SetFocus
'            Else
'                Do While Not RecPesq.EOF
'                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
'                    "(campo01) " & _
'                    "VALUES ('" & RecPesq.Fields.Item(0).Value & "/" & RecPesq.Fields.Item(1).Value & "/" & RecPesq.Fields.Item(2).Value & "/" & RecPesq.Fields.Item(3).Value & "/" & RecPesq.Fields.Item(4).Value & "/" & RecPesq.Fields.Item(5).Value & "')"
'
'                    RecPesq.MoveNext
'                Loop
'            End If
'            Desconecta
'
'            rptExtra_Etiqueta.Show
'        End If
'    End If
'End Sub

Private Sub CmdImprimirExt_Click()
    Dim RecPesq As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    Dim RecEstq As New ADODB.Recordset
    Dim CodProdTemp As Integer
    Dim VLIntCont As Integer
    Dim VLStrGravar As String
    Dim VLStrGriffe As String
    Dim VLStrCampo01 As String
    Dim VLStrCampo02 As String
    Dim VLStrCampo03 As String
    Dim VLStrCampo04 As String
    Dim VLStrCampo05 As String
    Dim VLStrCampo06 As String
    Dim VLStrCampo07 As String
    Dim VLStrCampo08 As String
    Dim VLStrCampo09 As String
    Dim VLStrCampo10 As String
    Dim VLStrCampo11 As String
    Dim VLStrCampo12 As String
    Dim VLStrCampo13 As String
    Dim VLStrCampo14 As String
    Dim VLStrCampo15 As String
    Dim VLStrCampo16 As String
    Dim VLStrCampo17 As String
    Dim VLStrCampo18 As String
    Dim VLStrCampo19 As String
    Dim VLStrCampo20 As String
    Dim VLStrCampo21 As String
    Dim VLStrCampo22 As String
    Dim VLStrCampo23 As String
    Dim VLStrCampo24 As String
    Dim VLStrCampo25 As String
    Dim VLStrCampo26 As String
    Dim VLStrCampo27 As String
    Dim VLStrCampo28 As String
    Dim VLStrCampo29 As String
    Dim VLStrCampo30 As String
    Dim VLStrCampo31 As String
    Dim VLStrCampo32 As String
    Dim VLStrCampo33 As String
    Dim VLStrCampo34 As String
    Dim VLStrCampo35 As String
    Dim VLStrCampo36 As String
    Dim VLStrCampo37 As String
    Dim VLIntCodCredTemp As Integer
    
    '============ Mala direta ============
    If OptMala.Value = True Then
        Conecta
        StrSql = "Select * from tb_Cliente where 0=0"
                
        '====== PESQUISAR POR CLIENTE ==========
        If TxtCliente.Text <> "" Then
            StrSql = StrSql + " and Nome like '%" & TxtCliente.Text & "%'"
        End If
                
        '====== PESQUISAR POR SEXO ==========
        If CboSexo.Text <> "" Then
            StrSql = StrSql + " and Sexo='" & CboSexo.Text & "'"
        End If
        
        '====== PESQUISAR POR DATA DE NASCIMENTO ==========
        If (TxtDtNiverCli1.Text <> "" And TxtDtNiverCli1.Text <> "__/__/____") And (TxtDtNiverCli2.Text <> "" And TxtDtNiverCli2.Text <> "__/__/____") Then
            StrSql = StrSql + " and DtNasc >=#" & FormataDataUS(TxtDtNiverCli1.Text) & "# and DtNasc <= #" & FormataDataUS(TxtDtNiverCli2.Text) & "#"
        
        ElseIf (TxtDtNiverCli1.Text <> "" And TxtDtNiverCli1.Text <> "__/__/____") And (TxtDtNiverCli2.Text = "" Or TxtDtNiverCli2.Text = "__/__/____") Then
            StrSql = StrSql + " and DtNasc =#" & FormataDataUS(TxtDtNiverCli1.Text) & "#"
        
        ElseIf (TxtDtNiverCli1.Text = "" Or TxtDtNiverCli1.Text = "__/__/____") And (TxtDtNiverCli2.Text <> "" And TxtDtNiverCli2.Text <> "__/__/____") Then
            StrSql = StrSql + " DtNasc =#" & FormataDataUS(TxtDtNiverCli2.Text) & "#"
        End If
        
        StrSql = StrSql + " order by Nome"
        RecPesq.Open StrSql, vgCon, 1, 3
        
        If RecPesq.EOF Then
            VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
            TxtCliente.SetFocus
        Else
            VLIntCont = 1
            Do While Not RecPesq.EOF
                If VLIntCont = 1 Then
                    VLStrCampo01 = RecPesq.Fields.Item(2).Value
                    VLStrCampo02 = RecPesq.Fields.Item(4).Value
                    VLStrCampo03 = RecPesq.Fields.Item(5).Value
                    VLStrCampo04 = RecPesq.Fields.Item(6).Value
                    VLStrCampo05 = RecPesq.Fields.Item(7).Value & "/" & RecPesq.Fields.Item(8).Value
                    VLIntCont = 2
                    
                ElseIf VLIntCont = 2 Then
                    VLStrCampo06 = RecPesq.Fields.Item(2).Value
                    VLStrCampo07 = RecPesq.Fields.Item(4).Value
                    VLStrCampo08 = RecPesq.Fields.Item(5).Value
                    VLStrCampo09 = RecPesq.Fields.Item(6).Value
                    VLStrCampo10 = RecPesq.Fields.Item(7).Value & "/" & RecPesq.Fields.Item(8).Value
                    VLIntCont = 1
                    
                    VLStrGravar = "sim"
                    
                End If
                
                RecPesq.MoveNext
                
                If RecPesq.EOF = True Or VLStrGravar = "sim" Then
                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10) " & _
                    "VALUES ('" & VLStrCampo01 & "','" & VLStrCampo02 & "','" & VLStrCampo03 & "','" & VLStrCampo04 & "','" & VLStrCampo05 & "','" & VLStrCampo06 & "','" & VLStrCampo07 & "','" & VLStrCampo08 & "','" & VLStrCampo09 & "','" & VLStrCampo10 & "')"
                
                    VLStrGravar = ""
                    VLStrCampo01 = ""
                    VLStrCampo02 = ""
                    VLStrCampo03 = ""
                    VLStrCampo04 = ""
                    VLStrCampo05 = ""
                    VLStrCampo06 = ""
                    VLStrCampo07 = ""
                    VLStrCampo08 = ""
                    VLStrCampo09 = ""
                    VLStrCampo10 = ""
                End If
            Loop
        End If
        Desconecta
        rptExtra_Mala.Show
    
    '============ Cartas de cobrança ============
    ElseIf OptCob.Value = True Then
        Conecta
                
        StrSql = "Select CR.Parcela,CR.CodCred,P.NumParc,P.Vencimento,P.Valor,C.Nome " & _
                 "From tb_Crediario as CR, tb_Crediario_Parcela as P, tb_Cliente as C " & _
                 "Where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and P.Quitado='não'"
                
        '====== PESQUISAR POR CLIENTE ==========
        If TxtClienteCob.Text <> "" Then
            StrSql = StrSql + " and C.Nome like '%" & TxtClienteCob.Text & "%'"
        End If
                
        '====== PESQUISAR POR DATA DO VENCIMENTO ==========
        If (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
            StrSql = StrSql + " and P.Vencimento >=#" & FormataDataUS(TxtDtVenc1.Text) & "# and P.Vencimento <= #" & FormataDataUS(TxtDtVenc2.Text) & "#"
        
        ElseIf (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text = "" Or TxtDtVenc2.Text = "__/__/____") Then
            StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc1.Text) & "#"
        
        ElseIf (TxtDtVenc1.Text = "" Or TxtDtVenc1.Text = "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
            StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc2.Text) & "#"
        End If
                
        StrSql = StrSql + " order by C.Nome,P.Vencimento desc"
        RecPesq.Open StrSql, vgCon, 1, 3
        
        If RecPesq.EOF Then
            VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
            TxtClienteCob.SetFocus
        Else
            Do While Not RecPesq.EOF
                VLIntCont = 1
                VLIntCodCredTemp = RecPesq!CodCred
                
                VLStrCampo01 = RecPesq!nome
                
                Do While (RecPesq!CodCred = VLIntCodCredTemp) And (RecPesq.EOF = False)
                    If VLIntCont = 1 Then
                        VLStrCampo02 = RecPesq!vencimento
                        VLStrCampo03 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo04 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 2 Then
                        VLStrCampo05 = RecPesq!vencimento
                        VLStrCampo06 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo07 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 3 Then
                        VLStrCampo08 = RecPesq!vencimento
                        VLStrCampo09 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo10 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 4 Then
                        VLStrCampo11 = RecPesq!vencimento
                        VLStrCampo12 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo13 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 5 Then
                        VLStrCampo14 = RecPesq!vencimento
                        VLStrCampo15 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo16 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 6 Then
                        VLStrCampo17 = RecPesq!vencimento
                        VLStrCampo18 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo19 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 7 Then
                        VLStrCampo20 = RecPesq!vencimento
                        VLStrCampo21 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo22 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 8 Then
                        VLStrCampo23 = RecPesq!vencimento
                        VLStrCampo24 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo25 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 9 Then
                        VLStrCampo26 = RecPesq!vencimento
                        VLStrCampo27 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo28 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 10 Then
                        VLStrCampo29 = RecPesq!vencimento
                        VLStrCampo30 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo31 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 11 Then
                        VLStrCampo32 = RecPesq!vencimento
                        VLStrCampo33 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo34 = FormataMoeda(RecPesq!valor)
                    ElseIf VLIntCont = 12 Then
                        VLStrCampo35 = RecPesq!vencimento
                        VLStrCampo36 = FormataNum(RecPesq!NumParc) & "/" & FormataNum(RecPesq!parcela)
                        VLStrCampo37 = FormataMoeda(RecPesq!valor)
                    End If
                    
                    VLIntCont = VLIntCont + 1
                    
                    RecPesq.MoveNext
                    
                    If RecPesq.EOF = True Then
                        Exit Do
                    End If
                Loop
                
                vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18,campo19,campo20,campo21,campo22,campo23,campo24,campo25,campo26,campo27,campo28,campo29,campo30,campo31,campo32,campo33,campo34,campo35,campo36,campo37) " & _
                "VALUES ('" & VLStrCampo01 & "','" & VLStrCampo02 & "','" & VLStrCampo03 & "','" & VLStrCampo04 & "','" & VLStrCampo05 & "','" & VLStrCampo06 & "','" & VLStrCampo07 & "','" & VLStrCampo08 & "','" & VLStrCampo09 & "','" & VLStrCampo10 & "','" & VLStrCampo11 & "','" & VLStrCampo12 & "','" & VLStrCampo13 & "','" & VLStrCampo14 & "','" & VLStrCampo15 & "','" & VLStrCampo16 & "','" & VLStrCampo17 & "','" & VLStrCampo18 & "','" & VLStrCampo19 & "','" & VLStrCampo20 & "','" & VLStrCampo21 & "','" & VLStrCampo22 & "','" & VLStrCampo23 & "','" & VLStrCampo24 & "','" & VLStrCampo25 & "','" & VLStrCampo26 & "','" & VLStrCampo27 & "','" & VLStrCampo28 & "','" & VLStrCampo29 & "','" & VLStrCampo30 & "','" & VLStrCampo31 & "','" & VLStrCampo32 & "','" & VLStrCampo33 & "','" & VLStrCampo34 & "','" & VLStrCampo35 & "','" & VLStrCampo36 & "','" & VLStrCampo37 & "')"
                
            Loop
        End If
        Desconecta
        
        If InStr(CboTipoCarta.Text, "simples") <> 0 Then
            rptExtra_CobrancaSimples.Show
            
        ElseIf InStr(CboTipoCarta.Text, "amigável") <> 0 Then
            rptExtra_CobrancaAmigavel.Show
        
        ElseIf InStr(CboTipoCarta.Text, "último") <> 0 Then
            rptExtra_CobrancaUltimoAviso.Show
        
        End If
    
    '============ Folhetos explicativos ============
    ElseIf OptExplic.Value = True Then
        If CboFolheto.Text = "" Then
            VPStrBox = MsgBox("Escolha o folheto que deseja imprimir", vbInformation, "Pró Ótica 2004 - Informação")
            CboFolheto.SetFocus
        Else
            If InStr(CboFolheto.Text, "catarata") <> 0 Then
                rptExtra_Folheto_Catarata.Show
                
            ElseIf InStr(CboFolheto.Text, "óculos") <> 0 Then
                rptExtra_Folheto_Oculos.Show
            
            ElseIf InStr(CboFolheto.Text, "coloridas") <> 0 Then
                rptExtra_Folheto_Lentes.Show
            
            ElseIf InStr(CboFolheto.Text, "Glaucoma") <> 0 Then
                rptExtra_Folheto_Glaucoma.Show
            
            ElseIf InStr(CboFolheto.Text, "contato") <> 0 Then
                rptExtra_Folheto_LentesContato.Show
            
            ElseIf InStr(CboFolheto.Text, "oculares") <> 0 Then
                rptExtra_Folheto_Oculares.Show
            
            End If
        End If
    
    '============ Aniversariantes ============
    ElseIf OptNiver.Value = True Then
        If TxtDia1.Text = "" And TxtDia2.Text = "" And TxtMes1.Text = "" And TxtMes2.Text = "" Then
            VPStrBox = MsgBox("Preencha o intervalo de dias e meses que se deseja imprimir", vbInformation, "Pró Ótica 2004 - Informação")
            TxtDia1.SetFocus
        
        ElseIf TxtDia1.Text = "" And TxtDia2.Text = "" Then
            VPStrBox = MsgBox("Preencha o intervalo de dias que se deseja imprimir", vbInformation, "Pró Ótica 2004 - Informação")
            TxtDia1.SetFocus
        
        ElseIf TxtMes1.Text = "" And TxtMes2.Text = "" Then
            VPStrBox = MsgBox("Preencha o intervalo de meses que se deseja imprimir", vbInformation, "Pró Ótica 2004 - Informação")
            TxtMes1.SetFocus
        
        Else
            Conecta
            StrSql = "Select * from tb_Cliente where 0=0"
                    
            '====== PESQUISAR POR DIA ==========
            If TxtDia1.Text <> "" And TxtDia2.Text <> "" Then
               StrSql = StrSql + " and Datepart('D',DtNasc) >=" & TxtDia1.Text & " and Datepart('D',DtNasc) <= " & TxtDia2.Text & ""
            
            ElseIf TxtDia1.Text <> "" And TxtDia2.Text = "" Then
               StrSql = StrSql + " and Datepart('D',DtNasc) =" & TxtDia1.Text & ""
            
            ElseIf TxtDia1.Text = "" And TxtDia2.Text <> "" Then
               StrSql = StrSql + " and Datepart('D',DtNasc) =" & TxtDia2.Text & ""
            
            End If
            
            '====== PESQUISAR POR MÊS ==========
            If TxtMes1.Text <> "" And TxtMes2.Text <> "" Then
               StrSql = StrSql + " and Datepart('M',DtNasc) >=" & TxtMes1.Text & " and Datepart('M',DtNasc) <= " & TxtMes2.Text & ""
            
            ElseIf TxtMes1.Text <> "" And TxtMes2.Text = "" Then
               StrSql = StrSql + " and Datepart('M',DtNasc) =" & TxtMes1.Text & ""
            
            ElseIf TxtMes1.Text = "" And TxtMes2.Text <> "" Then
               StrSql = StrSql + " and Datepart('M',DtNasc) =" & TxtMes2.Text & ""
            
            End If
            
            StrSql = StrSql + " order by Datepart('M',DtNasc), Datepart('D',DtNasc)"
            RecPesq.Open StrSql, vgCon, 1, 3
            
            If RecPesq.EOF Then
                VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
                TxtDia1.SetFocus
            Else
                Do While Not RecPesq.EOF
                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08) " & _
                    "VALUES ('" & RecPesq!nome & "','" & FormataData(RecPesq!dtnasc) & "','" & RecPesq!endereco & "','" & RecPesq!bairro & "','" & RecPesq!cep & "','" & RecPesq!cidade & "/" & RecPesq!Estado & "','" & RecPesq!telefone & "','" & RecPesq!email & "')"
                    
                    RecPesq.MoveNext
                Loop
            End If
            Desconecta
            
            rptExtra_Niver.Show
        End If
    
    '============ Etiquetas para armação ============
    ElseIf OptEtiqArm.Value = True Then
        If CboGriffe.Text = "" Then
            VPStrBox = MsgBox("Selecione a griffe", vbInformation, "Pró Ótica 2004 - Informação")
            CboGriffe.SetFocus
        Else
            Dim VLIntCodGriffe As Long
            Dim VLStrProduto As String
            
            VLIntCodGriffe = Mid(CboGriffe.Text, Len(CboGriffe.Text) - 10)
            VLStrGriffe = Trim(Mid(CboGriffe.Text, 1, Len(CboGriffe.Text) - 10))
            
            Conecta
            StrSql = "Select CodProd,Cor,Numero,Modelo,TamAro,TamPonte from tb_Produto where CodGriffe=" & VLIntCodGriffe
            RecPesq.Open StrSql, vgCon, 1, 3
            
            If RecPesq.EOF Then
                VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
                CboGriffe.SetFocus
            Else
            
                VLIntCont = 1
                Do While Not RecPesq.EOF
                    StrSql = "Select PrecoVenda from tb_Estoque where CodProd=" & RecPesq!CodProd
                    RecEstq.Open StrSql, vgCon, 1, 3
                    
                    If VLIntCont = 1 Then
                        VLStrCampo01 = VLStrGriffe & " / C" & RecPesq!cor & " / N" & RecPesq!Numero & " / M" & RecPesq!modelo & " / A" & RecPesq!TamAro & " / P" & RecPesq!TamPonte
                        VLStrCampo02 = RecEstq!precovenda
                        VLIntCont = 2
                        
                    ElseIf VLIntCont = 2 Then
                        VLStrCampo03 = VLStrGriffe & " / C" & RecPesq!cor & " / N" & RecPesq!Numero & " / M" & RecPesq!modelo & " / A" & RecPesq!TamAro & " / P" & RecPesq!TamPonte
                        VLStrCampo04 = RecEstq!precovenda
                        VLIntCont = 3
                        
                    ElseIf VLIntCont = 3 Then
                        VLStrCampo05 = VLStrGriffe & " / C" & RecPesq!cor & " / N" & RecPesq!Numero & " / M" & RecPesq!modelo & " / A" & RecPesq!TamAro & " / P" & RecPesq!TamPonte
                        VLStrCampo06 = RecEstq!precovenda
                        VLIntCont = 4
                        
                    ElseIf VLIntCont = 4 Then
                        VLStrCampo07 = VLStrGriffe & "/C" & RecPesq!cor & "/N" & RecPesq!Numero & "/M" & RecPesq!modelo & "/A" & RecPesq!TamAro & "/P" & RecPesq!TamPonteVLStrGriffe & " / C" & RecPesq!cor & " / N" & RecPesq!Numero & " / M" & RecPesq!modelo & " / A" & RecPesq!TamAro & " / P" & RecPesq!TamPonte
                        VLStrCampo08 = RecEstq!precovenda
                        VLIntCont = 1
                        
                        VLStrGravar = "sim"
                        
                    End If
                    
                    RecEstq.Close
                    
                    RecPesq.MoveNext
                    
                    If RecPesq.EOF = True Or VLStrGravar = "sim" Then
                        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08) " & _
                        "VALUES ('" & VLStrCampo01 & "','" & VLStrCampo02 & "','" & VLStrCampo03 & "','" & VLStrCampo04 & "','" & VLStrCampo05 & "','" & VLStrCampo06 & "','" & VLStrCampo07 & "','" & VLStrCampo08 & "')"
                    
                        VLStrGravar = ""
                        VLStrCampo01 = ""
                        VLStrCampo02 = ""
                        VLStrCampo03 = ""
                        VLStrCampo04 = ""
                        VLStrCampo05 = ""
                        VLStrCampo06 = ""
                        VLStrCampo07 = ""
                        VLStrCampo08 = ""
                    End If
                Loop
            
            End If
            Desconecta
            
            rptExtra_Etiqueta.Show
        End If
    
    '============ Proposta de crédito ============
    ElseIf OptPropCred.Value = True Then
        
        If TxtClienteProp.Text = "" And TxtCredstaProp.Text = "" And (TxtDtVendaProp01.Text = "" Or TxtDtVendaProp01.Text = "__/__/____") And (TxtDtVendaProp02.Text = "" Or TxtDtVendaProp02.Text = "__/__/____") Then
            VPStrBox = MsgBox("Preencha pelo menos um dos campos para a pesquisa.", vbInformation, "Pró Ótica 2004 - Informação")
            TxtClienteProp.SetFocus
        Else
            Conecta
            
            StrSql = "SELECT CR.CodCred FROM tb_Crediario as CR WHERE 0=0"
                    
            '====== PESQUISAR POR CLIENTE ==========
            If TxtClienteProp.Text <> "" Then
                StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",C.CodCli " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Cliente as C " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and C.CodCli=CR.CodCli and C.Nome like '%" & TxtClienteProp.Text & "%'"
                'StrSql = StrSql + " and C.Nome like '%" & TxtClienteCob.Text & "%'"
            End If
            
            '====== PESQUISAR POR CREDIARISTA ==========
            If TxtCredstaProp.Text <> "" Then
                'StrSql = StrSql + " and C.Nome like '%" & TxtClienteCob.Text & "%'"
                StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",CS.CodCredsta " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediarista as CS " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CS.CodCredsta=CR.CodCredsta and CS.Nome like '%" & TxtCredstaProp.Text & "%'"
            End If
            
            '====== PESQUISAR POR DATA DO CREDIÁRIO ==========
            If (TxtDtVendaProp01.Text <> "" And TxtDtVendaProp01.Text <> "__/__/____") And (TxtDtVendaProp02.Text <> "" And TxtDtVendaProp02.Text <> "__/__/____") Then
                StrSql = StrSql + " and CR.DtCred >=#" & FormataDataUS(TxtDtVendaProp01.Text) & "# and CR.DtCred <= #" & FormataDataUS(TxtDtVendaProp02.Text) & "#"
            
            ElseIf (TxtDtVendaProp01.Text <> "" And TxtDtVendaProp01.Text <> "__/__/____") And (TxtDtVendaProp02.Text = "" Or TxtDtVendaProp02.Text = "__/__/____") Then
                StrSql = StrSql + " and CR.DtCred =#" & FormataDataUS(TxtDtVendaProp01.Text) & "#"
            
            ElseIf (TxtDtVendaProp01.Text = "" Or TxtDtVendaProp01.Text = "__/__/____") And (TxtDtVendaProp02.Text <> "" And TxtDtVendaProp02.Text <> "__/__/____") Then
                StrSql = StrSql + " and CR.DtCred =#" & FormataDataUS(TxtDtVendaProp02.Text) & "#"
            End If
                    
            StrSql = StrSql + " order by CR.CodCred desc"
            RecPesq.Open StrSql, vgCon, 1, 3
            
            If RecPesq.EOF Then
                Desconecta
                VPStrBox = MsgBox("Pesquisa sem resultados", vbInformation, "Pró Ótica 2004 - Informação")
                TxtClienteProp.SetFocus
            Else
                VGIntPropCodCred = RecPesq!CodCred
                Desconecta
                VGStrAssinaturaProposta = "extraproposta"
                FrmAssinaturaOrc.Show
            End If
        End If
    End If
End Sub

Private Sub CmdImprimirForn_Click()
    Screen.MousePointer = vbHourglass
    
    Dim forn As String
    Dim tipo As String
    Dim endereco As String
    Dim bairro As String
    Dim cep As String
    Dim cidest As String
    Dim cnpj As String
    Dim email As String
    Dim resp As String
    Dim tel As String
    Dim cel As String
    Dim obs As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridFornecedor.MaxRows
        
        GridFornecedor.Col = 1
        GridFornecedor.Row = VLStrLinha
        forn = GridFornecedor.Text
        
        GridFornecedor.Col = 2
        GridFornecedor.Row = VLStrLinha
        tipo = GridFornecedor.Text
        
        GridFornecedor.Col = 3
        GridFornecedor.Row = VLStrLinha
        endereco = GridFornecedor.Text
        
        GridFornecedor.Col = 4
        GridFornecedor.Row = VLStrLinha
        bairro = GridFornecedor.Text
        
        GridFornecedor.Col = 5
        GridFornecedor.Row = VLStrLinha
        cep = GridFornecedor.Text
        
        GridFornecedor.Col = 6
        GridFornecedor.Row = VLStrLinha
        cidest = GridFornecedor.Text
        
        GridFornecedor.Col = 7
        GridFornecedor.Row = VLStrLinha
        cidest = cidest & "/" & GridFornecedor.Text
        
        GridFornecedor.Col = 8
        GridFornecedor.Row = VLStrLinha
        cnpj = GridFornecedor.Text
        
        GridFornecedor.Col = 9
        GridFornecedor.Row = VLStrLinha
        email = GridFornecedor.Text
        
        GridFornecedor.Col = 10
        GridFornecedor.Row = VLStrLinha
        resp = GridFornecedor.Text
        
        GridFornecedor.Col = 11
        GridFornecedor.Row = VLStrLinha
        tel = GridFornecedor.Text
        
        GridFornecedor.Col = 12
        GridFornecedor.Row = VLStrLinha
        cel = GridFornecedor.Text
        
        GridFornecedor.Col = 13
        GridFornecedor.Row = VLStrLinha
        obs = GridFornecedor.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12) " & _
        "VALUES ('" & forn & "','" & tipo & "','" & endereco & "','" & bairro & "','" & cep & "','" & cidest & " ','" & cnpj & "','" & email & "','" & resp & "','" & tel & "','" & cel & "','" & obs & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptFornecedor.Show

End Sub

Private Sub CmdImprimirMed_Click()
    Screen.MousePointer = vbHourglass
    
    Dim nome As String
    Dim clicons As String
    Dim crm As String
    Dim endereco As String
    Dim bairro As String
    Dim cep As String
    Dim cidest As String
    Dim datanasc As String
    Dim tel As String
    Dim cel As String
    Dim cpf As String
    Dim email As String
    Dim obs As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridMedico.MaxRows
        
        GridMedico.Col = 1
        GridMedico.Row = VLStrLinha
        nome = GridMedico.Text
        
        GridMedico.Col = 2
        GridMedico.Row = VLStrLinha
        clicons = GridMedico.Text
        
        GridMedico.Col = 3
        GridMedico.Row = VLStrLinha
        crm = GridMedico.Text
        
        GridMedico.Col = 4
        GridMedico.Row = VLStrLinha
        endereco = GridMedico.Text
        
        GridMedico.Col = 5
        GridMedico.Row = VLStrLinha
        bairro = GridMedico.Text
        
        GridMedico.Col = 6
        GridMedico.Row = VLStrLinha
        cep = GridMedico.Text
        
        GridMedico.Col = 7
        GridMedico.Row = VLStrLinha
        cidest = GridMedico.Text
        
        GridMedico.Col = 8
        GridMedico.Row = VLStrLinha
        cidest = cidest & "/" & GridMedico.Text
        
        GridMedico.Col = 9
        GridMedico.Row = VLStrLinha
        datanasc = GridMedico.Text
        
        GridMedico.Col = 10
        GridMedico.Row = VLStrLinha
        tel = GridMedico.Text
        
        GridMedico.Col = 11
        GridMedico.Row = VLStrLinha
        cel = GridMedico.Text
        
        GridMedico.Col = 12
        GridMedico.Row = VLStrLinha
        cpf = GridMedico.Text
        
        GridMedico.Col = 13
        GridMedico.Row = VLStrLinha
        email = GridMedico.Text
        
        GridMedico.Col = 14
        GridMedico.Row = VLStrLinha
        obs = GridMedico.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13) " & _
        "VALUES ('" & nome & "','" & clicons & "','" & crm & "','" & endereco & "','" & bairro & "','" & cep & "','" & cidest & "','" & datanasc & "','" & tel & "','" & cel & "','" & cpf & "','" & email & "','" & obs & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptMedico.Show

End Sub

Private Sub CmdImprimirOrc_Click()
    Screen.MousePointer = vbHourglass
    
    Dim data As String
    Dim vendedor As String
    Dim Armacao As String
    Dim valorarm As String
    Dim lente As String
    Dim valorlente As String
    Dim lentec As String
    Dim valorlentec As String
    Dim outro As String
    Dim valoroutro As String
    Dim totalvenda As String
    Dim parcelado As String
    Dim entrada As String
    Dim valorparc As String
    Dim valorprazo As String
    Dim validade As String
    Dim obs As String
    Dim cliente As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridOrcamento.MaxRows
        
        GridOrcamento.Col = 1
        GridOrcamento.Row = VLStrLinha
        data = GridOrcamento.Text
        
        GridOrcamento.Col = 2
        GridOrcamento.Row = VLStrLinha
        vendedor = GridOrcamento.Text
        
        GridOrcamento.Col = 3
        GridOrcamento.Row = VLStrLinha
        cliente = GridOrcamento.Text
        
        GridOrcamento.Col = 5
        GridOrcamento.Row = VLStrLinha
        Armacao = GridOrcamento.Text
        
        GridOrcamento.Col = 6
        GridOrcamento.Row = VLStrLinha
        valorarm = GridOrcamento.Text
        
        GridOrcamento.Col = 7
        GridOrcamento.Row = VLStrLinha
        lente = GridOrcamento.Text
        
        GridOrcamento.Col = 8
        GridOrcamento.Row = VLStrLinha
        valorlente = GridOrcamento.Text
        
        GridOrcamento.Col = 9
        GridOrcamento.Row = VLStrLinha
        lentec = GridOrcamento.Text
        
        GridOrcamento.Col = 10
        GridOrcamento.Row = VLStrLinha
        valorlentec = GridOrcamento.Text
        
        GridOrcamento.Col = 11
        GridOrcamento.Row = VLStrLinha
        outro = GridOrcamento.Text
        
        GridOrcamento.Col = 12
        GridOrcamento.Row = VLStrLinha
        valoroutro = GridOrcamento.Text
        
        GridOrcamento.Col = 13
        GridOrcamento.Row = VLStrLinha
        totalvenda = GridOrcamento.Text
        
        GridOrcamento.Col = 14
        GridOrcamento.Row = VLStrLinha
        parcelado = Mid(GridOrcamento.Text, 1, 2)
        
        GridOrcamento.Col = 15
        GridOrcamento.Row = VLStrLinha
        entrada = GridOrcamento.Text
        
        GridOrcamento.Col = 16
        GridOrcamento.Row = VLStrLinha
        valorparc = GridOrcamento.Text
        
        GridOrcamento.Col = 17
        GridOrcamento.Row = VLStrLinha
        valorprazo = GridOrcamento.Text
        
        GridOrcamento.Col = 18
        GridOrcamento.Row = VLStrLinha
        validade = GridOrcamento.Text
        
        GridOrcamento.Col = 19
        GridOrcamento.Row = VLStrLinha
        obs = GridOrcamento.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18) " & _
        "VALUES ('" & data & "','" & vendedor & "','" & Armacao & "','" & valorarm & "','" & lente & "','" & valorlente & "','" & lentec & "','" & valorlentec & "','" & outro & "','" & valoroutro & "','" & totalvenda & "','" & parcelado & "','" & entrada & "','" & valorparc & "','" & valorprazo & "','" & validade & "','" & obs & "','" & cliente & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptOrcamento.Show

End Sub

Private Sub CmdImprimirProd_Click()
    Screen.MousePointer = vbHourglass
    
    Dim forn As String
    Dim prod As String
    Dim Griffe As String
    Dim cor As String
    Dim num As String
    Dim modelo As String
    Dim aro As String
    Dim ponte As String
    Dim lente As String
    Dim chave As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridProduto.MaxRows
        
        GridProduto.Col = 1
        GridProduto.Row = VLStrLinha
        forn = GridProduto.Text
        
        GridProduto.Col = 2
        GridProduto.Row = VLStrLinha
        prod = GridProduto.Text
        
        GridProduto.Col = 3
        GridProduto.Row = VLStrLinha
        Griffe = GridProduto.Text
        
        GridProduto.Col = 4
        GridProduto.Row = VLStrLinha
        cor = GridProduto.Text
        
        GridProduto.Col = 5
        GridProduto.Row = VLStrLinha
        num = GridProduto.Text
        
        GridProduto.Col = 6
        GridProduto.Row = VLStrLinha
        modelo = GridProduto.Text
        
        GridProduto.Col = 7
        GridProduto.Row = VLStrLinha
        aro = GridProduto.Text
        
        GridProduto.Col = 8
        GridProduto.Row = VLStrLinha
        ponte = GridProduto.Text
        
        GridProduto.Col = 9
        GridProduto.Row = VLStrLinha
        lente = GridProduto.Text
        
        GridProduto.Col = 10
        GridProduto.Row = VLStrLinha
        chave = GridProduto.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10) " & _
        "VALUES ('" & forn & "','" & prod & "','" & Griffe & "','" & cor & "','" & num & "','" & modelo & "','" & aro & "','" & ponte & "','" & lente & "','" & chave & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptProduto.Show

End Sub

Private Sub CmdImprimirRec_Click()
    Screen.MousePointer = vbHourglass
    
    Dim cliente As String
    Dim medico As String
    Dim datarec As String
    Dim LODEsf As String
    Dim LODCil As String
    Dim LODEixo As String
    Dim LOEEsf As String
    Dim LOECil As String
    Dim LOEEixo As String
    Dim PODEsf As String
    Dim PODCil As String
    Dim PODEixo As String
    Dim POEEsf As String
    Dim POECil As String
    Dim POEEixo As String
    Dim ODDNP As String
    Dim OEDNP As String
    Dim ODAlt As String
    Dim OEAlt As String
    Dim ODAdicao As String
    Dim OEAdicao As String
    Dim AOAdicao As String
    Dim obs As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridReceita.MaxRows
        
        GridReceita.Col = 1
        GridReceita.Row = VLStrLinha
        cliente = GridReceita.Text
        
        GridReceita.Col = 2
        GridReceita.Row = VLStrLinha
        medico = GridReceita.Text
        
        GridReceita.Col = 3
        GridReceita.Row = VLStrLinha
        datarec = GridReceita.Text
        
        GridReceita.Col = 4
        GridReceita.Row = VLStrLinha
        LODEsf = GridReceita.Text
        
        GridReceita.Col = 5
        GridReceita.Row = VLStrLinha
        LODCil = GridReceita.Text
        
        GridReceita.Col = 6
        GridReceita.Row = VLStrLinha
        LODEixo = GridReceita.Text
        
        GridReceita.Col = 7
        GridReceita.Row = VLStrLinha
        LOEEsf = GridReceita.Text
        
        GridReceita.Col = 8
        GridReceita.Row = VLStrLinha
        LOECil = GridReceita.Text
        
        GridReceita.Col = 9
        GridReceita.Row = VLStrLinha
        LOEEixo = GridReceita.Text
        
        GridReceita.Col = 10
        GridReceita.Row = VLStrLinha
        PODEsf = GridReceita.Text
        
        GridReceita.Col = 11
        GridReceita.Row = VLStrLinha
        PODCil = GridReceita.Text
        
        GridReceita.Col = 12
        GridReceita.Row = VLStrLinha
        PODEixo = GridReceita.Text
        
        GridReceita.Col = 13
        GridReceita.Row = VLStrLinha
        POEEsf = GridReceita.Text
        
        GridReceita.Col = 14
        GridReceita.Row = VLStrLinha
        POECil = GridReceita.Text
        
        GridReceita.Col = 15
        GridReceita.Row = VLStrLinha
        POEEixo = GridReceita.Text
        
        GridReceita.Col = 16
        GridReceita.Row = VLStrLinha
        ODDNP = GridReceita.Text
        
        GridReceita.Col = 17
        GridReceita.Row = VLStrLinha
        OEDNP = GridReceita.Text
        
        GridReceita.Col = 18
        GridReceita.Row = VLStrLinha
        ODAlt = GridReceita.Text
        
        GridReceita.Col = 19
        GridReceita.Row = VLStrLinha
        OEAlt = GridReceita.Text
        
        GridReceita.Col = 20
        GridReceita.Row = VLStrLinha
        ODAdicao = GridReceita.Text
        
        GridReceita.Col = 21
        GridReceita.Row = VLStrLinha
        OEAdicao = GridReceita.Text
        
        GridReceita.Col = 22
        GridReceita.Row = VLStrLinha
        AOAdicao = GridReceita.Text
        
        GridReceita.Col = 23
        GridReceita.Row = VLStrLinha
        obs = GridReceita.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18,campo19,campo20,campo21,campo22,campo23) " & _
        "VALUES ('" & cliente & "','" & medico & "','" & datarec & "','" & LODEsf & "','" & LODCil & "','" & LODEixo & "','" & LOEEsf & "','" & LOECil & "','" & LOEEixo & "','" & PODEsf & "','" & PODCil & "','" & PODEixo & "','" & POEEsf & "','" & POECil & "','" & POEEixo & "','" & ODDNP & "','" & OEDNP & "','" & ODAlt & "','" & OEAlt & "','" & ODAdicao & "','" & OEAdicao & "','" & AOAdicao & "','" & obs & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
    
    rptReceita.Show
End Sub

Private Sub CmdImprimirVenda_Click()
    Screen.MousePointer = vbHourglass
    
    Dim RecVenda As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    
    Dim codvenda As Long
    Dim cliente As String
    Dim vendedor As String
    Dim datavenda As String
    Dim valorvenda As String
    Dim desconto As String
    Dim tipovenda As String
    Dim TipoPagto As String
    Dim tipoprod01 As String
    Dim descrprod01 As String
    Dim tipoprod02 As String
    Dim descrprod02 As String
    Dim tipoprod03 As String
    Dim descrprod03 As String
    Dim tipoprod04 As String
    Dim descrprod04 As String
    Dim tipoprod05 As String
    Dim descrprod05 As String
    Dim tipoprod06 As String
    Dim descrprod06 As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridVenda.MaxRows
        
        GridVenda.Col = 1
        GridVenda.Row = VLStrLinha
        cliente = GridVenda.Text
        
        GridVenda.Col = 2
        GridVenda.Row = VLStrLinha
        vendedor = GridVenda.Text
        
        GridVenda.Col = 3
        GridVenda.Row = VLStrLinha
        datavenda = GridVenda.Text
        
        GridVenda.Col = 4
        GridVenda.Row = VLStrLinha
        valorvenda = GridVenda.Text
        
        GridVenda.Col = 5
        GridVenda.Row = VLStrLinha
        desconto = GridVenda.Text
        
        GridVenda.Col = 6
        GridVenda.Row = VLStrLinha
        tipovenda = GridVenda.Text
        
        GridVenda.Col = 7
        GridVenda.Row = VLStrLinha
        TipoPagto = GridVenda.Text
        
        GridVenda.Col = 8
        GridVenda.Row = VLStrLinha
        codvenda = Val(GridVenda.Text)
        
        StrSql = "SELECT * FROM tb_Venda where CodVenda=" & codvenda
        RecVenda.Open StrSql, vgCon, 1, 3
        
        '=== Pegar produto 01 ==========
        StrSql = "SELECT * FROM tb_Produto where CodProd=" & RecVenda.Fields.Item(5).Value
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            tipoprod01 = RecProd.Fields.Item(3).Value
        Else
            tipoprod01 = ""
        End If
        
        If tipoprod01 <> "" Then
            If tipoprod01 = "Armação" Then
                descrprod01 = RecProd.Fields.Item(4).Value & "/" & RecProd.Fields.Item(5).Value & "/" & RecProd.Fields.Item(6).Value & "/" & RecProd.Fields.Item(7).Value & "/" & RecProd.Fields.Item(8).Value
            Else
                descrprod01 = RecProd.Fields.Item(9).Value & "/" & RecProd.Fields.Item(10).Value
            End If
        Else
            descrprod01 = ""
        End If
        '===============================
        
        RecProd.Close
        
        '=== Pegar produto 02 ==========
        StrSql = "SELECT * FROM tb_Produto where CodProd=" & RecVenda.Fields.Item(6).Value
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            tipoprod02 = RecProd.Fields.Item(3).Value
        Else
            tipoprod02 = ""
        End If
        
        If tipoprod02 <> "" Then
            If tipoprod02 = "Armação" Then
                descrprod02 = RecProd.Fields.Item(4).Value & "/" & RecProd.Fields.Item(5).Value & "/" & RecProd.Fields.Item(6).Value & "/" & RecProd.Fields.Item(7).Value & "/" & RecProd.Fields.Item(8).Value
            Else
                descrprod02 = RecProd.Fields.Item(9).Value & "/" & RecProd.Fields.Item(10).Value
            End If
        Else
            descrprod02 = ""
        End If
        '===============================
        
        RecProd.Close
        
        '=== Pegar produto 03 ==========
        StrSql = "SELECT * FROM tb_Produto where CodProd=" & RecVenda.Fields.Item(7).Value
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            tipoprod03 = RecProd.Fields.Item(3).Value
        Else
            tipoprod03 = ""
        End If
        
        If tipoprod03 <> "" Then
            If tipoprod03 = "Armação" Then
                descrprod03 = RecProd.Fields.Item(4).Value & "/" & RecProd.Fields.Item(5).Value & "/" & RecProd.Fields.Item(6).Value & "/" & RecProd.Fields.Item(7).Value & "/" & RecProd.Fields.Item(8).Value
            Else
                descrprod03 = RecProd.Fields.Item(9).Value & "/" & RecProd.Fields.Item(10).Value
            End If
        Else
            descrprod03 = ""
        End If
        '===============================
        
        RecProd.Close
        
        '=== Pegar produto 04 ==========
        StrSql = "SELECT * FROM tb_Produto where CodProd=" & RecVenda.Fields.Item(8).Value
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            tipoprod04 = RecProd.Fields.Item(3).Value
        Else
            tipoprod04 = ""
        End If
        
        If tipoprod04 <> "" Then
            If tipoprod04 = "Armação" Then
                descrprod04 = RecProd.Fields.Item(4).Value & "/" & RecProd.Fields.Item(5).Value & "/" & RecProd.Fields.Item(6).Value & "/" & RecProd.Fields.Item(7).Value & "/" & RecProd.Fields.Item(8).Value
            Else
                descrprod04 = RecProd.Fields.Item(9).Value & "/" & RecProd.Fields.Item(10).Value
            End If
        Else
            descrprod04 = ""
        End If
        '===============================
        
        RecProd.Close
        
        '=== Pegar produto 05 ==========
        StrSql = "SELECT * FROM tb_Produto where CodProd=" & RecVenda.Fields.Item(9).Value
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            tipoprod05 = RecProd.Fields.Item(3).Value
        Else
            tipoprod05 = ""
        End If
        
        If tipoprod05 <> "" Then
            If tipoprod05 = "Armação" Then
                descrprod05 = RecProd.Fields.Item(4).Value & "/" & RecProd.Fields.Item(5).Value & "/" & RecProd.Fields.Item(6).Value & "/" & RecProd.Fields.Item(7).Value & "/" & RecProd.Fields.Item(8).Value
            Else
                descrprod05 = RecProd.Fields.Item(9).Value & "/" & RecProd.Fields.Item(10).Value
            End If
        Else
            descrprod05 = ""
        End If
        '===============================
        
        RecProd.Close
        
        '=== Pegar produto 06 ==========
        StrSql = "SELECT * FROM tb_Produto where CodProd=" & RecVenda.Fields.Item(10).Value
        RecProd.Open StrSql, vgCon, 1, 3
        
        If Not RecProd.EOF Then
            tipoprod06 = RecProd.Fields.Item(3).Value
        Else
            tipoprod06 = ""
        End If
        
        If tipoprod06 <> "" Then
            If tipoprod06 = "Armação" Then
                descrprod06 = RecProd.Fields.Item(4).Value & "/" & RecProd.Fields.Item(5).Value & "/" & RecProd.Fields.Item(6).Value & "/" & RecProd.Fields.Item(7).Value & "/" & RecProd.Fields.Item(8).Value
            Else
                descrprod06 = RecProd.Fields.Item(9).Value & "/" & RecProd.Fields.Item(10).Value
            End If
        Else
            descrprod06 = ""
        End If
        '===============================
        
        RecProd.Close
        RecVenda.Close
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18,campo19) " & _
        "VALUES ('" & cliente & "','" & vendedor & "','" & datavenda & "','" & valorvenda & "','" & desconto & "','" & tipovenda & "','" & TipoPagto & "','" & tipoprod01 & "','" & descrprod01 & "','" & tipoprod02 & "','" & descrprod02 & "','" & tipoprod03 & "','" & descrprod03 & "','" & tipoprod04 & "','" & descrprod04 & "','" & tipoprod05 & "','" & descrprod05 & "','" & tipoprod06 & "','" & descrprod06 & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptVenda.Show

End Sub

Private Sub CmdIncluirCli_Click()
    FrmCliente_Inc.Show
End Sub

Private Sub CmdIncluirCredsta_Click()
    FrmCrediarista_Inc.Show
End Sub

Private Sub CmdIncluirCx_Click()
    FrmCaixa_Inc.Show
End Sub

Private Sub CmdIncluirAlterarEst_Click()
    FrmEstoque_Inc_Alt.Show
End Sub

Private Sub CmdIncluirExt_Click()
    If OptExplic.Value = True Then
        frmExtra_folheto_inc.Show
        
    ElseIf OptCob.Value = True Then
        frmExtra_cartacob_inc.Show
        
    End If
End Sub

Private Sub CmdIncluirForn_Click()
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirMed_Click()
    FrmMedico_Inc.Show
End Sub

Private Sub CmdIncluirOrc_Click()
    FrmOrcamento_Inc.Show
End Sub

Private Sub CmdIncluirProd_Click()
    FrmProduto_Inc.Show
End Sub

Private Sub CmdIncluirRec_Click()
    If VGIntCodCli = 0 Then
        VGStrForm = "receita"
        FrmVenda_Inc_Cli.Show
    Else
        FrmReceita_Inc.Show
    End If
End Sub

Private Sub CmdVendaRec_Click()
    FrmVenda_Inc.Show
End Sub

Private Sub CmdIncluirVenda_Click()
    FrmVenda_Inc_Cli.Show
End Sub

Private Sub CmdPagar_Click()
    FrmCaixa_APagar_Cons.Show
End Sub

Private Sub CmdPesqCli_Click()

    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Cliente where 0=0"
            
    '====== PESQUISAR POR NOME ==========
    If TxtNomeCli.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNomeCli.Text & "%'"
        VLStrOrder = VLStrOrder + "Nome,"
    End If
            
    '====== PESQUISAR POR CPF ==========
    If TxtCpfCli.Text <> "" Then
        StrSql = StrSql + " and Cpf='" & TxtCpfCli.Text & "'"
        VLStrOrder = VLStrOrder + "Cpf,"
    End If
    
    '====== PESQUISAR POR SEXO ==========
    If CboSexoCli.Text <> "" Then
        StrSql = StrSql + " and Sexo='" & CboSexoCli.Text & "'"
        VLStrOrder = VLStrOrder + "Sexo,"
    End If
            
    '====== PESQUISAR POR BAIRRO ==========
    If TxtBairroCli.Text <> "" Then
        StrSql = StrSql + " and Bairro like '%" & TxtBairroCli.Text & "%'"
        VLStrOrder = VLStrOrder + "Bairro,"
    End If
            
    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelCli.Text <> "" Then
        StrSql = StrSql + " and Telefone like '%" & TxtTelCli.Text & "%'"
        VLStrOrder = VLStrOrder + "Telefone,"
    End If
            
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by Nome"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridCliente
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqCred_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
''    StrSql = "Select C.Nome,CS.Nome,CR.CodCred,CR.DtCred,CR.TipoCred,CR.ValorVenda," & _
''             "CR.Parcela,CR.Juros,CR.ValorTotal,CR.TipoEntr,CR.ValorEntr,CR.NumBanco," & _
''             "CR.NumCheque,CP.CodParc,CP.Vencimento,CP.Valor,CP.Quitado,CP.NumParc,CS.CodCredsta " & _
''             "from tb_Cliente as C,tb_Crediarista as CS,tb_Crediario as CR," & _
''             "tb_Crediario_Parcela as CP where CP.CodCred=CR.CodCred and C.CodCli=CR.CodCli and CS.CodCredsta=CR.CodCredsta"
            
    StrSql = "SELECT * FROM tb_Crediario as CR WHERE 0=0"
            
    '====== PESQUISAR POR CLIENTE ==========
    If TxtCliCred.Text <> "" Then
        'StrSql = StrSql + " and C.Nome like '%" & TxtCliCred.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",C.nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Cliente as C " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and C.CodCli=CR.CodCli and C.Nome like '%" & TxtCliCred.Text & "%'"
        VLStrOrder = VLStrOrder + "C.Nome,"
    End If
            
    '====== PESQUISAR POR CREDIARISTA ==========
    If TxtCredstaCred.Text <> "" Then
        'StrSql = StrSql + " and CS.Nome like '%" & TxtCredstaCred.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",CS.CodCredsta,CS.nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediarista as CS " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CS.CodCredsta = CR.CodCredsta and CS.Nome like '%" & TxtCredstaCred.Text & "%'"
        VLStrOrder = VLStrOrder + "CS.Nome,"
    End If
            
    '====== PESQUISAR POR TIPO CREDIÁRIO ==========
    If CboTipoCred.Text <> "" Then
        StrSql = StrSql + " and CR.TipoCred='" & CboTipoCred.Text & "'"
        VLStrOrder = VLStrOrder + "CR.TipoCred,"
    End If
    
    '====== PESQUISAR POR DATA DO CREDIÁRIO ==========
    
    If (TxtDtCred1.Text <> "" And TxtDtCred1.Text <> "__/__/____") And (TxtDtCred2.Text <> "" And TxtDtCred2.Text <> "__/__/____") Then
        StrSql = StrSql + " and CR.DtCred >=#" & FormataDataUS(TxtDtCred1.Text) & "# and CR.DtCred <= #" & FormataDataUS(TxtDtCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CR.DtCred desc,"
    
    ElseIf (TxtDtCred1.Text <> "" And TxtDtCred1.Text <> "__/__/____") And (TxtDtCred2.Text = "" Or TxtDtCred2.Text = "__/__/____") Then
        StrSql = StrSql + " and CR.DtCred =#" & FormataDataUS(TxtDtCred1.Text) & "#"
        VLStrOrder = VLStrOrder + "CR.DtCred desc,"
    
    ElseIf (TxtDtCred1.Text = "" Or TxtDtCred1.Text = "__/__/____") And (TxtDtCred2.Text <> "" And TxtDtCred2.Text <> "__/__/____") Then
        StrSql = StrSql + " and CR.DtCred =#" & FormataDataUS(TxtDtCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CR.DtCred desc,"
    End If
            
    '====== PESQUISAR POR DATA DO VENCIMENTO ==========
    
    If (TxtDtVencCred1.Text <> "" And TxtDtVencCred1.Text <> "__/__/____") And (TxtDtVencCred2.Text <> "" And TxtDtVencCred2.Text <> "__/__/____") Then
        StrSql = StrSql + " and CP.Vencimento >=#" & FormataDataUS(TxtDtVencCred1.Text) & "# and CP.Vencimento <= #" & FormataDataUS(TxtDtVencCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CP.Vencimento desc,"
    
    ElseIf (TxtDtVencCred1.Text <> "" And TxtDtVencCred1.Text <> "__/__/____") And (TxtDtVencCred2.Text = "" Or TxtDtVencCred2.Text = "__/__/____") Then
        StrSql = StrSql + " and CP.Vencimento =#" & FormataDataUS(TxtDtVencCred1.Text) & "#"
        VLStrOrder = VLStrOrder + "CP.Vencimento desc,"
    
    ElseIf (TxtDtVencCred1.Text = "" Or TxtDtVencCred1.Text = "__/__/____") And (TxtDtVencCred2.Text <> "" And TxtDtVencCred2.Text <> "__/__/____") Then
        StrSql = StrSql + " and CP.Vencimento =#" & FormataDataUS(TxtDtVencCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CP.Vencimento desc,"
    End If
            
    '====== PESQUISAR POR CÓDIGO DA PARCELA ==========
    If TxtCodParcCred.Text <> "" Then
        'StrSql = StrSql + " and CP.CodParc=" & TxtCodParcCred.Text & ""
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",CP.codparc,CP.vencimento,CP.valor,CP.quitado,CP.NumParc " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_CrediarioParcela as CP " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CP.CodCred = CR.CodCred and CP.CodParc=" & TxtCodParcCred.Text & ""
    End If
        
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder & ",CR.DtCred desc"
    Else
        StrSql = StrSql + " order by CR.DtCred desc"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridCrediario
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqCx_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Caixa where 0=0"
            
    '====== PESQUISAR POR DATA DO MOVIMENTO ==========
    If IsDate(TxtDtMovCx1.Text) = False Then
        TxtDtMovCx1.Text = ""
    End If
    If IsDate(TxtDtMovCx2.Text) = False Then
        TxtDtMovCx2.Text = ""
    End If
        
    If TxtDtMovCx1.Text <> "" And TxtDtMovCx2.Text <> "" Then
        StrSql = StrSql + " AND DtMov >= #" & FormataDataUS(TxtDtMovCx1.Text) & "# AND DtMov <= #" & FormataDataUS(TxtDtMovCx2.Text) & "#"
        VLStrOrder = VLStrOrder + "DtMov desc,"
        
    ElseIf TxtDtMovCx1.Text <> "" And TxtDtMovCx2.Text = "" Then
        StrSql = StrSql + " AND DtMov = #" & FormataDataUS(TxtDtMovCx1.Text) & "#"
        VLStrOrder = VLStrOrder + "DtMov desc,"
    
    ElseIf TxtDtMovCx1.Text = "" And TxtDtMovCx2.Text <> "" Then
        StrSql = StrSql + " AND DtMov = #" & FormataDataUS(TxtDtMovCx2.Text) & "#"
        VLStrOrder = VLStrOrder + "DtMov desc,"
        
    End If
    
    If TxtDtMovCx1.Text = "" Then
        TxtDtMovCx1.Text = "__/__/____"
    End If
    
    If TxtDtMovCx2.Text = "" Then
        TxtDtMovCx2.Text = "__/__/____"
    End If
            
    '====== PESQUISAR POR TIPO DE PAGAMENTO ==========
    If CboTipoPagtoCx.Text <> "" Then
        StrSql = StrSql + " and TipoPagto='" & CboTipoPagtoCx.Text & "'"
        VLStrOrder = VLStrOrder + "TipoPagto,"
    End If
    
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by DtMov desc"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridCaixa
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqEst_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Estoque as E,tb_Produto as P where E.CodProd=P.CodProd"
            
    '====== PESQUISAR POR TIPO DE PRODUTO ==========
    If CboProdEst.Text <> "" Then
        StrSql = StrSql + " and P.TipoProd='" & CboProdEst.Text & "'"
        VLStrOrder = VLStrOrder + "P.TipoProd,"
    End If
            
    '====== PESQUISAR POR QTDE MÍNIMA ==========
    If TxtQtdeMinEst.Text <> "" Then
        StrSql = StrSql + " and E.QtdeMin=" & TxtQtdeMinEst.Text & ""
        VLStrOrder = VLStrOrder + "E.QtdeMin,"
    End If
    
    '====== PESQUISAR POR COR ==========
    If TxtCorEst.Text <> "" Then
        StrSql = StrSql + " and P.Cor='" & TxtCorEst.Text & "'"
        VLStrOrder = VLStrOrder + "P.Cor,"
    End If
    
    '====== PESQUISAR POR NÚMERO ==========
    If TxtNumEst.Text <> "" Then
        StrSql = StrSql + " and P.Numero='" & TxtNumEst.Text & "'"
        VLStrOrder = VLStrOrder + "P.Numero,"
    End If
    
    '====== PESQUISAR POR MODELO ==========
    If TxtModEst.Text <> "" Then
        StrSql = StrSql + " and P.Modelo like '%" & TxtModEst.Text & "%'"
        VLStrOrder = VLStrOrder + "P.Modelo,"
    End If
    
    '====== PESQUISAR POR ARO ==========
    If TxtAroEst.Text <> "" Then
        StrSql = StrSql + " and P.TamAro='" & TxtAroEst.Text & "'"
        VLStrOrder = VLStrOrder + "P.TamAro,"
    End If
    
    '====== PESQUISAR POR PONTE ==========
    If TxtPteEst.Text <> "" Then
        StrSql = StrSql + " and P.TamPonte='" & TxtPteEst.Text & "'"
        VLStrOrder = VLStrOrder + "P.TamPonte,"
    End If
    
    '====== PESQUISAR POR TIPO ==========
    If TxtTipoEst.Text <> "" Then
        StrSql = StrSql + " and P.Tipo like '%" & TxtTipoEst.Text & "%'"
        VLStrOrder = VLStrOrder + "P.Tipo,"
    End If
    
    '====== PESQUISAR POR CHAVE ==========
    If TxtChaEst.Text <> "" Then
        StrSql = StrSql + " and P.Chave like '%" & TxtChaEst.Text & "%'"
        VLStrOrder = VLStrOrder + "P.Chave,"
    End If
    
            
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by P.TipoProd"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridEstoque
    
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqForn_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Fornecedor where 0=0"
            
    '====== PESQUISAR POR FORNECEDOR ==========
    If TxtNomeForn.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNomeForn.Text & "%'"
        VLStrOrder = VLStrOrder + "Nome,"
    End If

    '====== PESQUISAR POR CNPJ ==========
    If TxtCnpjForn.Text <> "" Then
        StrSql = StrSql + " and CNPJ='" & TxtCnpjForn.Text & "'"
        VLStrOrder = VLStrOrder + "CNPJ,"
    End If
    
    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelForn.Text <> "" Then
        StrSql = StrSql + " and Telefone like '%" & TxtTelForn.Text & "%'"
        VLStrOrder = VLStrOrder + "Telefone,"
    End If
    
    '====== PESQUISAR POR TIPO ==========
    If TxtTipoForn.Text <> "" Then
        StrSql = StrSql + " and Tipo like '%" & TxtTipoForn.Text & "%'"
        VLStrOrder = VLStrOrder + "Tipo,"
    End If
    
            
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by Nome"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridFornecedor
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqMed_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Medico where 0=0"
            
    '====== PESQUISAR POR NOME ==========
    If TxtNomeMed.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNomeMed.Text & "%'"
        VLStrOrder = VLStrOrder + "Nome,"
    End If

    '====== PESQUISAR POR CRM ==========
    If TxtCrmMed.Text <> "" Then
        StrSql = StrSql + " and Crm='" & TxtCrmMed.Text & "'"
        VLStrOrder = VLStrOrder + "Crm,"
    End If

    '====== PESQUISAR POR CPF ==========
    If TxtCpfMed.Text <> "" Then
        StrSql = StrSql + " and Cpf='" & TxtCpfMed.Text & "'"
        VLStrOrder = VLStrOrder + "Cpf,"
    End If
    
    '====== PESQUISAR POR CLÍNICA/CONSULTÓRIO ==========
    If TxtCliConsMed.Text <> "" Then
        StrSql = StrSql + " and CliCons like '%" & TxtCliConsMed.Text & "%'"
        VLStrOrder = VLStrOrder + "CliCons,"
    End If
            
    '====== PESQUISAR POR BAIRRO ==========
    If TxtBairroMed.Text <> "" Then
        StrSql = StrSql + " and Bairro like '%" & TxtBairroMed.Text & "%'"
        VLStrOrder = VLStrOrder + "Bairro,"
    End If
            
    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelMed.Text <> "" Then
        StrSql = StrSql + " and Telefone like '%" & TxtTelMed.Text & "%'"
        VLStrOrder = VLStrOrder + "Telefone,"
    End If
            
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by Nome"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridMedico
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqOrc_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Orcamento as O,tb_Vendedor as V where V.CodVendedor=O.CodVendedor"
            
    '====== PESQUISAR POR CLIENTE ==========
    If TxtCliOrc.Text <> "" Then
        StrSql = StrSql + " and O.Nome like '%" & TxtCliOrc.Text & "%'"
        VLStrOrder = VLStrOrder + "O.Nome,"
    End If
            
    '====== PESQUISAR POR VENDEDOR ==========
    If TxtVendOrc.Text <> "" Then
        StrSql = StrSql + " and V.Nome like '%" & TxtVendOrc.Text & "%'"
        VLStrOrder = VLStrOrder + "V.Nome,"
    End If
    
    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelOrc.Text <> "" Then
        StrSql = StrSql + " and O.Telefone like '%" & TxtTelOrc.Text & "%'"
        VLStrOrder = VLStrOrder + "O.Telefone,"
    End If

    '====== PESQUISAR POR DATA DO ORÇAMENTO ==========
    If (TxtDtOrc1.Text <> "" And TxtDtOrc1.Text <> "__/__/____") And (TxtDtOrc2.Text <> "" And TxtDtOrc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and O.DtOrc >=#" & FormataDataUS(TxtDtOrc1.Text) & "# and O.DtOrc <= #" & FormataDataUS(TxtDtOrc2.Text) & "#"
        VLStrOrder = VLStrOrder + "O.DtOrc desc,"
    
    ElseIf (TxtDtOrc1.Text <> "" And TxtDtOrc1.Text <> "__/__/____") And (TxtDtOrc2.Text = "" Or TxtDtOrc2.Text = "__/__/____") Then
        StrSql = StrSql + " and O.DtOrc =#" & FormataDataUS(TxtDtOrc1.Text) & "#"
        VLStrOrder = VLStrOrder + "O.DtOrc desc,"
    
    ElseIf (TxtDtOrc1.Text = "" Or TxtDtOrc1.Text = "__/__/____") And (TxtDtOrc2.Text <> "" And TxtDtOrc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and O.DtOrc =#" & FormataDataUS(TxtDtOrc2.Text) & "#"
        VLStrOrder = VLStrOrder + "O.DtOrc desc,"
    End If

            
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by O.Nome"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridOrcamento
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqProd_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    Dim VLStrSql1 As String
    Dim VLStrSql2 As String
    Dim VLStrSql3 As String
    
    Conecta
    
    StrSql = "SELECT P.CodProd,P.TipoProd,P.Cor,P.Numero,P.Modelo,P.TamAro,P.TamPonte," & _
             "P.Tipo,P.Chave,F.Nome,P.CodGriffe FROM tb_Produto as P,tb_Fornecedor as F " & _
             "WHERE P.CodForn=F.CodForn"
            
    '====== PESQUISAR POR FORNECEDOR ==========
    If CboFornProd.Text <> "" Then
        StrSql = StrSql + " and F.Nome like '%" & CboFornProd.Text & "%'"
        VLStrOrder = VLStrOrder + "F.Nome,"
    End If
            
    '====== PESQUISAR POR TIPO DE PRODUTO ==========
    If CboTipoProd.Text <> "" Then
        StrSql = StrSql + " and P.TipoProd like '%" & CboTipoProd.Text & "%'"
        VLStrOrder = VLStrOrder + "P.TipoProd,"
    End If
            
    '====== PESQUISAR POR TIPO DE LENTE ==========
    If CboLenteProd.Text <> "" Then
        StrSql = StrSql + " and P.Tipo like '%" & CboLenteProd.Text & "%'"
        VLStrOrder = VLStrOrder + "P.Tipo,"
    End If
            
    '====== PESQUISAR POR GRIFFE ==========
    If CboGriffeProd.Text <> "" Then
        'StrSql = StrSql + "  and G.Nome like '%" & CboGriffeProd.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",G.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Griffe as G " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and G.CodGriffe=P.CodGriffe and G.Nome like '%" & CboGriffeProd.Text & "%'"
        'StrSql = 1 & ",G.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - 1)) & ",tb_Griffe as G " & Trim(Mid(StrSql, 1, InStr(StrSql, "WHERE"))) & " and G.Nome like '%" & CboGriffeProd.Text & "%'"
        'VLStrOrder = VLStrOrder + "G.Nome,"
    End If
            
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by P.TipoProd"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridProduto
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqRec_Click()
    
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "SELECT CodRec,R.CodCli,R.CodMed,DtRec,LODEsf,LODCil,LODEixo,LOEEsf," & _
    "LOECil,LOEEixo,PODEsf,PODCil,PODEixo,POEEsf,POECil,POEEixo,ODDNP,OEDNP,ODAlt," & _
    "OEAlt,ODAdicao,OEAdicao,AOAdicao,R.Obs,C.Nome,M.Nome " & _
    "FROM tb_Receita AS R,tb_Cliente As C,tb_Medico As M " & _
    "WHERE R.CodCli=C.CodCli AND R.CodMed=M.CodMed "
            
    VLStrOrder = "C.Nome,"
            
    '====== PESQUISAR POR CLIENTE ==========
    If TxtRecCliente.Text <> "" Then
        StrSql = StrSql + " AND C.Nome like '%" & TxtRecCliente.Text & "%'"
    End If
            
    '====== PESQUISAR POR MÉDICO ==========
    If TxtRecMedico.Text <> "" Then
        StrSql = StrSql + " AND M.Nome like '%" & TxtRecMedico.Text & "%'"
        VLStrOrder = VLStrOrder + "M.Nome,"
    End If
    
    '====== PESQUISAR POR DATA DA RECEITA ==========
    If IsDate(TxtDtRec1.Text) = False Then
        TxtDtRec1.Text = ""
    End If
    If IsDate(TxtDtRec2.Text) = False Then
        TxtDtRec2.Text = ""
    End If
        
    If TxtDtRec1.Text <> "" And TxtDtRec2.Text <> "" Then
        StrSql = StrSql + " AND R.DtRec >= #" & FormataDataUS(TxtDtRec1.Text) & "# AND R.DtRec <= #" & FormataDataUS(TxtDtRec2.Text) & "#"
        VLStrOrder = VLStrOrder + "R.DtRec desc,"
        
    ElseIf TxtDtRec1.Text <> "" And TxtDtRec2.Text = "" Then
        StrSql = StrSql + " AND R.DtRec = #" & FormataDataUS(TxtDtRec1.Text) & "#"
        VLStrOrder = VLStrOrder + "R.DtRec desc,"
    
    ElseIf TxtDtRec1.Text = "" And TxtDtRec2.Text <> "" Then
        StrSql = StrSql + " AND R.DtRec = #" & FormataDataUS(TxtDtRec2.Text) & "#"
        VLStrOrder = VLStrOrder + "R.DtRec desc,"
        
    End If
    
    If TxtDtRec1.Text = "" Then
        TxtDtRec1.Text = "__/__/____"
    End If
    
    If TxtDtRec2.Text = "" Then
        TxtDtRec2.Text = "__/__/____"
    End If
    
    
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by R.DtRec desc"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridReceita
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdPesqVenda_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "SELECT V.CodVenda,V.DtVenda,V.TipoVenda,V.Desconto,V.TotalVenda," & _
             "V.TipoPagto,V.CodCli,V.CodVendedor FROM tb_Venda as V WHERE 0=0"
            
    '====== PESQUISAR POR CLIENTE ==========
    If TxtCliVend.Text <> "" Then
        'StrSql = StrSql + " and C.Nome like '%" & TxtNomeCli.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",C.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Cliente as C " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and C.CodCli=V.CodCli and C.Nome like '%" & TxtNomeCli.Text & "%'"
        VLStrOrder = VLStrOrder + "C.Nome,"
    End If
            
    '====== PESQUISAR POR DATA DA VENDA ==========
    If (TxtDtVenda1.Text <> "" And TxtDtVenda1.Text <> "__/__/____") And (TxtDtVenda2.Text <> "" And TxtDtVenda2.Text <> "__/__/____") Then
        StrSql = StrSql + " and V.DtVenda >=#" & FormataDataUS(TxtDtVenda1.Text) & "# and V.DtVenda <= #" & FormataDataUS(TxtDtVenda2.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtVenda desc,"
    
    ElseIf (TxtDtVenda1.Text <> "" And TxtDtVenda1.Text <> "__/__/____") And (TxtDtVenda2.Text = "" Or TxtDtVenda2.Text = "__/__/____") Then
        StrSql = StrSql + " and V.DtVenda =#" & FormataDataUS(TxtDtVenda1.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtVenda desc,"
    
    ElseIf (TxtDtVenda1.Text = "" Or TxtDtVenda1.Text = "__/__/____") And (TxtDtVenda2.Text <> "" And TxtDtVenda2.Text <> "__/__/____") Then
        StrSql = StrSql + " and V.DtVenda =#" & FormataDataUS(TxtDtVenda2.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtVenda desc,"
    End If
            
    '====== PESQUISAR POR TIPO VENDA ==========
    If CboTipoVenda.Text <> "" Then
        StrSql = StrSql + " and V.TipoVenda='" & CboTipoVenda.Text & "'"
        VLStrOrder = VLStrOrder + "V.TipoVenda,"
    End If
    
    '====== PESQUISAR POR VENDEDOR ==========
    If TxtVendedor.Text <> "" Then
        'StrSql = StrSql + " and VR.Nome like '%" & TxtVendedor.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",VR.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Vendedor as VR " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and VR.CodVendedor=V.CodVendedor and VR.Nome like '%" & TxtVendedor.Text & "%'"
        VLStrOrder = VLStrOrder + "VR.Nome,"
    End If
            
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by V.DtVenda desc,V.CodVenda desc"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridVenda
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdQuitarCred_Click()
    If VPStrCrediarioQuitado = "sim" Then
        FrmCrediario_Quitado.Show
    Else
        FrmCrediario_Quitar.Show
    End If
End Sub

Private Sub CmdReceber_Click()
    FrmCaixa_AReceber_Cons.Show
End Sub

Private Sub CmdVendedorVenda_Click()
    FrmVendedor_Cons.Show
End Sub

Private Sub Form_Activate()
  
    '==== Verifica se tem alerta de estoque =====
    Conecta

    Dim RecAlerta As New ADODB.Recordset
    Dim RecVerif As New ADODB.Recordset
    
    StrSql = "Select Ativado From tb_Alerta"
    RecAlerta.Open StrSql, vgCon, 1, 3

    If RecAlerta!ativado = "sim" Then
        'Alerta está ativado
        ChkDesatAlerta.Value = 0
            
        'verifica se tem produto com qtde mínima
        StrSql = "SELECT QtdeMin FROM tb_Estoque WHERE QtdeProd <= QtdeMin"
        RecVerif.Open StrSql, vgCon, 1, 3
        
        If Not RecVerif.EOF Then
            Desconecta
            
            VPStrResponse = MsgBox("Existem produtos no estoque com quantidade mínima." & Chr(13) & "Deseja visualizar a listagem agora?", vbYesNo, "Pró Vendas 2004 - Alerta de Estoque")
            If VPStrResponse = vbYes Then
                FrmEstoque_Alerta.Show
            End If
        Else
            Desconecta
        End If
    Else
        Desconecta
        ChkDesatAlerta.Value = 1
    End If

    '============================================
End Sub

Private Sub Form_Load()
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Top = 270
    Left = 375
    Height = 10440
    Width = 14565
    
    FraVenda.BorderStyle = 0
    FraVenda.Visible = True
    
    FraCliente.Visible = False
    FraReceita.Visible = False
    FraMedico.Visible = False
    FraFornecedor.Visible = False
    FraEstoque.Visible = False
    FraProduto.Visible = False
    FraCrediario.Visible = False
    FraCaixa.Visible = False
    FraExtra.Visible = False
    FraOrcamento.Visible = False
    
    '===== Iniciando opção EXTRA =====
    FraExplic.Visible = False
    FraMala.Visible = False
    FraNiver.Visible = False
    FraEtiqArm.Visible = False
    FraCob.Visible = False

    CmdImprimirExt.Enabled = False
    '=================================
    
    '==== Montar CboSexoCli ==========
    CboSexoCli.AddItem ("")
    CboSexoCli.AddItem ("Feminino")
    CboSexoCli.AddItem ("Masculino")
    '=================================

    '==== Montar CboProdEst(Estoque)==
    Call MontaCboProdEst
    '=================================
    
    '==== Montar Cbos de Produtos ====
    Call MontaCbosProd
    '=================================
    
    '==== Montar Cbos de Crediarios ====
    Call MontaCboTipoCred
    '=================================
    
    '==== Monta campo de consulta do crediário =====
    TxtDtCred1.Text = "__/__/____"
    TxtDtCred2.Text = "__/__/____"
    TxtDtVencCred1.Text = "__/__/____"
    TxtDtVencCred2.Text = "__/__/____"
    '===============================================
    
    '==== Monta campos e cbos de Caixa =====
    TxtDtMovCx1.Text = FormataData(Date)
    TxtDtMovCx2.Text = FormataData(Date)
    
    Call MontaCboTipoPagtoCX
    '===============================================
    
    '==== Monta campo de consulta do orçamento =====
    TxtDtOrc1.Text = "__/__/____"
    TxtDtOrc2.Text = "__/__/____"
    '===============================================
    
    '==== Monta campo de consulta da venda =====
    TxtDtVenda1.Text = "__/__/____"
    TxtDtVenda2.Text = "__/__/____"
    Call MontaCboTipoVenda
    '===============================================
    
End Sub

Private Sub Form_Resize()
  TabPrincipal.Left = (MDIPrincipal.Width / 2) - (TabPrincipal.Width / 2)
  TabPrincipal.Top = (MDIPrincipal.Height / 3) - (TabPrincipal.Height / 3)
  
  FraCliente.Left = (MDIPrincipal.Width / 2) - (FraCliente.Width / 2)
  FraCliente.Top = (MDIPrincipal.Height / 3) - (FraCliente.Height / 3.5)
  
  FraReceita.Left = (MDIPrincipal.Width / 2) - (FraReceita.Width / 2)
  FraReceita.Top = (MDIPrincipal.Height / 3) - (FraReceita.Height / 3.5)
  
  FraMedico.Left = (MDIPrincipal.Width / 2) - (FraMedico.Width / 2)
  FraMedico.Top = (MDIPrincipal.Height / 3) - (FraMedico.Height / 3.5)
  
  FraFornecedor.Left = (MDIPrincipal.Width / 2) - (FraFornecedor.Width / 2)
  FraFornecedor.Top = (MDIPrincipal.Height / 3) - (FraFornecedor.Height / 3.5)
  
  FraEstoque.Left = (MDIPrincipal.Width / 2) - (FraEstoque.Width / 2)
  FraEstoque.Top = (MDIPrincipal.Height / 3) - (FraEstoque.Height / 3.5)
  
  FraProduto.Left = (MDIPrincipal.Width / 2) - (FraProduto.Width / 2)
  FraProduto.Top = (MDIPrincipal.Height / 3) - (FraProduto.Height / 3.5)
  
  FraCrediario.Left = (MDIPrincipal.Width / 2) - (FraCrediario.Width / 2)
  FraCrediario.Top = (MDIPrincipal.Height / 3) - (FraCrediario.Height / 3.5)
  
  FraCaixa.Left = (MDIPrincipal.Width / 2) - (FraCaixa.Width / 2)
  FraCaixa.Top = (MDIPrincipal.Height / 3) - (FraCaixa.Height / 3.5)
  
  FraExtra.Left = (MDIPrincipal.Width / 2) - (FraExtra.Width / 2)
  FraExtra.Top = (MDIPrincipal.Height / 3) - (FraExtra.Height / 3.5)
  
  FraOrcamento.Left = (MDIPrincipal.Width / 2) - (FraOrcamento.Width / 2)
  FraOrcamento.Top = (MDIPrincipal.Height / 3) - (FraOrcamento.Height / 3.5)
  
  FraVenda.Left = (MDIPrincipal.Width / 2) - (FraVenda.Width / 2)
  FraVenda.Top = (MDIPrincipal.Height / 3) - (FraVenda.Height / 3.5)
End Sub

Private Sub GridCaixa_Click(ByVal Col As Long, ByVal Row As Long)
    Dim VLStrLinha As Integer
    
    GridCaixa.Row = Row
    GridCaixa.Col = 7
    If GridCaixa.Text <> "" And GridCaixa.Text <> "CodCx" Then
        VGIntCodCx = GridCaixa.Text
        CmdAlterarCx.Enabled = True
        CmdExcluirCx.Enabled = True
    Else
        CmdAlterarCx.Enabled = False
        CmdExcluirCx.Enabled = False
    End If
    
End Sub

Private Sub GridCliente_Click(ByVal Col As Long, ByVal Row As Long)
    GridCliente.Row = Row
    GridCliente.Col = 16
    If GridCliente.Text <> "" And GridCliente.Text <> "CodCli" Then
        VGIntCodCli = GridCliente.Text
        CmdAlterarCli.Enabled = True
        CmdExcluirCli.Enabled = True
    Else
        CmdAlterarCli.Enabled = False
        CmdExcluirCli.Enabled = False
    End If
End Sub

Private Sub GridCrediario_Click(ByVal Col As Long, ByVal Row As Long)
    GridCrediario.Row = Row
    GridCrediario.Col = 13
    If GridCrediario.Text <> "Quitado" Then
        VPStrCrediarioQuitado = GridCrediario.Text
    End If
    
    GridCrediario.Row = Row
    GridCrediario.Col = 14
    If GridCrediario.Text <> "CodCred" And GridCrediario.Text <> "" Then
        VGIntCodCred = GridCrediario.Text
    
        GridCrediario.Row = Row
        GridCrediario.Col = 15
        If GridCrediario.Text <> "CodParc" Then
            VGIntCodParc = GridCrediario.Text
        End If
        
        GridCrediario.Row = Row
        GridCrediario.Col = 16
        If GridCrediario.Text <> "CodCredsta" Then
            VGIntCodCredsta = GridCrediario.Text
        End If
        
        CmdAlterarCred.Enabled = True
        CmdExcluirCred.Enabled = True
        CmdQuitarCred.Enabled = True
        CmdAlterarCredsta.Enabled = True
        CmdExcluirCredsta.Enabled = True
        CmdImprimirCred.Enabled = True
        CmdImprimirCredsta.Enabled = True
    Else
        CmdAlterarCred.Enabled = False
        CmdExcluirCred.Enabled = False
        CmdQuitarCred.Enabled = False
        CmdAlterarCredsta.Enabled = False
        CmdExcluirCredsta.Enabled = False
        CmdImprimirCred.Enabled = False
        CmdImprimirCredsta.Enabled = False
    End If
End Sub

Private Sub GridEstoque_Click(ByVal Col As Long, ByVal Row As Long)
    GridEstoque.Row = Row
    GridEstoque.Col = 8
    VGIntCodEst = GridEstoque.Text
    
    CmdIncluirAlterarEst.Enabled = True
    CmdExcluirEst.Enabled = True
End Sub

Private Sub GridFornecedor_Click(ByVal Col As Long, ByVal Row As Long)
    GridFornecedor.Row = Row
    GridFornecedor.Col = 14
    If GridFornecedor.Text <> "" And GridFornecedor.Text <> "CodForn" Then
        VGIntCodForn = GridFornecedor.Text
        
        CmdAlterarForn.Enabled = True
        CmdExcluirForn.Enabled = True
    Else
        CmdAlterarForn.Enabled = False
        CmdExcluirForn.Enabled = False
    End If
End Sub

Private Sub GridMedico_Click(ByVal Col As Long, ByVal Row As Long)
    GridMedico.Row = Row
    GridMedico.Col = 16
    If GridMedico.Text <> "" And GridMedico.Text <> "CodMed" Then
        VGIntCodMed = GridMedico.Text
        
        CmdAlterarMed.Enabled = True
        CmdExcluirMed.Enabled = True
    Else
        CmdAlterarMed.Enabled = False
        CmdExcluirMed.Enabled = False
    End If
End Sub

Private Sub GridOrcamento_Click(ByVal Col As Long, ByVal Row As Long)
    GridOrcamento.Row = Row
    GridOrcamento.Col = 20
    
    If GridOrcamento.Text <> "" And GridOrcamento.Text <> "CodOrc" Then
        VGIntCodOrc = GridOrcamento.Text
        CmdAlterarOrc.Enabled = True
        CmdExcluirOrc.Enabled = True
    Else
        CmdAlterarOrc.Enabled = False
        CmdExcluirOrc.Enabled = False
    End If
End Sub

Private Sub GridProduto_Click(ByVal Col As Long, ByVal Row As Long)
    GridProduto.Row = Row
    GridProduto.Col = 11
    VGIntCodProd = GridProduto.Text
    
    CmdAlterarProd.Enabled = True
    CmdExcluirProd.Enabled = True
End Sub

Private Sub GridReceita_Click(ByVal Col As Long, ByVal Row As Long)
    GridReceita.Row = Row
    GridReceita.Col = 24
    If GridReceita.Text <> "" And GridReceita.Text <> "CodRec" Then
    VGIntCodRec = GridReceita.Text
        GridReceita.Row = Row
        GridReceita.Col = 1
        VGStrNomeCli = GridReceita.Text
        
        CmdAlterarRec.Enabled = True
        CmdExcluirRec.Enabled = True
    Else
        CmdAlterarRec.Enabled = False
        CmdExcluirRec.Enabled = False
    End If
End Sub

Private Sub GridVenda_Click(ByVal Col As Long, ByVal Row As Long)
    GridVenda.Row = Row
    GridVenda.Col = 8
    If GridVenda.Text <> "" And GridVenda.Text <> "CodVenda" Then
        VGIntCodVenda = GridVenda.Text
        
        CmdExcluirVenda.Enabled = True
        CmdDetVenda.Enabled = True
        CmdCarne.Enabled = True
    Else
        CmdExcluirVenda.Enabled = False
        CmdDetVenda.Enabled = False
        CmdCarne.Enabled = False
    End If
    
    GridVenda.Row = Row
    GridVenda.Col = 7
    If GridVenda.Text = "Carnê" Then
       CmdCarne.Enabled = True
    Else
       CmdCarne.Enabled = False
    End If
End Sub

Private Sub OptCob_Click()
    FraExplic.Visible = False
    FraMala.Visible = False
    FraNiver.Visible = False
    FraEtiqArm.Visible = False
    FraCob.Visible = True
    FraPropCred.Visible = False

    CmdImprimirExt.Enabled = True
End Sub

Private Sub OptEtiqArm_Click()
    FraExplic.Visible = False
    FraMala.Visible = False
    FraNiver.Visible = False
    FraEtiqArm.Visible = True
    FraCob.Visible = False
    FraPropCred.Visible = False

    CmdImprimirExt.Enabled = True
    
    Call MontaGriffe
End Sub

Private Sub OptExplic_Click()
    FraExplic.Visible = True
    FraMala.Visible = False
    FraNiver.Visible = False
    FraEtiqArm.Visible = False
    FraCob.Visible = False
    FraPropCred.Visible = False
    
    CmdImprimirExt.Enabled = True
End Sub

Private Sub OptMala_Click()
    FraExplic.Visible = False
    FraMala.Visible = True
    FraNiver.Visible = False
    FraEtiqArm.Visible = False
    FraCob.Visible = False
    FraPropCred.Visible = False

    CmdImprimirExt.Enabled = True
End Sub

Private Sub OptNiver_Click()
    FraExplic.Visible = False
    FraMala.Visible = False
    FraNiver.Visible = True
    FraEtiqArm.Visible = False
    FraCob.Visible = False
    FraPropCred.Visible = False

    CmdImprimirExt.Enabled = True
End Sub

Private Sub OptPropCred_Click()
    FraExplic.Visible = False
    FraMala.Visible = False
    FraNiver.Visible = False
    FraEtiqArm.Visible = False
    FraCob.Visible = False
    FraPropCred.Visible = True

    CmdImprimirExt.Enabled = True
End Sub

Private Sub TabPrincipal_Click()
    
    If TabPrincipal.Tabs.Item(1).Selected = True Then
        '=== VENDA ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = True

        CmdExcluirVenda.Enabled = False
        CmdImprimirVenda.Enabled = False
        CmdDetVenda.Enabled = False
        CmdCarne.Enabled = False
        
        TxtCliVend.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(2).Selected = True Then
        '=== CLIENTE ===
        FraCliente.Visible = True
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False
        
        CmdAlterarCli.Enabled = False
        CmdExcluirCli.Enabled = False
        CmdImprimirCli.Enabled = False
        
        TxtNomeCli.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(3).Selected = True Then
        '=== RECEITA ===
        FraCliente.Visible = False
        FraReceita.Visible = True
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False
    
        CmdAlterarRec.Enabled = False
        CmdExcluirRec.Enabled = False
        CmdImprimirRec.Enabled = False
        
        TxtRecCliente.SetFocus
        
    ElseIf TabPrincipal.Tabs.Item(4).Selected = True Then
        '=== MÉDICO ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = True
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False
        
        CmdAlterarMed.Enabled = False
        CmdExcluirMed.Enabled = False
        CmdImprimirMed.Enabled = False
        
        TxtNomeMed.SetFocus
        
    ElseIf TabPrincipal.Tabs.Item(5).Selected = True Then
        '=== FORNECEDOR ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = True
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False
        
        CmdAlterarForn.Enabled = False
        CmdExcluirForn.Enabled = False
        CmdImprimirForn.Enabled = False
        
        TxtNomeForn.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(6).Selected = True Then
        '=== ESTOQUE ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = True
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False
    
        CmdIncluirAlterarEst.Enabled = False
        CmdExcluirEst.Enabled = False
        CmdImprimirEst.Enabled = False
        
        CboProdEst.SetFocus
        
    ElseIf TabPrincipal.Tabs.Item(7).Selected = True Then
        '=== PRODUTO ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = True
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False
    
        CmdAlterarProd.Enabled = False
        CmdExcluirProd.Enabled = False
        CmdImprimirProd.Enabled = False
        
        CboFornProd.SetFocus
        
    ElseIf TabPrincipal.Tabs.Item(8).Selected = True Then
        '=== CREDIÁRIO ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = True
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False

        CmdAlterarCred.Enabled = False
        CmdExcluirCred.Enabled = False
        CmdImprimirCred.Enabled = False
        CmdQuitarCred.Enabled = False
        CmdExcluirCredsta.Enabled = False
        CmdImprimirCredsta.Enabled = False
        
        TxtCliCred.SetFocus
        
    ElseIf TabPrincipal.Tabs.Item(9).Selected = True Then
        '=== CAIXA ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = True
        FraExtra.Visible = False
        FraOrcamento.Visible = False
        FraVenda.Visible = False
        
        CmdAlterarCx.Enabled = False
        CmdExcluirCx.Enabled = False
        CmdImprimirCx.Enabled = False
        
        TxtDtMovCx1.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(10).Selected = True Then
        '=== EXTRA ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = True
        FraOrcamento.Visible = False
        FraVenda.Visible = False
        
        'CboTipoCarta.Text = "Carta simples de cobrança"
        
    ElseIf TabPrincipal.Tabs.Item(11).Selected = True Then
        '=== ORÇAMENTO ===
        FraCliente.Visible = False
        FraReceita.Visible = False
        FraMedico.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraCrediario.Visible = False
        FraCaixa.Visible = False
        FraExtra.Visible = False
        FraOrcamento.Visible = True
        FraVenda.Visible = False
    
        CmdAlterarOrc.Enabled = False
        CmdExcluirOrc.Enabled = False
        CmdImprimirOrc.Enabled = False
        
        TxtCliOrc.SetFocus
        
    End If
End Sub

Sub MontaGridVenda()
    Dim VLIntLinha As Long
    Dim RecCli As New ADODB.Recordset
    Dim RecVend As New ADODB.Recordset
    
    If RecPesq.EOF Then
        LblNumTotalVend.Caption = "Nenhuma venda encontrada."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridVenda.Refresh
        GridVenda.MaxRows = 0
        
        CmdExcluirVenda.Enabled = False
        CmdImprimirVenda.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridVenda.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridVenda.Row = VLIntLinha
            GridVenda.Lock = True
            
            'Cliente
            StrSql = "Select Nome from tb_Cliente where CodCli=" & RecPesq!CodCli
            RecCli.Open StrSql, vgCon, 1, 3
            
            GridVenda.Col = 1
            If Not RecCli.EOF Then
                GridVenda.Text = VerificaNulo(RecCli!nome)
            Else
                GridVenda.Text = "- - - - - - - - - -"
            End If
            GridVenda.Lock = True
            
            RecCli.Close
            
            'Vendedor
            StrSql = "Select Nome from tb_Vendedor where CodVendedor=" & RecPesq!codvendedor
            RecVend.Open StrSql, vgCon, 1, 3
            
            GridVenda.Col = 2
            If Not RecVend.EOF Then
                GridVenda.Text = VerificaNulo(RecVend!nome)
            Else
                GridVenda.Text = "- - - - - - - - - -"
            End If
            GridVenda.Lock = True
            
            RecVend.Close
            
            'Data venda
            GridVenda.Col = 3
            GridVenda.Text = FormataData(RecPesq.Fields.Item(1).Value)
            GridVenda.Lock = True
            
            'Valor venda
            GridVenda.Col = 4
            GridVenda.Text = FormataMoeda(RecPesq.Fields.Item(4).Value)
            GridVenda.Lock = True
            
            'Desconto
            GridVenda.Col = 5
            If RecPesq.Fields.Item(3).Value <> "" And IsNull(RecPesq.Fields.Item(3).Value) = False Then
                GridVenda.Text = FormataNum(RecPesq.Fields.Item(3).Value) & "%"
            Else
                GridVenda.Text = ""
            End If
            GridVenda.Lock = True
            
            'Tipo Venda
            GridVenda.Col = 6
            GridVenda.Text = VerificaNulo(RecPesq.Fields.Item(2).Value)
            GridVenda.Lock = True
            
            'Tipo pagto
            GridVenda.Col = 7
            GridVenda.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridVenda.Lock = True
            
            'CodVenda
            GridVenda.Col = 8
            GridVenda.Text = Val(RecPesq.Fields.Item(0).Value)
            GridVenda.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridVenda.MaxRows = GridVenda.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE VENDAS PESQUISADOS =========
         GridVenda.MaxRows = GridVenda.MaxRows - 1
         
         If GridVenda.MaxRows = 1 Then
            LblNumTotalVend.Caption = FormataNum(GridVenda.MaxRows) & " venda encontrada."
         Else
            LblNumTotalVend.Caption = FormataNum(GridVenda.MaxRows) & " vendas encontradas."
         End If
         '================================================
         
         CmdImprimirVenda.Enabled = True
    End If

End Sub

Sub MontaGridCrediario()
    Dim VLIntCodCred As Long
    Dim VLIntLinha As Long
    Dim RecCli As New ADODB.Recordset
    Dim RecCredsta As New ADODB.Recordset
    Dim RecParc As New ADODB.Recordset
    
    If RecPesq.EOF Then
        LblNumTotalCred.Caption = "Nenhum crediário encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridCrediario.Refresh
        GridCrediario.MaxRows = 0
        
        CmdAlterarCredsta.Enabled = False
        CmdExcluirCredsta.Enabled = False
        CmdImprimirCredsta.Enabled = False
        CmdAlterarCred.Enabled = False
        CmdExcluirCred.Enabled = False
        CmdImprimirCred.Enabled = False
        CmdQuitarCred.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridCrediario.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridCrediario.Row = VLIntLinha
            GridCrediario.Lock = True
            
            'Cliente
            StrSql = "Select Nome from tb_Cliente where CodCli=" & RecPesq!CodCli
            RecCli.Open StrSql, vgCon, 1, 3
            
            GridCrediario.Col = 1
            If Not RecCli.EOF Then
                GridCrediario.Text = VerificaNulo(RecCli!nome)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Crediarista
            StrSql = "Select CodCredsta,Nome from tb_Crediarista where CodCredsta=" & RecPesq!CodCredsta
            RecCredsta.Open StrSql, vgCon, 1, 3
            
            GridCrediario.Col = 2
            If Not RecCredsta.EOF Then
                GridCrediario.Text = VerificaNulo(RecCredsta!nome)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Data Crediário
            GridCrediario.Col = 3
            GridCrediario.Text = FormataData(RecPesq!dtcred)
            GridCrediario.Lock = True
            
            'Tipo crediário
            GridCrediario.Col = 4
            GridCrediario.Text = VerificaNulo(RecPesq!tipocred)
            GridCrediario.Lock = True
            
            'Valor venda
            GridCrediario.Col = 5
            GridCrediario.Text = FormataMoeda(RecPesq!valorvenda)
            GridCrediario.Lock = True
            
            'Juros
            GridCrediario.Col = 6
            GridCrediario.Text = FormataNum(RecPesq!juros)
            GridCrediario.Lock = True
            
            'Valor total
            GridCrediario.Col = 7
            GridCrediario.Text = FormataMoeda(RecPesq!valortotal)
            GridCrediario.Lock = True
            
            'Tipo entrada
            GridCrediario.Col = 8
            If (RecPesq!numbanco <> "" And IsNull(RecPesq!numbanco) = False) And (RecPesq!numcheque <> "" And IsNull(RecPesq!numcheque) = False) Then
                GridCrediario.Text = VerificaNulo(RecPesqtipoentr) & " (" & RecPesq!numbanco & "/" & RecPesq!numcheque & ")"
            Else
                GridCrediario.Text = VerificaNulo(RecPesq!tipoentr)
            End If
            GridCrediario.Lock = True
            
            'Valor entrada
            GridCrediario.Col = 9
            If RecPesq!valorentr <> "" And IsNull(RecPesq!valorentr) = False Then
                GridCrediario.Text = FormataMoeda(RecPesq!valorentr)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Parcela
            StrSql = "Select CodParc,Vencimento,Valor,Quitado,NumParc " & _
                     "from tb_Crediario_Parcela where CodCred=" & RecPesq!CodCred
            RecParc.Open StrSql, vgCon, 1, 3
            
            GridCrediario.Col = 10
            If Not RecParc.EOF Then
                GridCrediario.Text = FormataNum(RecParc!NumParc) & "/" & FormataNum(RecPesq!parcela)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Vencimento
            GridCrediario.Col = 11
            If Not RecParc.EOF Then
                GridCrediario.Text = FormataData(RecParc!vencimento)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Valor
            GridCrediario.Col = 12
            If Not RecParc.EOF Then
                GridCrediario.Text = FormataMoeda(RecParc!valor)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Quitado
            GridCrediario.Col = 13
            If Not RecParc.EOF Then
                GridCrediario.Text = VerificaNulo(RecParc!quitado)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'CodCred
            GridCrediario.Col = 14
            GridCrediario.Text = Val(RecPesq!CodCred)
            GridCrediario.Lock = True
            
            'CodParc
            GridCrediario.Col = 15
            If Not RecParc.EOF Then
                GridCrediario.Text = Val(RecParc!codparc)
            Else
                GridCrediario.Text = Val(0)
            End If
            GridCrediario.Lock = True
            
            'CodCredsta
            GridCrediario.Col = 16
            GridCrediario.Text = Val(RecPesq!CodCredsta)
            GridCrediario.Lock = True
            
            RecCli.Close
            RecCredsta.Close
            RecParc.Close
            
            VLIntLinha = VLIntLinha + 1
            
            GridCrediario.MaxRows = GridCrediario.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE CREDIÁRIOS PESQUISADOS =========
         GridCrediario.MaxRows = GridCrediario.MaxRows - 1
         
         If GridCrediario.MaxRows = 1 Then
            LblNumTotalCred.Caption = FormataNum(GridCrediario.MaxRows) & " crediário encontrado."
         Else
            LblNumTotalCred.Caption = FormataNum(GridCrediario.MaxRows) & " crediários encontrados."
         End If
         '================================================
         
    End If

End Sub

Sub MontaGridEstoque()
    Dim VLIntCodEst As Long
    Dim VLIntLinha As Long
    Dim RecGrif As New ADODB.Recordset
    Dim Griffe As String
    
    If RecPesq.EOF Then
        LblNumTotalEst.Caption = "Nenhuma informação encontrada."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridEstoque.Refresh
        GridEstoque.MaxRows = 0
        
        CmdIncluirAlterarEst.Enabled = False
        CmdExcluirEst.Enabled = False
        CmdImprimirEst.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridEstoque.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridEstoque.Row = VLIntLinha
            GridEstoque.Lock = True
            
            'Tipo Produto
            GridEstoque.Col = 1
            GridEstoque.Text = VerificaNulo(RecPesq.Fields.Item(9).Value)
            GridEstoque.Lock = True
            
            'Produto
            If RecPesq.Fields.Item(8).Value <> 0 And RecPesq.Fields.Item(8).Value <> "" And IsNull(RecPesq.Fields.Item(8).Value) = False Then
                StrSql = "Select Nome From tb_Griffe where CodGriffe=" & RecPesq.Fields.Item(8).Value
                RecGrif.Open StrSql, vgCon, 1, 3
                
                If Not RecGrif.EOF Then
                    Griffe = RecGrif.Fields.Item(0).Value
                Else
                    Griffe = ""
                End If
                
                RecGrif.Close
                
            Else
                Griffe = ""
            End If
            
            GridEstoque.Col = 2
            If Griffe = "" Then
            'mostra dados para lentes
                GridEstoque.Text = VerificaNulo(RecPesq.Fields.Item(15).Value) & "/" & VerificaNulo(RecPesq.Fields.Item(16).Value)
            Else
            'mostra dados para armação
                GridEstoque.Text = Griffe & "/" & VerificaNulo(RecPesq.Fields.Item(10).Value) & "/" & VerificaNulo(RecPesq.Fields.Item(11).Value) & "/" & VerificaNulo(RecPesq.Fields.Item(12).Value) & "/" & VerificaNulo(RecPesq.Fields.Item(13).Value) & "/" & VerificaNulo(RecPesq.Fields.Item(14).Value)
            End If
            GridEstoque.Lock = True
            
            'Qtde Mínima
            GridEstoque.Col = 3
            GridEstoque.Text = VerificaNulo(RecPesq.Fields.Item(2).Value)
            GridEstoque.Lock = True
            
            'Qtde em estoque
            GridEstoque.Col = 4
            GridEstoque.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridEstoque.Lock = True
            
            'Preço Fabricante
            GridEstoque.Col = 5
            GridEstoque.Text = VerificaNulo(RecPesq.Fields.Item(17).Value)
            GridEstoque.Lock = True
            
            'Multiplicar
            GridEstoque.Col = 6
            GridEstoque.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridEstoque.Lock = True
            
            'Preço Venda
            GridEstoque.Col = 7
            GridEstoque.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridEstoque.Lock = True
            
            'CodEst
            GridEstoque.Col = 8
            GridEstoque.Text = Val(RecPesq.Fields.Item(0).Value)
            GridEstoque.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridEstoque.MaxRows = GridEstoque.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE INFORMAÇÕES DO ESTOQUE PESQUISADOS =========
         GridEstoque.MaxRows = GridEstoque.MaxRows - 1
         
         If GridEstoque.MaxRows = 1 Then
            LblNumTotalEst.Caption = FormataNum(GridEstoque.MaxRows) & " informação encontrada."
         Else
            LblNumTotalEst.Caption = FormataNum(GridEstoque.MaxRows) & " informações encontradas."
         End If
         '================================================
         
         CmdImprimirEst.Enabled = True
    End If

End Sub

Sub MontaGridCliente()
    Dim VLIntCodCli As Long
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalCli.Caption = "Nenhum cliente encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridCliente.Refresh
        GridCliente.MaxRows = 0
        
        CmdAlterarCli.Enabled = False
        CmdExcluirCli.Enabled = False
        CmdImprimirCli.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridCliente.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridCliente.Row = VLIntLinha
            GridCliente.Lock = True
            
            'Nome
            GridCliente.Col = 1
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(2).Value)
            GridCliente.Lock = True
            
            'Cliente desde
            GridCliente.Col = 2
            GridCliente.Text = FormataData(RecPesq.Fields.Item(1).Value)
            GridCliente.Lock = True
            
            'Sexo
            GridCliente.Col = 3
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridCliente.Lock = True
            
            'Endereço
            GridCliente.Col = 4
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridCliente.Lock = True
            
            'Bairro
            GridCliente.Col = 5
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridCliente.Lock = True
            
            'Cep
            GridCliente.Col = 6
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(6).Value)
            GridCliente.Lock = True
            
            'Cidade
            GridCliente.Col = 7
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(7).Value)
            GridCliente.Lock = True
            
            'Estado
            GridCliente.Col = 8
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(8).Value)
            GridCliente.Lock = True
            
            'Data Nascimento
            GridCliente.Col = 9
            GridCliente.Text = FormataData(VerificaNulo(RecPesq.Fields.Item(9).Value))
            GridCliente.Lock = True
            
            'Telefone
            GridCliente.Col = 10
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(10).Value)
            GridCliente.Lock = True
            
            'Celular
            GridCliente.Col = 11
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(11).Value)
            GridCliente.Lock = True
            
            'Fax
            GridCliente.Col = 12
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(12).Value)
            GridCliente.Lock = True
            
            'Cpf
            GridCliente.Col = 13
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(13).Value)
            GridCliente.Lock = True
            
            'Email
            GridCliente.Col = 14
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(14).Value)
            GridCliente.Lock = True
            
            'Observação
            GridCliente.Col = 15
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(15).Value)
            GridCliente.Lock = True
            
            'CodCli
            GridCliente.Col = 16
            GridCliente.Text = Val(RecPesq.Fields.Item(0).Value)
            GridCliente.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridCliente.MaxRows = GridCliente.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE CLIENTES PESQUISADOS =========
         GridCliente.MaxRows = GridCliente.MaxRows - 1
         
         If GridCliente.MaxRows = 1 Then
            LblNumTotalCli.Caption = FormataNum(GridCliente.MaxRows) & " cliente encontrado."
         Else
            LblNumTotalCli.Caption = FormataNum(GridCliente.MaxRows) & " clientes encontrados."
         End If
         '================================================
         
         CmdImprimirCli.Enabled = True
    End If

End Sub

Sub MontaGridReceita()
    
    Dim VLIntCodRec As Long
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalRec.Caption = "Nenhuma receita encontrada."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridReceita.Refresh
        GridReceita.MaxRows = 0
        
        CmdAlterarRec.Enabled = False
        CmdExcluirRec.Enabled = False
        CmdImprimirRec.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridReceita.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridReceita.Row = VLIntLinha
            GridReceita.Lock = True
            
            'Cliente
            GridReceita.Col = 1
            GridReceita.Text = RecPesq.Fields.Item(24).Value
            GridReceita.Lock = True
            
            'Médico
            GridReceita.Col = 2
            GridReceita.Text = RecPesq.Fields.Item(25).Value
            GridReceita.Lock = True
            
            'Data da Receita
            GridReceita.Col = 3
            GridReceita.Text = FormataData(RecPesq.Fields.Item(3).Value)
            GridReceita.Lock = True
            
            'Longe (OD) - Esf
            GridReceita.Col = 4
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridReceita.Lock = True
            
            'Longe (OD) - Cil
            GridReceita.Col = 5
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridReceita.Lock = True
            
            'Longe (OD) - Eixo
            GridReceita.Col = 6
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(6).Value)
            GridReceita.Lock = True
            
            'Longe (OE) - Esf
            GridReceita.Col = 7
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(7).Value)
            GridReceita.Lock = True
            
            'Longe (OE) - Cil
            GridReceita.Col = 8
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(8).Value)
            GridReceita.Lock = True
            
            'Longe (OE) - Eixo
            GridReceita.Col = 9
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(9).Value)
            GridReceita.Lock = True
            
            'Perto (OD) - Esf
            GridReceita.Col = 10
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(10).Value)
            GridReceita.Lock = True
            
            'Perto (OD) - Cil
            GridReceita.Col = 11
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(11).Value)
            GridReceita.Lock = True
            
            'Perto (OD) - Eixo
            GridReceita.Col = 12
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(12).Value)
            GridReceita.Lock = True
            
            'Perto (OE) - Esf
            GridReceita.Col = 13
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(13).Value)
            GridReceita.Lock = True
            
            'Perto (OE) - Cil
            GridReceita.Col = 14
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(14).Value)
            GridReceita.Lock = True
            
            'Perto (OE) - Eixo
            GridReceita.Col = 15
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(15).Value)
            GridReceita.Lock = True
            
            'DNP - OD.
            GridReceita.Col = 16
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(16).Value)
            GridReceita.Lock = True
            
            'DNP - OE.
            GridReceita.Col = 17
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(17).Value)
            GridReceita.Lock = True
            
            'Altura - OD.
            GridReceita.Col = 18
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(18).Value)
            GridReceita.Lock = True
            
            'Altura - OE.
            GridReceita.Col = 19
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(19).Value)
            GridReceita.Lock = True
            
            'Adição - OD.
            GridReceita.Col = 20
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(20).Value)
            GridReceita.Lock = True
            
            'Adição - OE.
            GridReceita.Col = 21
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(21).Value)
            GridReceita.Lock = True
            
            'Adição - AO.
            GridReceita.Col = 22
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(22).Value)
            GridReceita.Lock = True
            
            'Observação
            GridReceita.Col = 23
            GridReceita.Text = VerificaNulo(RecPesq.Fields.Item(23).Value)
            GridReceita.Lock = True
            
            'CodRec
            GridReceita.Col = 24
            GridReceita.Text = Val(RecPesq.Fields.Item(0).Value)
            GridReceita.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridReceita.MaxRows = GridReceita.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE CLIENTES PESQUISADOS =========
         GridReceita.MaxRows = GridReceita.MaxRows - 1
         
         If GridReceita.MaxRows = 1 Then
            LblNumTotalRec.Caption = FormataNum(GridReceita.MaxRows) & " receita encontrada."
         Else
            LblNumTotalRec.Caption = FormataNum(GridReceita.MaxRows) & " receitas encontradas."
         End If
         '================================================
         
         CmdImprimirRec.Enabled = True
    End If

End Sub

Sub MontaGridMedico()

    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalMed.Caption = "Nenhum médico encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridMedico.Refresh
        GridMedico.MaxRows = 0
        
        CmdAlterarMed.Enabled = False
        CmdExcluirMed.Enabled = False
        CmdImprimirMed.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridMedico.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridMedico.Row = VLIntLinha
            GridMedico.Lock = True
            
            'Nome
            GridMedico.Col = 1
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(1).Value)
            GridMedico.Lock = True
            
            'Clínica/Consultório
            GridMedico.Col = 2
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(2).Value)
            GridMedico.Lock = True
            
            'CRM
            GridMedico.Col = 3
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridMedico.Lock = True
            
            'Endereço
            GridMedico.Col = 4
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridMedico.Lock = True
            
            'Bairro
            GridMedico.Col = 5
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridMedico.Lock = True
            
            'Cep
            GridMedico.Col = 6
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(6).Value)
            GridMedico.Lock = True
            
            'Cidade
            GridMedico.Col = 7
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(7).Value)
            GridMedico.Lock = True
            
            'Estado
            GridMedico.Col = 8
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(8).Value)
            GridMedico.Lock = True
            
            'Data Nascimento
            GridMedico.Col = 9
            GridMedico.Text = FormataData(VerificaNulo(RecPesq.Fields.Item(9).Value))
            GridMedico.Lock = True
            
            'Telefone
            GridMedico.Col = 10
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(10).Value)
            GridMedico.Lock = True
            
            'Celular
            GridMedico.Col = 11
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(11).Value)
            GridMedico.Lock = True
            
            'Fax
            GridMedico.Col = 12
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(12).Value)
            GridMedico.Lock = True
            
            'Cpf
            GridMedico.Col = 13
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(13).Value)
            GridMedico.Lock = True
            
            'Email
            GridMedico.Col = 14
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(14).Value)
            GridMedico.Lock = True
            
            'Observação
            GridMedico.Col = 15
            GridMedico.Text = VerificaNulo(RecPesq.Fields.Item(15).Value)
            GridMedico.Lock = True
            
            'CodMed
            GridMedico.Col = 16
            GridMedico.Text = Val(RecPesq.Fields.Item(0).Value)
            GridMedico.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridMedico.MaxRows = GridMedico.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE MÉDICOS PESQUISADOS =========
         GridMedico.MaxRows = GridMedico.MaxRows - 1
         
         If GridMedico.MaxRows = 1 Then
            LblNumTotalMed.Caption = FormataNum(GridMedico.MaxRows) & " médico encontrado."
         Else
            LblNumTotalMed.Caption = FormataNum(GridMedico.MaxRows) & " médicos encontrados."
         End If
         '================================================
         
         CmdImprimirMed.Enabled = True
    End If

End Sub


Sub MontaGridFornecedor()

    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalForn.Caption = "Nenhum fornecedor encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridFornecedor.Refresh
        GridFornecedor.MaxRows = 0
        
        CmdAlterarForn.Enabled = False
        CmdExcluirForn.Enabled = False
        CmdImprimirForn.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridFornecedor.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridFornecedor.Row = VLIntLinha
            GridFornecedor.Lock = True
            
            'Fornecedor
            GridFornecedor.Col = 1
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridFornecedor.Lock = True
            
            'Tipo
            GridFornecedor.Col = 2
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(2).Value)
            GridFornecedor.Lock = True
            
            'Endereço
            GridFornecedor.Col = 3
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridFornecedor.Lock = True
            
            'Bairro
            GridFornecedor.Col = 4
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridFornecedor.Lock = True
            
            'Cep
            GridFornecedor.Col = 5
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(6).Value)
            GridFornecedor.Lock = True
            
            'Cidade
            GridFornecedor.Col = 6
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(7).Value)
            GridFornecedor.Lock = True
            
            'Estado
            GridFornecedor.Col = 7
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(8).Value)
            GridFornecedor.Lock = True
            
            'CNPJ
            GridFornecedor.Col = 8
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(9).Value)
            GridFornecedor.Lock = True
            
            'Email
            GridFornecedor.Col = 9
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(10).Value)
            GridFornecedor.Lock = True
            
            'Responsável
            GridFornecedor.Col = 10
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(11).Value)
            GridFornecedor.Lock = True
            
            'Telefone
            GridFornecedor.Col = 11
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(12).Value)
            GridFornecedor.Lock = True
            
            'Celular
            GridFornecedor.Col = 12
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(13).Value)
            GridFornecedor.Lock = True
            
            'Observação
            GridFornecedor.Col = 13
            GridFornecedor.Text = VerificaNulo(RecPesq.Fields.Item(14).Value)
            GridFornecedor.Lock = True
            
            'CodForn
            GridFornecedor.Col = 14
            GridFornecedor.Text = Val(RecPesq.Fields.Item(0).Value)
            GridFornecedor.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridFornecedor.MaxRows = GridFornecedor.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE FORNECEDORES PESQUISADOS =========
         GridFornecedor.MaxRows = GridFornecedor.MaxRows - 1
         
         If GridFornecedor.MaxRows = 1 Then
            LblNumTotalForn.Caption = FormataNum(GridFornecedor.MaxRows) & " fornecedor encontrado."
         Else
            LblNumTotalForn.Caption = FormataNum(GridFornecedor.MaxRows) & " fornecedores encontrados."
         End If
         '================================================
         
         CmdImprimirForn.Enabled = True
    End If

End Sub

Sub MontaGridOrcamento()
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalOrc.Caption = "Nenhum orçamento encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridOrcamento.Refresh
        GridOrcamento.MaxRows = 0
        
        CmdAlterarOrc.Enabled = False
        CmdExcluirOrc.Enabled = False
        CmdImprimirOrc.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridOrcamento.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridOrcamento.Row = VLIntLinha
            GridOrcamento.Lock = True
            
            'Data
            GridOrcamento.Col = 1
            GridOrcamento.Text = FormataData(RecPesq.Fields.Item(2).Value)
            GridOrcamento.Lock = True
            
            'Vendedor
            GridOrcamento.Col = 2
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(22).Value)
            GridOrcamento.Lock = True
            
            'Cliente
            GridOrcamento.Col = 3
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridOrcamento.Lock = True
            
            'Telefone
            GridOrcamento.Col = 4
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridOrcamento.Lock = True
            
            'Armação
            GridOrcamento.Col = 5
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridOrcamento.Lock = True
            
            'Valor armação
            GridOrcamento.Col = 6
            If RecPesq.Fields.Item(6).Value <> "" And IsNull(RecPesq.Fields.Item(6).Value) = False Then
                GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(6).Value)
            Else
                GridOrcamento.Text = ""
            End If
            GridOrcamento.Lock = True
            
            'Lente
            GridOrcamento.Col = 7
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(7).Value)
            GridOrcamento.Lock = True
            
            'Valor Lente
            GridOrcamento.Col = 8
            If RecPesq.Fields.Item(8).Value <> "" And IsNull(RecPesq.Fields.Item(8).Value) = False Then
                GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(8).Value)
            Else
                GridOrcamento.Text = ""
            End If
            GridOrcamento.Lock = True
            
            'Lente de contato
            GridOrcamento.Col = 9
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(9).Value)
            GridOrcamento.Lock = True
            
            'Valor lente de contato
            GridOrcamento.Col = 10
            If RecPesq.Fields.Item(10).Value <> "" And IsNull(RecPesq.Fields.Item(10).Value) = False Then
                GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(10).Value)
            Else
                GridOrcamento.Text = ""
            End If
            GridOrcamento.Lock = True
            
            'Outros
            GridOrcamento.Col = 11
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(11).Value)
            GridOrcamento.Lock = True
            
            'Valor outros
            GridOrcamento.Col = 12
            If RecPesq.Fields.Item(12).Value <> "" And IsNull(RecPesq.Fields.Item(12).Value) = False Then
                GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(12).Value)
            Else
                GridOrcamento.Text = ""
            End If
            GridOrcamento.Lock = True
            
            'Total da venda
            GridOrcamento.Col = 13
            GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(13).Value)
            GridOrcamento.Lock = True
            
            'Parcelado
            GridOrcamento.Col = 14
            GridOrcamento.Text = FormataNum(RecPesq.Fields.Item(14).Value) & " vezes"
            GridOrcamento.Lock = True
            
            'Entrada
            GridOrcamento.Col = 15
            GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(16).Value)
            GridOrcamento.Lock = True
            
            'Valor da parcela
            GridOrcamento.Col = 16
            GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(17).Value)
            GridOrcamento.Lock = True
            
            'Valor a prazo
            GridOrcamento.Col = 17
            GridOrcamento.Text = FormataMoeda(RecPesq.Fields.Item(18).Value)
            GridOrcamento.Lock = True
            
            'Validade
            GridOrcamento.Col = 18
            GridOrcamento.Text = FormataData(RecPesq.Fields.Item(19).Value)
            GridOrcamento.Lock = True
            
            'Observação
            GridOrcamento.Col = 19
            GridOrcamento.Text = VerificaNulo(RecPesq.Fields.Item(20).Value)
            GridOrcamento.Lock = True
            
            'CodOrc
            GridOrcamento.Col = 20
            GridOrcamento.Text = Val(RecPesq.Fields.Item(0).Value)
            GridOrcamento.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridOrcamento.MaxRows = GridOrcamento.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE CLIENTES PESQUISADOS =========
         GridOrcamento.MaxRows = GridOrcamento.MaxRows - 1
         
         If GridOrcamento.MaxRows = 1 Then
            LblNumTotalOrc.Caption = FormataNum(GridOrcamento.MaxRows) & " orçamento encontrado."
         Else
            LblNumTotalOrc.Caption = FormataNum(GridOrcamento.MaxRows) & " orçamentos encontrados."
         End If
         '================================================
         
         CmdImprimirOrc.Enabled = True
    End If

End Sub

Sub MontaGridProduto()
    Dim VLIntCodProd As Long
    Dim VLIntLinha As Long
    Dim RecGrif As New ADODB.Recordset
    
    If RecPesq.EOF Then
        LblNumTotalProd.Caption = "Nenhum produto encontrado."
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridProduto.Refresh
        GridProduto.MaxRows = 0
        
        CmdAlterarProd.Enabled = False
        CmdExcluirProd.Enabled = False
        CmdImprimirProd.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridProduto.MaxRows = VLIntLinha
        
        Do While Not RecPesq.EOF
                 
            GridProduto.Row = VLIntLinha
            GridProduto.Lock = True
            
            'Fornecedor
            GridProduto.Col = 1
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(9).Value)
            GridProduto.Lock = True
            
            'Tipo produto
            GridProduto.Col = 2
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(1).Value)
            GridProduto.Lock = True
            
            'Griffe
            GridProduto.Col = 3
            If IsNull(RecPesq.Fields.Item(10).Value) = False And RecPesq.Fields.Item(10).Value <> 0 Then
                StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecPesq.Fields.Item(10).Value
                RecGrif.Open StrSql, vgCon, 1, 3
                
                GridProduto.Text = VerificaNulo(RecGrif.Fields.Item(0).Value)
                RecGrif.Close
            Else
                GridProduto.Text = ""
            End If
            GridProduto.Lock = True
            
            'Cor
            GridProduto.Col = 4
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(2).Value)
            GridProduto.Lock = True
            
            'Número
            GridProduto.Col = 5
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridProduto.Lock = True
            
            'Modelo
            GridProduto.Col = 6
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridProduto.Lock = True
            
            'Tamanho Aro
            GridProduto.Col = 7
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(5).Value)
            GridProduto.Lock = True
            
            'Tamanho Ponte
            GridProduto.Col = 8
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(6).Value)
            GridProduto.Lock = True
            
            'Tipo de lente
            GridProduto.Col = 9
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(7).Value)
            GridProduto.Lock = True
            
            'Chave
            GridProduto.Col = 10
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(8).Value)
            GridProduto.Lock = True
            
            'CodProd
            GridProduto.Col = 11
            GridProduto.Text = Val(RecPesq.Fields.Item(0).Value)
            GridProduto.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridProduto.MaxRows = GridProduto.MaxRows + 1
            
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE CLIENTES PESQUISADOS =========
         GridProduto.MaxRows = GridProduto.MaxRows - 1
         
         If GridProduto.MaxRows = 1 Then
            LblNumTotalProd.Caption = FormataNum(GridProduto.MaxRows) & " produto encontrado."
         Else
            LblNumTotalProd.Caption = FormataNum(GridProduto.MaxRows) & " produtos encontrados."
         End If
         '================================================
         
         CmdImprimirProd.Enabled = True
    End If

End Sub

Sub MontaGridCaixa()
    Dim VLIntCodCx As Long
    Dim VLIntLinha As Long
    Dim VLIntCredito As Long
    Dim VLIntDebito As Long
    Dim VLIntVenda As Long
    Dim VLStrCorVermelho  As String
    
    VLStrCorVermelho = &HC0&
    
    If RecPesq.EOF Then
        LblNumTotalCx.Caption = "Nenhum movimento de caixa encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridCaixa.Refresh
        GridCaixa.MaxRows = 0
        
        CmdAlterarCx.Enabled = False
        CmdExcluirCx.Enabled = False
        CmdImprimirCx.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridCaixa.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
            
            GridCaixa.Row = VLIntLinha
            GridCaixa.Lock = True
            
            'Cod. Venda
            GridCaixa.Col = 1
            GridCaixa.Text = FormataNum(RecPesq.Fields.Item(1).Value)
            GridCaixa.Lock = True
                        
            'Descrição
            GridCaixa.Col = 2
            GridCaixa.TypeMaxEditLen = 255
            GridCaixa.Text = VerificaNulo(RecPesq.Fields.Item(6).Value)
            GridCaixa.Lock = True
            
            'Data Movimento
            GridCaixa.Col = 3
            GridCaixa.Text = FormataData(RecPesq.Fields.Item(2).Value)
            GridCaixa.Lock = True
            
            'Tipo Movimento
            GridCaixa.Col = 4
            GridCaixa.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridCaixa.Lock = True
            
            'Credito
            'MsgBox ("tipo credito=" & RecPesq.Fields.Item(5).Value)
            
            GridCaixa.Col = 5
            If RecPesq.Fields.Item(5).Value = "credito" Then
                'MsgBox ("valor credito=" & FormataMoeda(RecPesq.Fields.Item(4).Value))
                GridCaixa.Text = FormataMoeda(RecPesq.Fields.Item(4).Value)
                'MsgBox ("variavel=" & CCur(GridCaixa.Text))
                VLIntCredito = VLIntCredito + CCur(GridCaixa.Text)
                'MsgBox ("VLIntCredito=" & VLIntCredito)
                
                If RecPesq.Fields.Item(1).Value <> 0 And IsNull(RecPesq.Fields.Item(1).Value) = False Then
                    VLIntVenda = VLIntVenda + CCur(GridCaixa.Text)
                End If
            Else
                GridCaixa.Text = ""
            End If
            GridCaixa.Lock = True
            
            'Débito
            GridCaixa.Col = 6
            If RecPesq.Fields.Item(5).Value = "debito" Then
                GridCaixa.Text = FormataMoeda(RecPesq.Fields.Item(4).Value)
                VLIntDebito = VLIntDebito + CCur(GridCaixa.Text)
            Else
                GridCaixa.Text = ""
            End If
            GridCaixa.Lock = True
            
            'CodCx
            GridCaixa.Col = 7
            GridCaixa.Text = Val(RecPesq.Fields.Item(0).Value)
            GridCaixa.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridCaixa.MaxRows = GridCaixa.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         GridCaixa.Row = GridCaixa.MaxRows
         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True
         
         
         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows
         
         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL VENDA DO DIA:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntVenda)
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True
         
         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows
         
         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL CRÉDITO:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntCredito)
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True
         
         
         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows
         
         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL DÉBITO:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntDebito)
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True
         
         
         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows
         
         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL MOVIMENTO DO DIA:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntCredito - VLIntDebito)
         If InStr(GridCaixa.Text, "-") <> 0 Then
            GridCaixa.ForeColor = VLStrCorVermelho
         End If
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True
         
         
         '===== CONTAGEM DE MOVIMENTOS PESQUISADOS =========
         If (GridCaixa.MaxRows - 5) = 1 Then
            LblNumTotalCx.Caption = FormataNum((GridCaixa.MaxRows - 5)) & " movimento de caixa encontrado."
         Else
            LblNumTotalCx.Caption = FormataNum((GridCaixa.MaxRows - 5)) & " movimentos de caixa encontrados."
         End If
         '================================================
         
         CmdImprimirCx.Enabled = True
    End If

End Sub


Private Sub TxtCpfCli_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCpfMed_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCrmMed_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, letras minúsculas, letras maiúsculas e / - backspace e enter ===
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And KeyAscii <> 45 And KeyAscii <> 47 And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtCred1_GotFocus()
    TxtDtCred1.Text = ""
End Sub

Private Sub TxtDtCred2_GotFocus()
    TxtDtCred2.Text = ""
End Sub

Private Sub TxtDtMovCx1_Click()
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtMovCx1_GotFocus()
    If TxtDtMovCx1.Text = "__/__/____" Then
        TxtDtMovCx1.Text = ""
    End If
End Sub

Private Sub TxtDtMovCx1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtMovCx1_LostFocus()
    Dim VLStrData As String
    
    If TxtDtMovCx1.Text <> "" Then
        VLStrData = VerificaData(TxtDtMovCx1.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtMovCx1.SetFocus
        Else
            TxtDtMovCx1.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtMovCx1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtMovCx2_Click()
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtMovCx2_GotFocus()
    If TxtDtMovCx2.Text = "__/__/____" Then
        TxtDtMovCx2.Text = ""
    End If
End Sub

Private Sub TxtDtMovCx2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtDtMovCx2_LostFocus()
    Dim VLStrData As String
    
    If TxtDtMovCx2.Text <> "" Then
        VLStrData = VerificaData(TxtDtMovCx2.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtMovCx2.SetFocus
        Else
            TxtDtMovCx2.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtMovCx2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtNiver1_GotFocus()
    TxtDtNiver1.Text = ""
End Sub

Private Sub TxtDtNiver1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNiver1_LostFocus()

    Dim VLStrData As String

    If TxtDtNiver1.Text <> "" Then
        VLStrData = VerificaData(TxtDtNiver1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtNiver1.SetFocus
        Else
            TxtDtNiver1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtNiver1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtNiver2_GotFocus()
    TxtDtNiver2.Text = ""
End Sub

Private Sub TxtDtNiver2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNiver2_LostFocus()

    Dim VLStrData As String

    If TxtDtNiver2.Text <> "" Then
        VLStrData = VerificaData(TxtDtNiver2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtNiver2.SetFocus
        Else
            TxtDtNiver2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtNiver2.Text = "__/__/____"
    End If
End Sub


Private Sub TxtDtNiverCli1_GotFocus()
    TxtDtNiverCli1.Text = ""
End Sub

Private Sub TxtDtNiverCli1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNiverCli1_LostFocus()

    Dim VLStrData As String

    If TxtDtNiverCli1.Text <> "" Then
        VLStrData = VerificaData(TxtDtNiverCli1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtNiverCli1.SetFocus
        Else
            TxtDtNiverCli1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtNiverCli1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtNiverCli2_GotFocus()
    TxtDtNiverCli2.Text = ""
End Sub

Private Sub TxtDtNiverCli2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNiverCli2_LostFocus()

    Dim VLStrData As String

    If TxtDtNiverCli2.Text <> "" Then
        VLStrData = VerificaData(TxtDtNiverCli2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtNiverCli2.SetFocus
        Else
            TxtDtNiverCli2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtNiverCli2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtOrc1_GotFocus()
    If TxtDtOrc1.Text = "__/__/____" Then
        TxtDtOrc1.Text = ""
    End If
End Sub

Private Sub TxtDtOrc1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtOrc1_LostFocus()
    Dim VLStrData As String
    
    If TxtDtOrc1.Text <> "" Then
        VLStrData = VerificaData(TxtDtOrc1.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtOrc1.SetFocus
        Else
            TxtDtOrc1.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtOrc1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtOrc2_GotFocus()
    If TxtDtOrc2.Text = "__/__/____" Then
        TxtDtOrc2.Text = ""
    End If
End Sub

Private Sub TxtDtOrc2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtOrc2_LostFocus()
    Dim VLStrData As String
    
    If TxtDtOrc2.Text <> "" Then
        VLStrData = VerificaData(TxtDtOrc2.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtOrc2.SetFocus
        Else
            TxtDtOrc2.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtOrc2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenc1_GotFocus()
    TxtDtVenc1.Text = ""
End Sub

Private Sub TxtDtVenc1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc1_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtVenc1.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenc1.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVenc1.SetFocus
        Else
            TxtDtVenc1.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVenc1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenc2_GotFocus()
    TxtDtVenc2.Text = ""
End Sub

Private Sub TxtDtVenc2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc2_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtVenc2.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenc2.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVenc2.SetFocus
        Else
            TxtDtVenc2.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVenc2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVencCred1_GotFocus()
    TxtDtVencCred1.Text = ""
End Sub

Private Sub TxtDtVencCred2_GotFocus()
    TxtDtVencCred2.Text = ""
End Sub

Private Sub TxtDtCred1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtCred1_LostFocus()
    Dim VLStrData As String
    
    If TxtDtCred1.Text <> "" Then
        VLStrData = VerificaData(TxtDtCred1.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtCred1.SetFocus
        Else
            TxtDtCred1.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtCred1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtCred2_LostFocus()
    Dim VLStrData As String
    
    If TxtDtCred2.Text <> "" Then
        VLStrData = VerificaData(TxtDtCred2.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtCred2.SetFocus
        Else
            TxtDtCred2.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtCred2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVencCred1_LostFocus()
    Dim VLStrData As String
    
    If TxtDtVencCred1.Text <> "" Then
        VLStrData = VerificaData(TxtDtVencCred1.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVencCred1.SetFocus
        Else
            TxtDtVencCred1.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVencCred1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVencCred2_LostFocus()
    Dim VLStrData As String
    
    If TxtDtVencCred2.Text <> "" Then
        VLStrData = VerificaData(TxtDtVencCred2.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVencCred2.SetFocus
        Else
            TxtDtVencCred2.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVencCred2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtCred2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVencCred1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVencCred2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtRec1_GotFocus()
    TxtDtRec1.Text = ""
End Sub

Private Sub TxtDtRec1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtRec1_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtRec1.Text <> "" Then
        VLStrData = VerificaData(TxtDtRec1.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtRec1.SetFocus
        Else
            TxtDtRec1.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtRec1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtRec2_GotFocus()
    TxtDtRec2.Text = ""
End Sub

Private Sub TxtDtRec2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtRec2_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtRec2.Text <> "" Then
        VLStrData = VerificaData(TxtDtRec2.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtRec2.SetFocus
        Else
            TxtDtRec2.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtRec2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenda1_GotFocus()
    If TxtDtVenda1.Text = "__/__/____" Then
        TxtDtVenda1.Text = ""
    End If
End Sub

Private Sub TxtDtVenda1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenda1_LostFocus()
    Dim VLStrData As String
    
    If TxtDtVenda1.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenda1.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVenda1.SetFocus
        Else
            TxtDtVenda1.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVenda1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenda2_GotFocus()
    If TxtDtVenda2.Text = "__/__/____" Then
        TxtDtVenda2.Text = ""
    End If
End Sub

Private Sub TxtDtVenda2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenda2_LostFocus()
    Dim VLStrData As String
    
    If TxtDtVenda2.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenda2.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVenda2.SetFocus
        Else
            TxtDtVenda2.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVenda2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDia1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDia2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtMes1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtMes2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtQtdeMinEst_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTelCli_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTelForn_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTelMed_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Sub MontaCboProdEst()
    Conecta
    
    Dim RecProd As New ADODB.Recordset
    
    StrSql = "Select distinct TipoProd From tb_Produto"
    RecProd.Open StrSql, vgCon, 1, 3
    
    CboProdEst.AddItem ("")
    
    Do While Not RecProd.EOF
        CboProdEst.AddItem (RecProd.Fields.Item(0).Value)
    RecProd.MoveNext
    Loop
    
    Desconecta
End Sub

Sub MontaCbosProd()
    Conecta
    
    Dim RecForn As New ADODB.Recordset
    Dim RecGrif As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    Dim RecLente As New ADODB.Recordset
    
    StrSql = "Select distinct Nome From tb_Fornecedor"
    RecForn.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct Nome From tb_Griffe"
    RecGrif.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct TipoProd From tb_Produto"
    RecProd.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct Tipo From tb_Produto"
    RecLente.Open StrSql, vgCon, 1, 3
    
    CboFornProd.AddItem ("")
    Do While Not RecForn.EOF
        CboFornProd.AddItem (RecForn.Fields.Item(0).Value)
    RecForn.MoveNext
    Loop
    
    CboGriffeProd.AddItem ("")
    Do While Not RecGrif.EOF
        CboGriffeProd.AddItem (RecGrif.Fields.Item(0).Value)
    RecGrif.MoveNext
    Loop
    
    CboTipoProd.AddItem ("")
    Do While Not RecProd.EOF
        CboTipoProd.AddItem (RecProd.Fields.Item(0).Value)
    RecProd.MoveNext
    Loop
    
    CboLenteProd.AddItem ("")
    Do While Not RecLente.EOF
        If RecLente.Fields.Item(0).Value <> "" And IsNull(RecLente.Fields.Item(0).Value) = False Then
            CboLenteProd.AddItem (RecLente.Fields.Item(0).Value)
        End If
    RecLente.MoveNext
    Loop
    
    Desconecta
End Sub

Sub MontaCboTipoCred()
    Conecta
    
    Dim RecTipo As New ADODB.Recordset
    
    StrSql = "Select distinct TipoCred From tb_Crediario"
    RecTipo.Open StrSql, vgCon, 1, 3
    
    CboTipoCred.AddItem ("")
    Do While Not RecTipo.EOF
        CboTipoCred.AddItem (RecTipo.Fields.Item(0).Value)
        RecTipo.MoveNext
    Loop
    
    Desconecta
End Sub

Sub MontaCboTipoPagtoCX()
    Conecta
    
    Dim RecTipo As New ADODB.Recordset
    StrSql = "Select distinct TipoPagto From tb_Caixa order by TipoPagto"
    RecTipo.Open StrSql, vgCon, 1, 3
    
    CboTipoPagtoCx.AddItem ("")
    Do While Not RecTipo.EOF
        CboTipoPagtoCx.AddItem (RecTipo.Fields.Item(0).Value)
        RecTipo.MoveNext
    Loop
    Desconecta
End Sub

Sub MontaCboTipoVenda()
    Conecta
    
    Dim RecTipo As New ADODB.Recordset
    
    StrSql = "Select distinct TipoVenda From tb_Venda order by TipoVenda"
    RecTipo.Open StrSql, vgCon, 1, 3
    
    CboTipoVenda.AddItem ("")
    Do While Not RecTipo.EOF
        CboTipoVenda.AddItem (RecTipo.Fields.Item(0).Value)
        RecTipo.MoveNext
    Loop
    
    Desconecta
End Sub

Private Sub TxtTelOrc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Sub MontaGriffe()
    Conecta
    
    Dim RecGriffe As New ADODB.Recordset
    
    StrSql = "Select G.CodGriffe,G.Nome From tb_Griffe as G,tb_Produto as P where G.CodGriffe=P.CodGriffe"
    RecGriffe.Open StrSql, vgCon, 1, 3
    
    Do While Not RecGriffe.EOF
        CboGriffe.AddItem (RecGriffe.Fields.Item(1).Value & "                                                                                                         " & RecGriffe.Fields.Item(0).Value)
        RecGriffe.MoveNext
    Loop
    
    Desconecta
End Sub

Private Sub TxtDtVendaProp01_GotFocus()
    TxtDtVendaProp01.Text = ""
End Sub

Private Sub TxtDtVendaProp01_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVendaProp01_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtVendaProp01.Text <> "" Then
        VLStrData = VerificaData(TxtDtVendaProp01.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVendaProp01.SetFocus
        Else
            TxtDtVendaProp01.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVendaProp01.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVendaProp02_GotFocus()
    TxtDtVendaProp02.Text = ""
End Sub

Private Sub TxtDtVendaProp02_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVendaProp02_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtVendaProp02.Text <> "" Then
        VLStrData = VerificaData(TxtDtVendaProp02.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVendaProp02.SetFocus
        Else
            TxtDtVendaProp02.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVendaProp02.Text = "__/__/____"
    End If
End Sub
