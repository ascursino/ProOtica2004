VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmVenda_Inc 
   Caption         =   "Inclusão de Venda"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
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
   Icon            =   "FrmVenda_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   7920
   Begin VB.Frame FraVista 
      Caption         =   "À vista"
      Height          =   2895
      Left            =   240
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox TxtTotalVista 
         Height          =   285
         Left            =   1200
         TabIndex        =   99
         ToolTipText     =   "Valor total da venda"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TxtVendaVista 
         Height          =   285
         Left            =   1200
         TabIndex        =   98
         ToolTipText     =   "Valor da venda"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtDescVista 
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         ToolTipText     =   "Desconto na venda"
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton OptDin 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         ToolTipText     =   "Pagamento em dinheiro"
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptChq 
         Caption         =   "Cheque"
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         ToolTipText     =   "Pagamento em cheque"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtDigVista 
         Height          =   285
         Left            =   6960
         TabIndex        =   28
         ToolTipText     =   "Dígito do número do cheque"
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox TxtChequeVista 
         Height          =   285
         Left            =   5640
         TabIndex        =   27
         ToolTipText     =   "Número do cheque"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtBancoVista 
         Height          =   285
         Left            =   5640
         TabIndex        =   26
         ToolTipText     =   "Número do banco do cheque"
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmVenda_Inc.frx":0CCA
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmVenda_Inc.frx":0D2E
         TabIndex        =   33
         Top             =   960
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmVenda_Inc.frx":0D98
         TabIndex        =   34
         Top             =   1440
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmVenda_Inc.frx":0DFC
         TabIndex        =   35
         Top             =   960
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "FrmVenda_Inc.frx":0E56
         TabIndex        =   36
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "FrmVenda_Inc.frx":0EBA
         TabIndex        =   37
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Frame FraPrazoCheque 
      Caption         =   "A prazo - cheque"
      Height          =   2895
      Left            =   240
      TabIndex        =   38
      Top             =   3240
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton CmdParcInc 
         Caption         =   "Incluir parcelas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   97
         ToolTipText     =   "Incluir parcelas do crediário"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton CmdCrediarista 
         Caption         =   "Crediarista"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   96
         ToolTipText     =   "Escolher crediarista"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Entrada"
         Height          =   1455
         Left            =   2880
         TabIndex        =   80
         Top             =   720
         Width           =   4455
         Begin VB.Frame FraChqEntrDin 
            Height          =   1215
            Left            =   1680
            TabIndex        =   91
            Top             =   120
            Visible         =   0   'False
            Width           =   2655
            Begin VB.TextBox TxtValorEntrDinChq 
               Height          =   285
               Left            =   1080
               TabIndex        =   92
               ToolTipText     =   "Valor da entrada"
               Top             =   480
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel53 
               Height          =   255
               Left            =   360
               OleObjectBlob   =   "FrmVenda_Inc.frx":0F20
               TabIndex        =   93
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.OptionButton OptChqSemEntr 
            Caption         =   "Sem entrada"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            ToolTipText     =   "Venda sem entrada"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Frame FraChqEntrChq 
            Height          =   1215
            Left            =   1680
            TabIndex        =   83
            Top             =   120
            Visible         =   0   'False
            Width           =   2655
            Begin VB.TextBox TxtValorEntrChequeChq 
               Height          =   285
               Left            =   960
               TabIndex        =   94
               ToolTipText     =   "Valor da entrada"
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox TxtDigChq 
               Height          =   285
               Left            =   2280
               TabIndex        =   86
               ToolTipText     =   "Dígito do número do cheque"
               Top             =   480
               Width           =   255
            End
            Begin VB.TextBox TxtChequeChq 
               Height          =   285
               Left            =   960
               TabIndex        =   85
               ToolTipText     =   "Número do cheque"
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox TxtBancoChq 
               Height          =   285
               Left            =   960
               TabIndex        =   84
               ToolTipText     =   "Nùmero do banco do cheque"
               Top             =   120
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel50 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmVenda_Inc.frx":0F84
               TabIndex        =   87
               Top             =   120
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmVenda_Inc.frx":0FE8
               TabIndex        =   88
               Top             =   480
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel52 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "FrmVenda_Inc.frx":104E
               TabIndex        =   89
               Top             =   840
               Width           =   615
            End
         End
         Begin VB.OptionButton OptChqEntrChq 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "Pagamento da entrada em cheque"
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton OptChqEntrDin 
            Caption         =   "Dinheiro"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            ToolTipText     =   "Pagamento da entrada em dinheiro"
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox TxtVendaChq 
         Height          =   285
         Left            =   1080
         TabIndex        =   79
         ToolTipText     =   "Valor da venda"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxtTotalVendaChq 
         Height          =   285
         Left            =   1080
         TabIndex        =   78
         ToolTipText     =   "Valor total da venda"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox CboPrazoChqParc 
         Height          =   315
         ItemData        =   "FrmVenda_Inc.frx":10B2
         Left            =   1080
         List            =   "FrmVenda_Inc.frx":10B4
         Style           =   2  'Dropdown List
         TabIndex        =   40
         ToolTipText     =   "Quantidade de parcelas"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox TxtPrazoChqJuros 
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         ToolTipText     =   "Juros da venda"
         Top             =   1200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":10B6
         TabIndex        =   41
         Top             =   840
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":111A
         TabIndex        =   42
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":1184
         TabIndex        =   43
         Top             =   1200
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":11E8
         TabIndex        =   44
         Top             =   1560
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEntrChq 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":124C
         TabIndex        =   45
         Top             =   2280
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblParcChq 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":12C8
         TabIndex        =   46
         Top             =   2520
         Width           =   5055
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCredstaCheque 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":1354
         TabIndex        =   95
         Top             =   360
         Width           =   5775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "FrmVenda_Inc.frx":13C4
         TabIndex        =   101
         Top             =   1200
         Width           =   255
      End
   End
   Begin VB.Frame FraFinalizVenda 
      Caption         =   "Finalização da venda"
      Height          =   3735
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   7695
      Begin VB.Frame FraPrazoCarne 
         Caption         =   "A prazo - carnê"
         Height          =   2895
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Visible         =   0   'False
         Width           =   7455
         Begin VB.CommandButton CmdCrediarista 
            Caption         =   "Crediarista"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   6120
            TabIndex        =   77
            ToolTipText     =   "Escolher crediarista"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Frame Frame5 
            Caption         =   "Entrada"
            Height          =   1455
            Left            =   2880
            TabIndex        =   55
            Top             =   720
            Width           =   4455
            Begin VB.Frame FraCarEntrDin 
               Height          =   1215
               Left            =   1680
               TabIndex        =   66
               Top             =   120
               Visible         =   0   'False
               Width           =   2655
               Begin VB.TextBox TxtValorEntrDinCar 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   75
                  ToolTipText     =   "Valor da entrada"
                  Top             =   480
                  Width           =   1215
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel64 
                  Height          =   255
                  Left            =   360
                  OleObjectBlob   =   "FrmVenda_Inc.frx":141E
                  TabIndex        =   67
                  Top             =   480
                  Width           =   615
               End
            End
            Begin VB.OptionButton OptCarSemEntr 
               Caption         =   "Sem entrada"
               Height          =   255
               Left            =   120
               TabIndex        =   65
               ToolTipText     =   "Venda sem entrada"
               Top             =   1080
               Width           =   1455
            End
            Begin VB.OptionButton OptCarEntrDin 
               Caption         =   "Dinheiro"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               ToolTipText     =   "Pagamento da entrada em dinheiro"
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton OptCarEntrChq 
               Caption         =   "Cheque"
               Height          =   255
               Left            =   120
               TabIndex        =   63
               ToolTipText     =   "Pagamento da entrada em cheque"
               Top             =   720
               Width           =   975
            End
            Begin VB.Frame FraCarEntrChq 
               Height          =   1215
               Left            =   1680
               TabIndex        =   56
               Top             =   120
               Visible         =   0   'False
               Width           =   2655
               Begin VB.TextBox TxtValorEntrChequeCar 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   74
                  ToolTipText     =   "Valor da entrada"
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.TextBox TxtBancoCar 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   59
                  ToolTipText     =   "Número do banco do cheque"
                  Top             =   120
                  Width           =   495
               End
               Begin VB.TextBox TxtChequeCar 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   58
                  ToolTipText     =   "Número do cheque"
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.TextBox TxtDigCar 
                  Height          =   285
                  Left            =   2280
                  TabIndex        =   57
                  ToolTipText     =   "Dígito do número do cheque"
                  Top             =   480
                  Width           =   255
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel65 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":1482
                  TabIndex        =   60
                  Top             =   120
                  Width           =   615
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel66 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":14E6
                  TabIndex        =   61
                  Top             =   480
                  Width           =   735
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel67 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":154C
                  TabIndex        =   62
                  Top             =   840
                  Width           =   615
               End
            End
         End
         Begin VB.TextBox TxtPrazoCarJuros 
            Height          =   285
            Left            =   1080
            TabIndex        =   54
            ToolTipText     =   "Juros da venda"
            Top             =   1200
            Width           =   495
         End
         Begin VB.ComboBox CboPrazoCarParc 
            Height          =   315
            ItemData        =   "FrmVenda_Inc.frx":15B0
            Left            =   1080
            List            =   "FrmVenda_Inc.frx":15B2
            Style           =   2  'Dropdown List
            TabIndex        =   53
            ToolTipText     =   "Quantidade de parcelas"
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox TxtVendaCar 
            Height          =   285
            Left            =   1080
            TabIndex        =   52
            ToolTipText     =   "Valor da venda"
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TxtTotalVendaCar 
            Height          =   285
            Left            =   1080
            TabIndex        =   51
            ToolTipText     =   "Valor total da venda"
            Top             =   1560
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel58 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":15B4
            TabIndex        =   68
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel59 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1618
            TabIndex        =   69
            Top             =   1920
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel60 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1682
            TabIndex        =   70
            Top             =   1200
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel61 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":16E6
            TabIndex        =   71
            Top             =   1560
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblEntrCar 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":174A
            TabIndex        =   72
            Top             =   2280
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblParcCar 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":17C6
            TabIndex        =   73
            Top             =   2520
            Width           =   6495
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblCredstaCarne 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1852
            TabIndex        =   76
            Top             =   360
            Width           =   5775
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "FrmVenda_Inc.frx":18C2
            TabIndex        =   102
            Top             =   1200
            Width           =   255
         End
      End
      Begin VB.Frame FraTipoPrazo 
         Height          =   495
         Left            =   4680
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.OptionButton OptPrazoCarne 
            Caption         =   "Carnê"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "Venda a prazo em carnê"
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton OptPrazoCheque 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   1080
            TabIndex        =   48
            ToolTipText     =   "Venda a prazo em cheque"
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.OptionButton OptPrazo 
         Caption         =   "A prazo"
         Height          =   255
         Left            =   3600
         TabIndex        =   18
         ToolTipText     =   "Venda a prazo"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptVista 
         Caption         =   "À vista"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         ToolTipText     =   "Venda à vista"
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Inc.frx":191C
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FraProduto 
      Caption         =   "Produto 01"
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton CmdLimparProd 
         Caption         =   "Limpar produto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   100
         ToolTipText     =   "Limpar produto"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdProx 
         Caption         =   "Próximo >>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   24
         ToolTipText     =   "Próximo produto"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton CmdAnt 
         Caption         =   "<< Anterior"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   23
         ToolTipText     =   "Produto anterior"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TxtQtdeProd 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Quantidade do produto"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtDescrProd 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Código do produto"
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox TxtPrecoUnit 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Preço unitário do produto"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtValorTotal 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         ToolTipText     =   "Valor total do produto"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton CmdVerProd 
         Caption         =   "Ver produto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   1
         ToolTipText     =   "Ver produto"
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":1990
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":19F8
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":1A6E
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":1ADC
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtdeProdEst 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmVenda_Inc.frx":1B4C
         TabIndex        =   15
         Top             =   1440
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc.frx":1BC0
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoProd 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmVenda_Inc.frx":1C32
         TabIndex        =   21
         Top             =   360
         Width           =   4815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtdeProdEstTemp 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmVenda_Inc.frx":1C8E
         TabIndex        =   22
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.ComboBox CboVendedor 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   8
      ToolTipText     =   "Nome do vendedor"
      Top             =   6360
      Width           =   6495
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
      TabIndex        =   7
      Top             =   6720
      Width           =   7695
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   240
         OleObjectBlob   =   "FrmVenda_Inc.frx":1CEA
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
         Left            =   6360
         TabIndex        =   6
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
         Left            =   5040
         TabIndex        =   5
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmVenda_Inc.frx":1F1E
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
End
Attribute VB_Name = "FrmVenda_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public VPIntParcela As Integer
Public VPIntParcTemp As Integer

Public VPIntProd As Integer
Public VPStrVenda As String

Public VPStrTipoProd01 As String
Public VPStrDescrProd01 As String
Public VPStrPrecoUnit01 As String
Public VPStrQtdeProd01 As String
Public VPStrValorTotal01 As String
Public VPStrQtdeProdEst01 As String

Public VPStrTipoProd02 As String
Public VPStrDescrProd02 As String
Public VPStrPrecoUnit02 As String
Public VPStrQtdeProd02 As String
Public VPStrValorTotal02 As String
Public VPStrQtdeProdEst02 As String

Public VPStrTipoProd03 As String
Public VPStrDescrProd03 As String
Public VPStrPrecoUnit03 As String
Public VPStrQtdeProd03 As String
Public VPStrValorTotal03 As String
Public VPStrQtdeProdEst03 As String

Public VPStrTipoProd04 As String
Public VPStrDescrProd04 As String
Public VPStrPrecoUnit04 As String
Public VPStrQtdeProd04 As String
Public VPStrValorTotal04 As String
Public VPStrQtdeProdEst04 As String

Public VPStrTipoProd05 As String
Public VPStrDescrProd05 As String
Public VPStrPrecoUnit05 As String
Public VPStrQtdeProd05 As String
Public VPStrValorTotal05 As String
Public VPStrQtdeProdEst05 As String

Public VPStrTipoProd06 As String
Public VPStrDescrProd06 As String
Public VPStrPrecoUnit06 As String
Public VPStrQtdeProd06 As String
Public VPStrValorTotal06 As String
Public VPStrQtdeProdEst06 As String

Public VPIntCodCredTemp As Long

Private Sub CboPrazoCarParc_Click()
    Dim restparc As String
    
    If TxtTotalVendaCar.Text = "" Then
        TxtTotalVendaCar.Text = TxtVendaCar.Text
    End If
    
    If TxtTotalVendaCar.Text <> "" Then
        restparc = CCur(TxtTotalVendaCar.Text) - ((CCur(TxtTotalVendaCar.Text) * 20) / 100)
        
        LblEntrCar.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaCar.Text) * 20) / 100))
        
        If restparc = "0" And CboPrazoCarParc.Text = "00" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoCarParc.Text))
        End If
    End If
End Sub

Private Sub CboPrazoChqParc_Click()
    Dim restparc As String
    
    If TxtTotalVendaChq.Text = "" Then
        TxtTotalVendaChq.Text = TxtVendaChq.Text
    End If
    
    If TxtTotalVendaChq.Text <> "" Then
        restparc = CCur(TxtTotalVendaChq.Text) - ((CCur(TxtTotalVendaChq.Text) * 20) / 100)
        
        LblEntrChq.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaChq.Text) * 20) / 100))
        
        If restparc = "0" And CboPrazoChqParc.Text = "00" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoChqParc.Text))
        End If
    
    End If
End Sub

Private Sub CmdAnt_Click()
    VPIntProd = VPIntProd - 1
    
    If VPStrValorTotal01 = "0" Then
        VPStrValorTotal01 = ""
    End If
    
    If VPStrValorTotal02 = "0" Then
        VPStrValorTotal02 = ""
    End If
    
    If VPStrValorTotal03 = "0" Then
        VPStrValorTotal03 = ""
    End If
    
    If VPStrValorTotal04 = "0" Then
        VPStrValorTotal04 = ""
    End If
    
    If VPStrValorTotal05 = "0" Then
        VPStrValorTotal05 = ""
    End If
    
    If VPStrValorTotal06 = "0" Then
        VPStrValorTotal06 = ""
    End If
    
    FraProduto.Caption = "Produto " & FormataNum(VPIntProd)
    
    If VPIntProd = 1 Then
        CmdAnt.Enabled = False
        
        VPStrTipoProd02 = LblTipoProd.Caption
        VPStrDescrProd02 = TxtDescrProd.Text
        VPStrPrecoUnit02 = TxtPrecoUnit.Text
        VPStrQtdeProd02 = TxtQtdeProd.Text
        VPStrValorTotal02 = TxtValorTotal.Text
        VPStrQtdeProdEst02 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd01
        TxtDescrProd.Text = VPStrDescrProd01
        TxtPrecoUnit.Text = VPStrPrecoUnit01
        TxtQtdeProd.Text = VPStrQtdeProd01
        TxtValorTotal.Text = VPStrValorTotal01
        LblQtdeProdEst.Caption = VPStrQtdeProdEst01
        
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 2 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd03 = LblTipoProd.Caption
        VPStrDescrProd03 = TxtDescrProd.Text
        VPStrPrecoUnit03 = TxtPrecoUnit.Text
        VPStrQtdeProd03 = TxtQtdeProd.Text
        VPStrValorTotal03 = TxtValorTotal.Text
        VPStrQtdeProdEst03 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd02
        TxtDescrProd.Text = VPStrDescrProd02
        TxtPrecoUnit.Text = VPStrPrecoUnit02
        TxtQtdeProd.Text = VPStrQtdeProd02
        TxtValorTotal.Text = VPStrValorTotal02
        LblQtdeProdEst.Caption = VPStrQtdeProdEst02
        
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 3 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd04 = LblTipoProd.Caption
        VPStrDescrProd04 = TxtDescrProd.Text
        VPStrPrecoUnit04 = TxtPrecoUnit.Text
        VPStrQtdeProd04 = TxtQtdeProd.Text
        VPStrValorTotal04 = TxtValorTotal.Text
        VPStrQtdeProdEst04 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd03
        TxtDescrProd.Text = VPStrDescrProd03
        TxtPrecoUnit.Text = VPStrPrecoUnit03
        TxtQtdeProd.Text = VPStrQtdeProd03
        TxtValorTotal.Text = VPStrValorTotal03
        LblQtdeProdEst.Caption = VPStrQtdeProdEst03
        
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 4 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd05 = LblTipoProd.Caption
        VPStrDescrProd05 = TxtDescrProd.Text
        VPStrPrecoUnit05 = TxtPrecoUnit.Text
        VPStrQtdeProd05 = TxtQtdeProd.Text
        VPStrValorTotal05 = TxtValorTotal.Text
        VPStrQtdeProdEst05 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd04
        TxtDescrProd.Text = VPStrDescrProd04
        TxtPrecoUnit.Text = VPStrPrecoUnit04
        TxtQtdeProd.Text = VPStrQtdeProd04
        TxtValorTotal.Text = VPStrValorTotal04
        LblQtdeProdEst.Caption = VPStrQtdeProdEst04
        
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 5 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd06 = LblTipoProd.Caption
        VPStrDescrProd06 = TxtDescrProd.Text
        VPStrPrecoUnit06 = TxtPrecoUnit.Text
        VPStrQtdeProd06 = TxtQtdeProd.Text
        VPStrValorTotal06 = TxtValorTotal.Text
        VPStrQtdeProdEst06 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd05
        TxtDescrProd.Text = VPStrDescrProd05
        TxtPrecoUnit.Text = VPStrPrecoUnit05
        TxtQtdeProd.Text = VPStrQtdeProd05
        TxtValorTotal.Text = VPStrValorTotal05
        LblQtdeProdEst.Caption = VPStrQtdeProdEst05

        CmdProx.Enabled = True
    End If
End Sub

Private Sub CmdCrediarista_Click(Index As Integer)
    VGStrCredLista = "venda"
    FrmCrediarista_Lista.Show
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdLimparProd_Click()
    LblTipoProd.Caption = ""
    TxtDescrProd.Text = ""
    TxtPrecoUnit.Text = ""
    TxtQtdeProd.Text = ""
    TxtValorTotal.Text = ""
    LblQtdeProdEst.Caption = ""
    
    TxtValorTotal.SetFocus
End Sub

Private Sub CmdOK_Click()
    If OptPrazo.Value = True And LblCredstaCheque.Caption = "Crediarista:" And LblCredstaCarne.Caption = "Crediarista:" Then
        VPStrBox = MsgBox("Você deve escolher um crediarista para este crediário.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        Screen.MousePointer = vbHourglass
        
        Dim RecVenda As New ADODB.Recordset
        Dim RecEst As New ADODB.Recordset
        Dim RecCx As New ADODB.Recordset
        Dim RecCred As New ADODB.Recordset
        Dim RecCredParc As New ADODB.Recordset
        
        Dim VLIntCountProd As Integer
        Dim VLIntCodTemp As Long
        Dim VLIntCodVendaTemp As Long
        Dim parcelatemp As Integer
        Dim VLIntCodCred As Long
        
        parcelatemp = 1
                
        Conecta
        
        If VPStrVenda = "vista" Then
            
            '============== INCLUIR VENDA =============================
            StrSql = "SELECT * FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
                
            RecVenda.AddNew
            If CboVendedor.Text = "" Then
                RecVenda("CodVendedor") = 0
            Else
                RecVenda("CodVendedor") = Trim(Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10))
            End If
            RecVenda("CodCli") = VGIntCodCli
            RecVenda("CodCred") = 0
            RecVenda("DtVenda") = FormataDataUS(Date)
            If VPStrDescrProd01 <> "" Then
                RecVenda("CodProd01") = Mid(VPStrDescrProd01, 1, InStr(VPStrDescrProd01, "/") - 1)
            Else
                RecVenda("CodProd01") = 0
            End If
            If VPStrDescrProd02 <> "" Then
                RecVenda("CodProd02") = Mid(VPStrDescrProd02, 1, InStr(VPStrDescrProd02, "/") - 1)
            Else
                RecVenda("CodProd02") = 0
            End If
            
            If VPStrDescrProd03 <> "" Then
                RecVenda("CodProd03") = Mid(VPStrDescrProd03, 1, InStr(VPStrDescrProd03, "/") - 1)
            Else
                RecVenda("CodProd03") = 0
            End If
            
            If VPStrDescrProd04 <> "" Then
                RecVenda("CodProd04") = Mid(VPStrDescrProd04, 1, InStr(VPStrDescrProd04, "/") - 1)
            Else
                RecVenda("CodProd04") = 0
            End If
            
            If VPStrDescrProd05 <> "" Then
                RecVenda("CodProd05") = Mid(VPStrDescrProd05, 1, InStr(VPStrDescrProd05, "/") - 1)
            Else
                RecVenda("CodProd05") = 0
            End If
            
            If VPStrDescrProd06 <> "" Then
                RecVenda("CodProd06") = Mid(VPStrDescrProd06, 1, InStr(VPStrDescrProd06, "/") - 1)
            Else
                RecVenda("CodProd06") = 0
            End If
            
            RecVenda("CodForn01") = 0
            RecVenda("CodForn02") = 0
            RecVenda("CodForn03") = 0
            RecVenda("CodForn04") = 0
            RecVenda("CodForn05") = 0
            RecVenda("CodForn06") = 0
            RecVenda("Qtde01") = Val(VPStrQtdeProd01)
            RecVenda("Qtde02") = Val(VPStrQtdeProd02)
            RecVenda("Qtde03") = Val(VPStrQtdeProd03)
            RecVenda("Qtde04") = Val(VPStrQtdeProd04)
            RecVenda("Qtde05") = Val(VPStrQtdeProd05)
            RecVenda("Qtde06") = Val(VPStrQtdeProd06)
            
            If VPStrValorTotal01 <> "" Then
                RecVenda("ValorVenda01") = CCur(VPStrValorTotal01)
            Else
                RecVenda("ValorVenda01") = ""
            End If
            
            If VPStrValorTotal02 <> "" Then
                RecVenda("ValorVenda02") = CCur(VPStrValorTotal02)
            Else
                RecVenda("ValorVenda02") = ""
            End If
            
            If VPStrValorTotal03 <> "" Then
                RecVenda("ValorVenda03") = CCur(VPStrValorTotal03)
            Else
                RecVenda("ValorVenda03") = ""
            End If
            
            If VPStrValorTotal04 <> "" Then
                RecVenda("ValorVenda04") = CCur(VPStrValorTotal04)
            Else
                RecVenda("ValorVenda04") = ""
            End If
            
            If VPStrValorTotal05 <> "" Then
                RecVenda("ValorVenda05") = CCur(VPStrValorTotal05)
            Else
                RecVenda("ValorVenda05") = ""
            End If
            
            If VPStrValorTotal06 <> "" Then
                RecVenda("ValorVenda06") = CCur(VPStrValorTotal06)
            Else
                RecVenda("ValorVenda06") = ""
            End If
            
            RecVenda("TipoVenda") = "À vista"
            RecVenda("SubTotalVenda") = CCur(TxtVendaVista.Text)
            RecVenda("Desconto") = TxtDescVista.Text
            RecVenda("TotalVenda") = CCur(TxtTotalVista.Text)
            
            If OptDin.Value = True Then
                RecVenda("TipoPagto") = "Dinheiro"
                RecVenda("NumBanco") = 0
                RecVenda("NumCheque") = ""
            
            ElseIf OptChq.Value = True Then
                RecVenda("TipoPagto") = "Cheque"
                RecVenda("NumBanco") = TxtBancoVista.Text
                If TxtDigVista.Text <> "" Then
                    RecVenda("NumCheque") = TxtChequeVista.Text & "-" & TxtDigVista.Text
                Else
                    RecVenda("NumCheque") = TxtChequeVista.Text
                End If
            End If
            RecVenda.Update
                
            RecVenda.Close
            
            StrSql = "SELECT MAX(CodVenda) FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
            
            VLIntCodVendaTemp = RecVenda.Fields.Item(0).Value
            VGIntCodVendaRel = RecVenda.Fields.Item(0).Value
            
            RecVenda.Close
            
            '================ RETIRAR QTDE DO ESTOQUE ========================
            VLIntCountProd = 1
            VLIntCodTemp = 0
            
            Do While VLIntCountProd <= 6
                
                If VLIntCountProd = 1 Then
                    If VPStrDescrProd01 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd01, 1, InStr(VPStrDescrProd01, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd01)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 2 Then
                    If VPStrDescrProd02 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd02, 1, InStr(VPStrDescrProd02, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd02)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 3 Then
                    If VPStrDescrProd03 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd03, 1, InStr(VPStrDescrProd03, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd03)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 4 Then
                    If VPStrDescrProd04 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd04, 1, InStr(VPStrDescrProd04, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd04)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 5 Then
                    If VPStrDescrProd05 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd05, 1, InStr(VPStrDescrProd05, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd05)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 6 Then
                    If VPStrDescrProd06 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd06, 1, InStr(VPStrDescrProd06, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd06)
                        RecEst.Update
                        RecEst.Close
                    End If
                End If
                
                VLIntCodTemp = 0
                VLIntCountProd = VLIntCountProd + 1
            Loop
                
            '================= INCLUIR NO MOVIMENTO DE CAIXA =========================
            
            StrSql = "SELECT * FROM tb_Caixa"
            RecCx.Open StrSql, vgCon, 1, 3
            
            RecCx.AddNew
            RecCx("CodVenda") = VLIntCodVendaTemp
            RecCx("DtMov") = FormataDataUS(Date)
            RecCx("TipoMov") = "Venda à vista"
            RecCx("Valor") = CCur(TxtTotalVista.Text)
            RecCx("TipoValor") = "credito"
            RecCx("Descricao") = "Venda à vista - Cliente: " & VGStrNomeCli
            
            If OptDin.Value = True Then
                RecCx("TipoPagto") = "Dinheiro"
            ElseIf OptChq.Value = True Then
                RecCx("TipoPagto") = "Cheque"
            End If
            
            RecCx.Update
        
            Desconecta
            
            VPStrBox = MsgBox("Venda efetuada.", vbInformation, "Pró Ótica 2004 - Informação")
            
            Unload Me
            
            MDIPrincipal.Enabled = True
            MDIPrincipal.WindowState = 2
            
            Screen.MousePointer = vbNormal
        
        ElseIf VPStrVenda = "prazocheque" Then
            
            '============== INCLUIR CREDIÁRIO =============================
            StrSql = "SELECT * FROM tb_Crediario"
            RecCred.Open StrSql, vgCon, 1, 3
            
            RecCred.AddNew
            RecCred("CodCredsta") = VGIntCodCredstaVenda
            RecCred("CodCli") = VGIntCodCli
            RecCred("DtCred") = FormataDataUS(Date)
            RecCred("TipoCred") = "Cheque"
            RecCred("ValorVenda") = CCur(TxtVendaChq.Text)
            RecCred("Parcela") = CboPrazoChqParc.Text
            RecCred("Juros") = TxtPrazoChqJuros.Text
            RecCred("ValorTotal") = CCur(TxtTotalVendaChq.Text)
            
            If OptChqSemEntr.Value = True Then
                RecCred("TipoEntr") = "Sem entrada"
                RecCred("ValorEntr") = ""
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptChqEntrDin.Value = True Then
                RecCred("TipoEntr") = "Dinheiro"
                RecCred("ValorEntr") = CCur(TxtValorEntrDinChq.Text)
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptChqEntrChq.Value = True Then
                RecCred("TipoEntr") = "Cheque"
                RecCred("ValorEntr") = CCur(TxtValorEntrChequeChq.Text)
                
                If TxtBancoCar.Text <> "" Then
                    RecCred("NumBanco") = TxtBancoCar.Text
                Else
                    RecCred("NumBanco") = 0
                End If
                
                If TxtDigChq.Text <> "" Then
                    RecCred("NumCheque") = TxtChequeChq.Text & "-" & TxtDigChq.Text
                Else
                    RecCred("NumCheque") = TxtChequeChq.Text
                End If
            End If
            
            RecCred.Update
            
            RecCred.Close
            
            StrSql = "SELECT MAX(CodCred) FROM tb_Crediario where CodCli=" & VGIntCodCli
            RecCred.Open StrSql, vgCon, 1, 3
            
            VPIntCodCredTemp = RecCred.Fields.Item(0).Value
            
            '============== INCLUIR PARCELAS DO CREDIÁRIO =============================
            If CboPrazoChqParc.Text <> "" Then
                
                Do While parcelatemp <= Val(CboPrazoChqParc.Text)
                
                    StrSql = "SELECT * FROM tb_Crediario_Parcela"
                    RecCredParc.Open StrSql, vgCon, 1, 3
                     
                    RecCredParc.AddNew
                    RecCredParc("CodCred") = RecCred.Fields.Item(0).Value
                    RecCredParc("NumParc") = parcelatemp
                    
                    If parcelatemp = 1 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData01)
                        RecCredParc("Valor") = CCur(VGStrValor01)
                        RecCredParc("NumBanco") = VGStrBanco01
                        RecCredParc("NumCheque") = VGStrChequeDig01
                    
                    ElseIf parcelatemp = 2 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData02)
                        RecCredParc("Valor") = CCur(VGStrValor02)
                        RecCredParc("NumBanco") = VGStrBanco02
                        RecCredParc("NumCheque") = VGStrChequeDig02
                    
                    ElseIf parcelatemp = 3 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData03)
                        RecCredParc("Valor") = CCur(VGStrValor03)
                        RecCredParc("NumBanco") = VGStrBanco03
                        RecCredParc("NumCheque") = VGStrChequeDig03
                    
                    ElseIf parcelatemp = 4 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData04)
                        RecCredParc("Valor") = CCur(VGStrValor04)
                        RecCredParc("NumBanco") = VGStrBanco04
                        RecCredParc("NumCheque") = VGStrChequeDig04
                    
                    ElseIf parcelatemp = 5 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData05)
                        RecCredParc("Valor") = CCur(VGStrValor05)
                        RecCredParc("NumBanco") = VGStrBanco05
                        RecCredParc("NumCheque") = VGStrChequeDig05
                    
                    ElseIf parcelatemp = 6 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData06)
                        RecCredParc("Valor") = CCur(VGStrValor06)
                        RecCredParc("NumBanco") = VGStrBanco06
                        RecCredParc("NumCheque") = VGStrChequeDig06
                    
                    ElseIf parcelatemp = 7 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData07)
                        RecCredParc("Valor") = CCur(VGStrValor07)
                        RecCredParc("NumBanco") = VGStrBanco07
                        RecCredParc("NumCheque") = VGStrChequeDig07
                    
                    ElseIf parcelatemp = 8 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData08)
                        RecCredParc("Valor") = CCur(VGStrValor08)
                        RecCredParc("NumBanco") = VGStrBanco08
                        RecCredParc("NumCheque") = VGStrChequeDig08
                    
                    ElseIf parcelatemp = 9 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData09)
                        RecCredParc("Valor") = CCur(VGStrValor09)
                        RecCredParc("NumBanco") = VGStrBanco09
                        RecCredParc("NumCheque") = VGStrChequeDig09
                    
                    ElseIf parcelatemp = 10 Then
                        RecCredParc("Vencimento") = FormataDataUS(VGStrData10)
                        RecCredParc("Valor") = CCur(VGStrValor10)
                        RecCredParc("NumBanco") = VGStrBanco10
                        RecCredParc("NumCheque") = VGStrChequeDig10
                    End If
                    
                    RecCredParc("Quitado") = "não"
                    RecCredParc.Update
                         
                    RecCredParc.Close
                                    
                    parcelatemp = parcelatemp + 1
                Loop
                
            End If
            
            
            '============== INCLUIR VENDA =============================
            StrSql = "SELECT * FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
                
            RecVenda.AddNew
            RecVenda("CodVendedor") = Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10)
            RecVenda("CodCli") = VGIntCodCli
            RecVenda("CodCred") = RecCred.Fields.Item(0).Value
            RecVenda("DtVenda") = FormataDataUS(Date)
            
            If VPStrDescrProd01 <> "" Then
                RecVenda("CodProd01") = Mid(VPStrDescrProd01, 1, InStr(VPStrDescrProd01, "/") - 1)
            Else
                RecVenda("CodProd01") = 0
            End If
            
            If VPStrDescrProd02 <> "" Then
                RecVenda("CodProd02") = Mid(VPStrDescrProd02, 1, InStr(VPStrDescrProd02, "/") - 1)
            Else
                RecVenda("CodProd02") = 0
            End If
            
            If VPStrDescrProd03 <> "" Then
                RecVenda("CodProd03") = Mid(VPStrDescrProd03, 1, InStr(VPStrDescrProd03, "/") - 1)
            Else
                RecVenda("CodProd03") = 0
            End If
            
            If VPStrDescrProd04 <> "" Then
                RecVenda("CodProd04") = Mid(VPStrDescrProd04, 1, InStr(VPStrDescrProd04, "/") - 1)
            Else
                RecVenda("CodProd04") = 0
            End If
            
            If VPStrDescrProd05 <> "" Then
                RecVenda("CodProd05") = Mid(VPStrDescrProd05, 1, InStr(VPStrDescrProd05, "/") - 1)
            Else
                RecVenda("CodProd05") = 0
            End If
            
            If VPStrDescrProd06 <> "" Then
                RecVenda("CodProd06") = Mid(VPStrDescrProd06, 1, InStr(VPStrDescrProd06, "/") - 1)
            Else
                RecVenda("CodProd06") = 0
            End If
            
            RecVenda("CodForn01") = 0
            RecVenda("CodForn02") = 0
            RecVenda("CodForn03") = 0
            RecVenda("CodForn04") = 0
            RecVenda("CodForn05") = 0
            RecVenda("CodForn06") = 0
            RecVenda("Qtde01") = Val(VPStrQtdeProd01)
            RecVenda("Qtde02") = Val(VPStrQtdeProd02)
            RecVenda("Qtde03") = Val(VPStrQtdeProd03)
            RecVenda("Qtde04") = Val(VPStrQtdeProd04)
            RecVenda("Qtde05") = Val(VPStrQtdeProd05)
            RecVenda("Qtde06") = Val(VPStrQtdeProd06)
            If VPStrValorTotal01 <> "" Then
                RecVenda("ValorVenda01") = CCur(VPStrValorTotal01)
            Else
                RecVenda("ValorVenda01") = ""
            End If
            
            If VPStrValorTotal02 <> "" Then
                RecVenda("ValorVenda02") = CCur(VPStrValorTotal02)
            Else
                RecVenda("ValorVenda02") = ""
            End If
            
            If VPStrValorTotal03 <> "" Then
                RecVenda("ValorVenda03") = CCur(VPStrValorTotal03)
            Else
                RecVenda("ValorVenda03") = ""
            End If
            
            If VPStrValorTotal04 <> "" Then
                RecVenda("ValorVenda04") = CCur(VPStrValorTotal04)
            Else
                RecVenda("ValorVenda04") = ""
            End If
            
            If VPStrValorTotal05 <> "" Then
                RecVenda("ValorVenda05") = CCur(VPStrValorTotal05)
            Else
                RecVenda("ValorVenda05") = ""
            End If
            
            If VPStrValorTotal06 <> "" Then
                RecVenda("ValorVenda06") = CCur(VPStrValorTotal06)
            Else
                RecVenda("ValorVenda06") = ""
            End If
            RecVenda("TipoVenda") = "A prazo - Cheque"
            RecVenda("SubTotalVenda") = CCur(TxtVendaChq.Text)
            RecVenda("Desconto") = ""
            RecVenda("TotalVenda") = CCur(TxtTotalVendaChq.Text)
            RecVenda("TipoPagto") = "Cheque"
            RecVenda("NumBanco") = 0
            RecVenda("NumCheque") = ""
            RecVenda.Update
                
            RecVenda.Close
            RecCred.Close
            
            StrSql = "SELECT MAX(CodVenda) FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
            
            VLIntCodVendaTemp = RecVenda.Fields.Item(0).Value
            VGIntCodVendaRel = RecVenda.Fields.Item(0).Value
            
            RecVenda.Close
            
            '================ RETIRAR QTDE DO ESTOQUE ========================
            VLIntCountProd = 1
            VLIntCodTemp = 0
            
            Do While VLIntCountProd <= 6
                
                If VLIntCountProd = 1 Then
                    If VPStrDescrProd01 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd01, 1, InStr(VPStrDescrProd01, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd01)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 2 Then
                    If VPStrDescrProd02 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd02, 1, InStr(VPStrDescrProd02, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd02)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 3 Then
                    If VPStrDescrProd03 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd03, 1, InStr(VPStrDescrProd03, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd03)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 4 Then
                    If VPStrDescrProd04 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd04, 1, InStr(VPStrDescrProd04, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd04)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 5 Then
                    If VPStrDescrProd05 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd05, 1, InStr(VPStrDescrProd05, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd05)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 6 Then
                    If VPStrDescrProd06 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd06, 1, InStr(VPStrDescrProd06, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd06)
                        RecEst.Update
                        RecEst.Close
                    End If
                End If
                
                VLIntCodTemp = 0
                VLIntCountProd = VLIntCountProd + 1
            Loop
                
                
            '================= INCLUIR NO MOVIMENTO DE CAIXA =========================
            
            If OptChqEntrDin.Value = True Or OptChqEntrChq.Value = True Then
                StrSql = "SELECT * FROM tb_Caixa"
                RecCx.Open StrSql, vgCon, 1, 3
                
                RecCx.AddNew
                RecCx("CodVenda") = VLIntCodVendaTemp
                RecCx("DtMov") = FormataDataUS(Date)
                RecCx("TipoMov") = "Entrada de venda"
                
                If OptChqEntrDin.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrDinChq.Text)
                    
                ElseIf OptChqEntrChq.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrChequeChq.Text)
                End If
                
                RecCx("TipoValor") = "credito"
                RecCx("Descricao") = "Entrada de venda a prazo em cheque  - Cliente: " & VGStrNomeCli
                
                If OptChqEntrDin.Value = True Then
                    RecCx("TipoPagto") = "Dinheiro"
                    
                ElseIf OptChqEntrChq.Value = True Then
                    RecCx("TipoPagto") = "Cheque"
                End If
                
                RecCx.Update
            End If
            
            Desconecta
            
            VPStrResponse = MsgBox("Venda efetuada." & Chr(13) & Chr(13) & "Deseja imprimir a proposta de crédito agora?", vbYesNo, "Pró Ótica 2004 - Informação")
            
            Unload Me
            
            MDIPrincipal.Enabled = True
            MDIPrincipal.WindowState = 2
            
            Screen.MousePointer = vbNormal
            
            If VPStrResponse = vbYes Then
                Call MontaImpressaoProposta
            End If
            
        ElseIf VPStrVenda = "prazocarne" Then
            Dim VLStrData As String
            Dim VLStrValor As String
            
            '============== INCLUIR CREDIÁRIO =============================
            StrSql = "SELECT * FROM tb_Crediario"
            RecCred.Open StrSql, vgCon, 1, 3
            
            RecCred.AddNew
            RecCred("CodCredsta") = VGIntCodCredstaVenda
            RecCred("CodCli") = VGIntCodCli
            RecCred("DtCred") = FormataDataUS(Date)
            RecCred("TipoCred") = "Carnê"
            RecCred("ValorVenda") = CCur(TxtVendaCar.Text)
            RecCred("Parcela") = CboPrazoCarParc.Text
            RecCred("Juros") = TxtPrazoCarJuros.Text
            RecCred("ValorTotal") = CCur(TxtTotalVendaCar.Text)
            
            If OptCarSemEntr.Value = True Then
                RecCred("TipoEntr") = "Sem entrada"
                RecCred("ValorEntr") = ""
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptCarEntrDin.Value = True Then
                RecCred("TipoEntr") = "Dinheiro"
                RecCred("ValorEntr") = CCur(TxtValorEntrDinCar.Text)
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptCarEntrChq.Value = True Then
                RecCred("TipoEntr") = "Cheque"
                RecCred("ValorEntr") = CCur(TxtValorEntrChequeCar.Text)
                
                If TxtBancoCar.Text <> "" Then
                    RecCred("NumBanco") = TxtBancoCar.Text
                Else
                    RecCred("NumBanco") = 0
                End If
                
                If TxtDigCar.Text <> "" Then
                    RecCred("NumCheque") = TxtChequeCar.Text & "-" & TxtDigCar.Text
                Else
                    RecCred("NumCheque") = TxtChequeCar.Text
                End If
            End If
            
            RecCred.Update
            
            RecCred.Close
            
            StrSql = "SELECT MAX(CodCred) FROM tb_Crediario where CodCli=" & VGIntCodCli
            RecCred.Open StrSql, vgCon, 1, 3
            
            VPIntCodCredTemp = RecCred.Fields.Item(0).Value
            
            '============== INCLUIR PARCELAS DO CREDIÁRIO =============================
            If CboPrazoCarParc.Text <> "" Then
                
                VLStrData = Date
                VLStrValor = CCur(Mid(LblParcCar.Caption, InStr(LblParcCar.Caption, "R$")))
                
                Do While parcelatemp <= Val(CboPrazoCarParc.Text)
                
                    StrSql = "SELECT * FROM tb_Crediario_Parcela"
                    RecCredParc.Open StrSql, vgCon, 1, 3
                     
                    RecCredParc.AddNew
                    RecCredParc("CodCred") = RecCred.Fields.Item(0).Value
                    RecCredParc("NumParc") = parcelatemp
                    
                    VLStrData = DateSerial(Year(VLStrData), Month(VLStrData), Day(VLStrData) + 30)
                    
                    RecCredParc("Vencimento") = FormataDataUS(VLStrData)
                    RecCredParc("Valor") = VLStrValor
                    RecCredParc("Quitado") = "não"
                    RecCredParc("NumBanco") = 0
                    RecCredParc("NumCheque") = ""
                    RecCredParc.Update
                         
                    RecCredParc.Close
                                    
                    parcelatemp = parcelatemp + 1
                Loop
                
            End If
            
            
            '============== INCLUIR VENDA =============================
            StrSql = "SELECT * FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
                
            RecVenda.AddNew
            If CboVendedor.Text = "" Then
                RecVenda("CodVendedor") = "0"
            Else
                RecVenda("CodVendedor") = Trim(Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10))
            End If
            RecVenda("CodCli") = VGIntCodCli
            RecVenda("CodCred") = RecCred.Fields.Item(0).Value
            RecVenda("DtVenda") = FormataDataUS(Date)
            
            If VPStrDescrProd01 <> "" Then
                RecVenda("CodProd01") = Mid(VPStrDescrProd01, 1, InStr(VPStrDescrProd01, "/") - 1)
            Else
                RecVenda("CodProd01") = 0
            End If
            
            If VPStrDescrProd02 <> "" Then
                RecVenda("CodProd02") = Mid(VPStrDescrProd02, 1, InStr(VPStrDescrProd02, "/") - 1)
            Else
                RecVenda("CodProd02") = 0
            End If
            
            If VPStrDescrProd03 <> "" Then
                RecVenda("CodProd03") = Mid(VPStrDescrProd03, 1, InStr(VPStrDescrProd03, "/") - 1)
            Else
                RecVenda("CodProd03") = 0
            End If
            
            If VPStrDescrProd04 <> "" Then
                RecVenda("CodProd04") = Mid(VPStrDescrProd04, 1, InStr(VPStrDescrProd04, "/") - 1)
            Else
                RecVenda("CodProd04") = 0
            End If
            
            If VPStrDescrProd05 <> "" Then
                RecVenda("CodProd05") = Mid(VPStrDescrProd05, 1, InStr(VPStrDescrProd05, "/") - 1)
            Else
                RecVenda("CodProd05") = 0
            End If
            
            If VPStrDescrProd06 <> "" Then
                RecVenda("CodProd06") = Mid(VPStrDescrProd06, 1, InStr(VPStrDescrProd06, "/") - 1)
            Else
                RecVenda("CodProd06") = 0
            End If
            
            RecVenda("CodForn01") = 0
            RecVenda("CodForn02") = 0
            RecVenda("CodForn03") = 0
            RecVenda("CodForn04") = 0
            RecVenda("CodForn05") = 0
            RecVenda("CodForn06") = 0
            RecVenda("Qtde01") = Val(VPStrQtdeProd01)
            RecVenda("Qtde02") = Val(VPStrQtdeProd02)
            RecVenda("Qtde03") = Val(VPStrQtdeProd03)
            RecVenda("Qtde04") = Val(VPStrQtdeProd04)
            RecVenda("Qtde05") = Val(VPStrQtdeProd05)
            RecVenda("Qtde06") = Val(VPStrQtdeProd06)
            If VPStrValorTotal01 <> "" Then
                RecVenda("ValorVenda01") = CCur(VPStrValorTotal01)
            Else
                RecVenda("ValorVenda01") = ""
            End If
            
            If VPStrValorTotal02 <> "" Then
                RecVenda("ValorVenda02") = CCur(VPStrValorTotal02)
            Else
                RecVenda("ValorVenda02") = ""
            End If
            
            If VPStrValorTotal03 <> "" Then
                RecVenda("ValorVenda03") = CCur(VPStrValorTotal03)
            Else
                RecVenda("ValorVenda03") = ""
            End If
            
            If VPStrValorTotal04 <> "" Then
                RecVenda("ValorVenda04") = CCur(VPStrValorTotal04)
            Else
                RecVenda("ValorVenda04") = ""
            End If
            
            If VPStrValorTotal05 <> "" Then
                RecVenda("ValorVenda05") = CCur(VPStrValorTotal05)
            Else
                RecVenda("ValorVenda05") = ""
            End If
            
            If VPStrValorTotal06 <> "" Then
                RecVenda("ValorVenda06") = CCur(VPStrValorTotal06)
            Else
                RecVenda("ValorVenda06") = ""
            End If
            RecVenda("TipoVenda") = "A prazo - Carnê"
            RecVenda("SubTotalVenda") = CCur(TxtVendaCar.Text)
            RecVenda("Desconto") = ""
            RecVenda("TotalVenda") = CCur(TxtTotalVendaCar.Text)
            RecVenda("TipoPagto") = "Carnê"
            RecVenda("NumBanco") = 0
            RecVenda("NumCheque") = ""
            RecVenda.Update
                
            RecVenda.Close
            RecCred.Close
            
            StrSql = "SELECT MAX(CodVenda) FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
            
            VLIntCodVendaTemp = RecVenda.Fields.Item(0).Value
            VGIntCodVendaRel = RecVenda.Fields.Item(0).Value
            
            RecVenda.Close
            
            '================ RETIRAR QTDE DO ESTOQUE ========================
            VLIntCountProd = 1
            VLIntCodTemp = 0
            
            Do While VLIntCountProd <= 6
                
                If VLIntCountProd = 1 Then
                    If VPStrDescrProd01 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd01, 1, InStr(VPStrDescrProd01, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd01)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 2 Then
                    If VPStrDescrProd02 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd02, 1, InStr(VPStrDescrProd02, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd02)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 3 Then
                    If VPStrDescrProd03 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd03, 1, InStr(VPStrDescrProd03, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd03)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 4 Then
                    If VPStrDescrProd04 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd04, 1, InStr(VPStrDescrProd04, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd04)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 5 Then
                    If VPStrDescrProd05 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd05, 1, InStr(VPStrDescrProd05, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd05)
                        RecEst.Update
                        RecEst.Close
                    End If
                
                ElseIf VLIntCountProd = 6 Then
                    If VPStrDescrProd06 <> "" Then
                        VLIntCodTemp = Mid(VPStrDescrProd06, 1, InStr(VPStrDescrProd06, "/") - 1)
                        
                        StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & VLIntCodTemp
                        RecEst.Open StrSql, vgCon, 1, 3
                        
                        RecEst("QtdeProd") = Int(RecEst.Fields.Item(0).Value) - Int(VPStrQtdeProd06)
                        RecEst.Update
                        RecEst.Close
                    End If
                End If
                
                VLIntCodTemp = 0
                VLIntCountProd = VLIntCountProd + 1
            Loop
                
                
            '================= INCLUIR NO MOVIMENTO DE CAIXA =========================
            
            If OptCarEntrDin.Value = True Or OptCarEntrChq.Value = True Then
                StrSql = "SELECT * FROM tb_Caixa"
                RecCx.Open StrSql, vgCon, 1, 3
                
                RecCx.AddNew
                RecCx("CodVenda") = VLIntCodVendaTemp
                RecCx("DtMov") = FormataDataUS(Date)
                RecCx("TipoMov") = "Entrada de venda"
                
                If OptCarEntrDin.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrDinCar.Text)
                    
                ElseIf OptCarEntrChq.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrChequeCar.Text)
                End If
                
                RecCx("TipoValor") = "credito"
                RecCx("Descricao") = "Entrada de venda a prazo em carnê  - Cliente: " & VGStrNomeCli
                
                If OptCarEntrDin.Value = True Then
                    RecCx("TipoPagto") = "Dinheiro"
                    
                ElseIf OptCarEntrChq.Value = True Then
                    RecCx("TipoPagto") = "Cheque"
                End If
                
                RecCx.Update
            End If
            
            Desconecta
            
            VPStrResponse = MsgBox("Venda efetuada." & Chr(13) & Chr(13) & "Deseja imprimir a proposta de crédito agora?", vbYesNo, "Pró Ótica 2004 - Informação")
            
            Unload Me
            
            MDIPrincipal.Enabled = True
            MDIPrincipal.WindowState = 2
            
            Screen.MousePointer = vbNormal
            
            If VPStrResponse = vbYes Then
                Call MontaImpressaoProposta
            End If
        
        End If
        
    End If
   
End Sub

Private Sub CmdParcInc_Click()
    If CboPrazoChqParc.Text = "" Then
        VPStrBox = MsgBox("Você deve selecionar a quantidade de parcelas.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        FrmParcela_Inc.Show
    End If
End Sub

Private Sub CmdProx_Click()
    VPIntProd = VPIntProd + 1
            
    If VPStrValorTotal01 = "0" Then
        VPStrValorTotal01 = ""
    End If
    
    If VPStrValorTotal02 = "0" Then
        VPStrValorTotal02 = ""
    End If
    
    If VPStrValorTotal03 = "0" Then
        VPStrValorTotal03 = ""
    End If
    
    If VPStrValorTotal04 = "0" Then
        VPStrValorTotal04 = ""
    End If
    
    If VPStrValorTotal05 = "0" Then
        VPStrValorTotal05 = ""
    End If
    
    If VPStrValorTotal06 = "0" Then
        VPStrValorTotal06 = ""
    End If
    
    FraProduto.Caption = "Produto " & FormataNum(VPIntProd)
    
    If VPIntProd = 2 Then
        CmdAnt.Enabled = True

        VPStrTipoProd01 = LblTipoProd.Caption
        VPStrDescrProd01 = TxtDescrProd.Text
        VPStrPrecoUnit01 = TxtPrecoUnit.Text
        VPStrQtdeProd01 = TxtQtdeProd.Text
        VPStrValorTotal01 = TxtValorTotal.Text
        VPStrQtdeProdEst01 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd02
        TxtDescrProd.Text = VPStrDescrProd02
        TxtPrecoUnit.Text = VPStrPrecoUnit02
        TxtQtdeProd.Text = VPStrQtdeProd02
        TxtValorTotal.Text = VPStrValorTotal02
        LblQtdeProdEst.Caption = VPStrQtdeProdEst02
    
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 3 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd02 = LblTipoProd.Caption
        VPStrDescrProd02 = TxtDescrProd.Text
        VPStrPrecoUnit02 = TxtPrecoUnit.Text
        VPStrQtdeProd02 = TxtQtdeProd.Text
        VPStrValorTotal02 = TxtValorTotal.Text
        VPStrQtdeProdEst02 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd03
        TxtDescrProd.Text = VPStrDescrProd03
        TxtPrecoUnit.Text = VPStrPrecoUnit03
        TxtQtdeProd.Text = VPStrQtdeProd03
        TxtValorTotal.Text = VPStrValorTotal03
        LblQtdeProdEst.Caption = VPStrQtdeProdEst03
    
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 4 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd03 = LblTipoProd.Caption
        VPStrDescrProd03 = TxtDescrProd.Text
        VPStrPrecoUnit03 = TxtPrecoUnit.Text
        VPStrQtdeProd03 = TxtQtdeProd.Text
        VPStrValorTotal03 = TxtValorTotal.Text
        VPStrQtdeProdEst03 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd04
        TxtDescrProd.Text = VPStrDescrProd04
        TxtPrecoUnit.Text = VPStrPrecoUnit04
        TxtQtdeProd.Text = VPStrQtdeProd04
        TxtValorTotal.Text = VPStrValorTotal04
        LblQtdeProdEst.Caption = VPStrQtdeProdEst04
    
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 5 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd04 = LblTipoProd.Caption
        VPStrDescrProd04 = TxtDescrProd.Text
        VPStrPrecoUnit04 = TxtPrecoUnit.Text
        VPStrQtdeProd04 = TxtQtdeProd.Text
        VPStrValorTotal04 = TxtValorTotal.Text
        VPStrQtdeProdEst04 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd05
        TxtDescrProd.Text = VPStrDescrProd05
        TxtPrecoUnit.Text = VPStrPrecoUnit05
        TxtQtdeProd.Text = VPStrQtdeProd05
        TxtValorTotal.Text = VPStrValorTotal05
        LblQtdeProdEst.Caption = VPStrQtdeProdEst05
    
        CmdProx.Enabled = True
        
    ElseIf VPIntProd = 6 Then
        CmdAnt.Enabled = True
        
        VPStrTipoProd05 = LblTipoProd.Caption
        VPStrDescrProd05 = TxtDescrProd.Text
        VPStrPrecoUnit05 = TxtPrecoUnit.Text
        VPStrQtdeProd05 = TxtQtdeProd.Text
        VPStrValorTotal05 = TxtValorTotal.Text
        VPStrQtdeProdEst05 = LblQtdeProdEst.Caption
        
        LblTipoProd.Caption = VPStrTipoProd06
        TxtDescrProd.Text = VPStrDescrProd06
        TxtPrecoUnit.Text = VPStrPrecoUnit06
        TxtQtdeProd.Text = VPStrQtdeProd06
        TxtValorTotal.Text = VPStrValorTotal06
        LblQtdeProdEst.Caption = VPStrQtdeProdEst06
    
        CmdProx.Enabled = False
    End If
    
End Sub

Private Sub CmdVerProd_Click()
    FrmVenda_Inc_Prod.Show
End Sub

Private Sub Form_Resize()
  FrmVenda_Inc.Left = (MDIPrincipal.Width / 2) - (FrmVenda_Inc.Width / 2)
  FrmVenda_Inc.Top = (MDIPrincipal.Height / 3) - (FrmVenda_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 8070
    Width = 8040
    Top = 135
    Left = 3180
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    LblTipoProd.Caption = ""
    LblQtdeProdEst.Visible = False
    CmdAnt.Enabled = False
    VPIntProd = 1
    
    Call MontaCboVendedor
    Call MontaParcelas
    
    If VGStrVendaRapida = "sim" Then
        VGStrVendaRapida = ""
        OptPrazo.Enabled = False
    Else
        OptPrazo.Enabled = True
    End If
    
End Sub

Private Sub OptCarEntrChq_Click()
    FraCarEntrDin.Visible = False
    FraCarEntrChq.Visible = True

    Dim restparc As String
    
    If TxtTotalVendaCar.Text = "" Then
        TxtTotalVendaCar.Text = TxtVendaCar.Text
    End If
    
    If TxtTotalVendaCar.Text <> "" And CboPrazoCarParc.Text <> "" Then
        restparc = CCur(TxtTotalVendaCar.Text) - ((CCur(TxtTotalVendaCar.Text) * 20) / 100)
        
        LblEntrCar.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaCar.Text) * 20) / 100))
        
        If CboPrazoCarParc.Text = "00" And restparc = "0" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoCarParc.Text))
        End If
    End If
End Sub

Private Sub OptCarEntrDin_Click()
    FraCarEntrDin.Visible = True
    FraCarEntrChq.Visible = False

    Dim restparc As String
    
    If TxtTotalVendaCar.Text = "" Then
        TxtTotalVendaCar.Text = TxtVendaCar.Text
    End If
    
    If TxtTotalVendaCar.Text <> "" And CboPrazoCarParc.Text <> "" Then
        restparc = CCur(TxtTotalVendaCar.Text) - ((CCur(TxtTotalVendaCar.Text) * 20) / 100)

        LblEntrCar.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaCar.Text) * 20) / 100))
        
        If CboPrazoCarParc = "00" And restparc = "0" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoCarParc.Text))
        End If
    End If
End Sub

Private Sub OptCarSemEntr_Click()
    FraCarEntrDin.Visible = False
    FraCarEntrChq.Visible = False

    If TxtTotalVendaCar.Text <> "" And CboPrazoCarParc.Text <> "" Then
        LblEntrCar.Caption = "Entrada: R$ 0,00"
        
        If TxtTotalVendaCar.Text = "R$ 0,00" And CboPrazoCarParc.Text = "00" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(CCur(TxtTotalVendaCar.Text / CboPrazoCarParc.Text))
        End If
    
    End If
End Sub

Private Sub OptChqEntrChq_Click()
    FraChqEntrChq.Visible = True
    FraChqEntrDin.Visible = False

    Dim restparc As String
    
    If TxtTotalVendaChq.Text = "" Then
        TxtTotalVendaChq.Text = TxtVendaChq.Text
    End If
    
    If TxtTotalVendaChq.Text <> "" And CboPrazoChqParc.Text <> "" Then
        restparc = CCur(TxtTotalVendaChq.Text) - ((CCur(TxtTotalVendaChq.Text) * 20) / 100)
        
        LblEntrChq.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaChq.Text) * 20) / 100))
        
        If restparc = "0" And CboPrazoChqParc.Text = "00" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoChqParc.Text))
        End If
    End If
End Sub

Private Sub OptChqEntrDin_Click()
    FraChqEntrChq.Visible = False
    FraChqEntrDin.Visible = True

    Dim restparc As String
    
    If TxtTotalVendaChq.Text = "" Then
        TxtTotalVendaChq.Text = TxtVendaChq.Text
    End If
    
    If TxtTotalVendaChq.Text <> "" And CboPrazoChqParc.Text <> "" Then
        restparc = CCur(TxtTotalVendaChq.Text) - ((CCur(TxtTotalVendaChq.Text) * 20) / 100)
        
        LblEntrChq.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaChq.Text) * 20) / 100))
        
        If restparc = "0" And CboPrazoChqParc.Text = "00" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoChqParc.Text))
        End If
    End If
End Sub

Private Sub OptChqSemEntr_Click()
    FraChqEntrChq.Visible = False
    FraChqEntrDin.Visible = False
    
    If TxtTotalVendaChq.Text <> "" And CboPrazoChqParc.Text <> "" Then
        LblEntrChq.Caption = "Entrada: R$ 0,00"
        
        If TxtTotalVendaChq.Text = "R$ 0,00" And CboPrazoChqParc.Text = "00" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(CCur(TxtTotalVendaChq.Text / CboPrazoChqParc.Text))
        End If
    End If
End Sub

Private Sub OptCredCarne_Click()
    FraCredCarne.Visible = True
    FraCredCheque.Visible = False
End Sub

Private Sub OptCredCheque_Click()
    FraCredCarne.Visible = False
    FraCredCheque.Visible = True
End Sub

Private Sub OptPrazo_Click()
    OptPrazoCheque.Value = False
    OptPrazoCarne.Value = False
    
    FraVista.Visible = False
    FraPrazoCarne.Visible = False
    FraPrazoCheque.Visible = False
    FraTipoPrazo.Visible = True
End Sub

Private Sub OptPrazoCarne_Click()
    VPStrVenda = "prazocarne"
    FraVista.Visible = False
    FraPrazoCarne.Visible = True
    FraPrazoCheque.Visible = False
    FraTipoPrazo.Visible = True
    
    If VPIntProd = 1 Then
        VPStrValorTotal01 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        
        VPStrDescrProd01 = TxtDescrProd.Text
        VPStrQtdeProd01 = TxtQtdeProd.Text
    End If
    
    If VPIntProd = 2 Then
        VPStrValorTotal02 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        
        VPStrDescrProd02 = TxtDescrProd.Text
        VPStrQtdeProd02 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 3 Then
        VPStrValorTotal03 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd03 = TxtDescrProd.Text
        VPStrQtdeProd03 = TxtQtdeProd.Text
        
    End If
    
    If VPIntProd = 4 Then
        VPStrValorTotal04 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd04 = TxtDescrProd.Text
        VPStrQtdeProd04 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 5 Then
        VPStrValorTotal05 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd05 = TxtDescrProd.Text
        VPStrQtdeProd05 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 6 Then
        VPStrValorTotal06 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd06 = TxtDescrProd.Text
        VPStrQtdeProd06 = TxtQtdeProd.Text
    
    End If
    
    If OptVista.Value = True Then
        TxtVendaVista.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVista.SetFocus
        OptDin.Value = True
        
    ElseIf OptPrazoCheque.Value = True Then
        TxtVendaChq.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVendaChq.SetFocus
        OptChqEntrDin.Value = True
        
    ElseIf OptPrazoCarne.Value = True Then
        TxtVendaCar.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVendaCar.SetFocus
        OptCarEntrDin.Value = True
    
    End If
    
End Sub

Private Sub OptPrazoCheque_Click()
    VPStrVenda = "prazocheque"
    FraVista.Visible = False
    FraPrazoCarne.Visible = False
    FraPrazoCheque.Visible = True
    FraTipoPrazo.Visible = True
    
    If VPIntProd = 1 Then
        VPStrValorTotal01 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        
        VPStrDescrProd01 = TxtDescrProd.Text
        VPStrQtdeProd01 = TxtQtdeProd.Text
    End If
    
    If VPIntProd = 2 Then
        VPStrValorTotal02 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        
        VPStrDescrProd02 = TxtDescrProd.Text
        VPStrQtdeProd02 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 3 Then
        VPStrValorTotal03 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd03 = TxtDescrProd.Text
        VPStrQtdeProd03 = TxtQtdeProd.Text
        
    End If
    
    If VPIntProd = 4 Then
        VPStrValorTotal04 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd04 = TxtDescrProd.Text
        VPStrQtdeProd04 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 5 Then
        VPStrValorTotal05 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd05 = TxtDescrProd.Text
        VPStrQtdeProd05 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 6 Then
        VPStrValorTotal06 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd06 = TxtDescrProd.Text
        VPStrQtdeProd06 = TxtQtdeProd.Text
    
    End If
    
    If OptVista.Value = True Then
        TxtVendaVista.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVista.SetFocus
        OptDin.Value = True
        
    ElseIf OptPrazoCheque.Value = True Then
        TxtVendaChq.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVendaChq.SetFocus
        OptChqEntrDin.Value = True
        
    ElseIf OptPrazoCarne.Value = True Then
        TxtVendaCar.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVendaCar.SetFocus
        OptCarEntrDin.Value = True
    
    End If
End Sub

Private Sub OptVista_Click()
    VPStrVenda = "vista"
    FraVista.Visible = True
    FraPrazoCarne.Visible = False
    FraPrazoCheque.Visible = False
    FraTipoPrazo.Visible = False
    
    If VPIntProd = 1 Then
        VPStrValorTotal01 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        
        VPStrDescrProd01 = TxtDescrProd.Text
        VPStrQtdeProd01 = TxtQtdeProd.Text
    End If
    
    If VPIntProd = 2 Then
        VPStrValorTotal02 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        
        VPStrDescrProd02 = TxtDescrProd.Text
        VPStrQtdeProd02 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 3 Then
        VPStrValorTotal03 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd03 = TxtDescrProd.Text
        VPStrQtdeProd03 = TxtQtdeProd.Text
        
    End If
    
    If VPIntProd = 4 Then
        VPStrValorTotal04 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd04 = TxtDescrProd.Text
        VPStrQtdeProd04 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 5 Then
        VPStrValorTotal05 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd05 = TxtDescrProd.Text
        VPStrQtdeProd05 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 6 Then
        VPStrValorTotal06 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd06 = TxtDescrProd.Text
        VPStrQtdeProd06 = TxtQtdeProd.Text
    
    End If
    
    If OptVista.Value = True Then
        TxtVendaVista.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVista.SetFocus
        OptDin.Value = True
        
    ElseIf OptPrazoCheque.Value = True Then
        TxtVendaChq.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVendaChq.SetFocus
        OptChqEntrDin.Value = True
        
    ElseIf OptPrazoCarne.Value = True Then
        TxtVendaCar.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
        TxtTotalVendaCar.SetFocus
        OptCarEntrDin.Value = True
    
    End If
    
End Sub

Private Sub TxtDescrProd_LostFocus()
    Dim RecEst As New ADODB.Recordset
    
    If TxtDescrProd.Text <> "" Then
        Conecta
        
        StrSql = "Select QtdeProd,PrecoVenda from tb_Estoque where CodProd=" & VGIntCodProd
        RecEst.Open StrSql, vgCon, 1, 3
        
        If RecEst.EOF Then
            Desconecta
            VPStrBox = MsgBox("Este produto ainda não tem informações de estoque.", vbInformation, "Pró Ótica 2004 - Informação")
            TxtPrecoUnit.Text = ""
            TxtQtdeProd.Text = ""
            TxtValorTotal.Text = ""
            LblQtdeProdEstTemp.Caption = ""
            LblQtdeProdEst.Caption = ""
            LblQtdeProdEst.Visible = False
        Else
            TxtPrecoUnit.Text = FormataMoeda(RecEst.Fields.Item(1).Value)
            LblQtdeProdEstTemp.Caption = RecEst.Fields.Item(0).Value
            LblQtdeProdEst.Caption = "Em estoque: " & FormataNum(LblQtdeProdEstTemp.Caption)
            LblQtdeProdEst.Visible = True
            
            Desconecta
            
            TxtPrecoUnit.SetFocus
        End If
    End If
End Sub

Private Sub TxtPrazoCarJuros_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrazoCarJuros_LostFocus()
    If TxtPrazoCarJuros.Text <> "" And TxtVendaCar.Text <> "" Then
        TxtTotalVendaCar.Text = FormataMoeda(CCur(TxtVendaCar.Text) + (CCur(TxtVendaCar.Text) * TxtPrazoCarJuros.Text) / 100)
    
    ElseIf TxtPrazoCarJuros.Text = "" And TxtVendaCar.Text <> "" Then
        TxtTotalVendaCar.Text = FormataMoeda(TxtVendaCar.Text)
    
    End If
End Sub

Private Sub TxtPrazoChqJuros_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrazoChqJuros_LostFocus()
    If TxtPrazoChqJuros.Text <> "" And TxtVendaChq.Text <> "" Then
        TxtTotalVendaChq.Text = FormataMoeda(CCur(TxtVendaChq.Text) + (CCur(TxtVendaChq.Text) * TxtPrazoChqJuros.Text) / 100)
    
    ElseIf TxtPrazoChqJuros.Text = "" And TxtVendaChq.Text <> "" Then
        TxtTotalVendaChq.Text = FormataMoeda(TxtVendaChq.Text)
    
    End If
End Sub

Private Sub TxtPrecoUnit_LostFocus()
    If TxtPrecoUnit.Text <> "" And TxtQtdeProd.Text <> "" Then
        TxtValorTotal.Text = FormataMoeda(CCur(TxtPrecoUnit.Text) * Int(TxtQtdeProd.Text))
    End If
End Sub

Private Sub TxtQtdeProd_LostFocus()
    If TxtQtdeProd.Text <> "" And TxtPrecoUnit.Text <> "" Then
        TxtValorTotal.Text = FormataMoeda(CCur(TxtPrecoUnit.Text) * Int(TxtQtdeProd.Text))
        
        If (Int(LblQtdeProdEstTemp.Caption) - Int(TxtQtdeProd.Text)) < 0 Then
            VPStrBox = MsgBox("Estoque não possui a quantidade informada.", vbInformation, "Pró Ótica 2004 - Informação")
            TxtQtdeProd.SetFocus
        Else
           LblQtdeProdEst.Caption = "Em estoque: " & FormataNum(Int(LblQtdeProdEstTemp.Caption) - Int(TxtQtdeProd.Text))
           LblQtdeProdEst.Visible = True
        End If
    End If
End Sub

Sub MontaCboVendedor()
    Conecta
    
    Dim RecVend As New ADODB.Recordset
    
    StrSql = "SELECT CodVendedor,Nome FROM tb_Vendedor order by Nome"
    RecVend.Open StrSql, vgCon, 1, 3
    
    CboVendedor.AddItem ("                                                                                                                 0")
    Do While Not RecVend.EOF
        CboVendedor.AddItem (RecVend.Fields.Item(1).Value & "                                                                                                      " & RecVend.Fields.Item(0).Value)
        RecVend.MoveNext
    Loop
    
    Desconecta
    
End Sub

Private Sub TxtTotalVista_GotFocus()
    If TxtVendaVista.Text <> "" And TxtDescVista.Text <> "" Then
        TxtTotalVista.Text = FormataMoeda(CCur(TxtVendaVista.Text) - ((CCur(TxtVendaVista.Text) * TxtDescVista.Text) / 100))
    
    ElseIf TxtVendaVista.Text <> "" And TxtDescVista.Text = "" Then
        TxtTotalVista.Text = FormataMoeda(TxtVendaVista.Text)
    
    End If
End Sub

Private Sub TxtValorEntrChequeCar_LostFocus()
    If TxtValorEntrChequeCar.Text <> "" Then
        TxtValorEntrChequeCar.Text = FormataMoeda(TxtValorEntrChequeCar.Text)
    End If
End Sub

Private Sub TxtValorEntrChequeChq_LostFocus()
    If TxtValorEntrChequeChq.Text <> "" Then
        TxtValorEntrChequeChq.Text = FormataMoeda(TxtValorEntrChequeChq.Text)
    End If
End Sub

Private Sub TxtValorEntrDinCar_LostFocus()
    If TxtValorEntrDinCar.Text <> "" Then
        TxtValorEntrDinCar.Text = FormataMoeda(TxtValorEntrDinCar.Text)
    End If
End Sub

Private Sub TxtValorEntrDinChq_LostFocus()
    If TxtValorEntrDinChq.Text <> "" Then
        TxtValorEntrDinChq.Text = FormataMoeda(TxtValorEntrDinChq.Text)
    End If
End Sub

Private Sub TxtValorTotal_LostFocus()
    
    If VPIntProd = 1 Then
        VPStrValorTotal01 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd01 = TxtDescrProd.Text
        VPStrQtdeProd01 = TxtQtdeProd.Text
        
    End If
    
    If VPIntProd = 2 Then
        VPStrValorTotal02 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd02 = TxtDescrProd.Text
        VPStrQtdeProd02 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 3 Then
        VPStrValorTotal03 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd03 = TxtDescrProd.Text
        VPStrQtdeProd03 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 4 Then
        VPStrValorTotal04 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd04 = TxtDescrProd.Text
        VPStrQtdeProd04 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 5 Then
        VPStrValorTotal05 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd05 = TxtDescrProd.Text
        VPStrQtdeProd05 = TxtQtdeProd.Text
    
    End If
    
    If VPIntProd = 6 Then
        VPStrValorTotal06 = TxtValorTotal.Text
        If VPStrValorTotal01 = "" Then
            VPStrValorTotal01 = 0
        End If
        
        If VPStrValorTotal02 = "" Then
            VPStrValorTotal02 = 0
        End If
        
        If VPStrValorTotal03 = "" Then
            VPStrValorTotal03 = 0
        End If
        
        If VPStrValorTotal04 = "" Then
            VPStrValorTotal04 = 0
        End If
        
        If VPStrValorTotal05 = "" Then
            VPStrValorTotal05 = 0
        End If
        
        If VPStrValorTotal06 = "" Then
            VPStrValorTotal06 = 0
        End If
        VPStrDescrProd06 = TxtDescrProd.Text
        VPStrQtdeProd06 = TxtQtdeProd.Text
    
    End If
    
    If OptVista.Value = True Then
        TxtVendaVista.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
    
    ElseIf OptPrazoCheque.Value = True Then
        TxtVendaChq.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
    
    ElseIf OptPrazoCarne.Value = True Then
        TxtVendaCar.Text = FormataMoeda(CCur(VPStrValorTotal01) + CCur(VPStrValorTotal02) + CCur(VPStrValorTotal03) + CCur(VPStrValorTotal04) + CCur(VPStrValorTotal05) + CCur(VPStrValorTotal06))
    
    End If
End Sub

Private Sub TxtVendaCar_LostFocus()
    If TxtPrazoCarJuros.Text = "" Then
        TxtTotalVendaCar.Text = FormataMoeda(TxtVendaCar.Text)
    Else
        TxtTotalVendaCar.Text = FormataMoeda(CCur(TxtVendaCar.Text) + (CCur(TxtVendaCar.Text) * TxtPrazoCarJuros.Text) / 100)
    End If
End Sub

Private Sub TxtVendaChq_LostFocus()
    If TxtPrazoChqJuros.Text = "" Then
        TxtTotalVendaChq.Text = FormataMoeda(TxtVendaChq.Text)
    Else
        TxtTotalVendaChq.Text = FormataMoeda(CCur(TxtVendaChq.Text) + (CCur(TxtVendaChq.Text) * TxtPrazoChqJuros.Text) / 100)
    End If
End Sub

Sub MontaImpressaoProposta()
    Screen.MousePointer = vbHourglass
    
    Dim RecCred As New ADODB.Recordset
    Dim RecCredParc As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    Dim RecCredsta As New ADODB.Recordset
    Dim RecMed As New ADODB.Recordset
    Dim RecRec As New ADODB.Recordset
    Dim RecAux As New ADODB.Recordset
    Dim VLStrNomeMed As String
    Dim VLStrCRMMed As String
    Dim VLStrCPFMed As String
    Dim parctemp As Integer
    
    Conecta
    
    If VGStrProposta = "imprimir" Then
        VGStrProposta = ""
        VPIntCodCredTemp = VGIntCodCredTemp
    End If
    
    '=== Pega informações do crediário =======
    StrSql = "Select CodCredsta,CodCli,DtCred,TipoCred,ValorVenda,Parcela,Juros,ValorTotal,TipoEntr,ValorEntr " & _
             "From tb_Crediario where CodCred=" & VPIntCodCredTemp
    RecCred.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações das parcelas crediário =======
    StrSql = "Select Vencimento,Valor From tb_Crediario_Parcela where CodCred=" & VPIntCodCredTemp
    RecCredParc.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do crediarista =======
    StrSql = "Select Nome,Endereco,Bairro,Cep,Cidade,Estado,DtNasc,Telefone,CPF " & _
             "From tb_Crediarista where CodCredsta=" & RecCred.Fields.Item(0).Value
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do cliente =======
    StrSql = "Select Nome,Endereco,Bairro,Cep,Cidade,Estado,DtNasc,Telefone,CPF " & _
             "From tb_Cliente where CodCli=" & RecCred.Fields.Item(1).Value
    RecCli.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do médico =======
    StrSql = "Select CodMed From tb_Receita where CodCli=" & RecCred.Fields.Item(1).Value
    RecRec.Open StrSql, vgCon, 1, 3

    If Not RecRec.EOF Then
        StrSql = "Select Nome,CRM,Cpf From tb_Medico where CodMed=" & RecRec.Fields.Item(0).Value
        RecMed.Open StrSql, vgCon, 1, 3
        VLStrNomeMed = RecMed!nome
        VLStrCRMMed = RecMed!crm
        VLStrCPFMed = RecMed!cpf
    Else
        VLStrNomeMed = ""
        VLStrCRMMed = ""
        VLStrCPFMed = ""
    End If
    
    '=== Insere informações na tabela auxiliar =======
    StrSql = "Select * From tb_Auxiliar"
    RecAux.Open StrSql, vgCon, 1, 3
    
    RecAux.AddNew
    RecAux("Campo01") = RecCredsta!nome
    RecAux("Campo02") = FormataData(RecCredsta!dtnasc)
    RecAux("Campo03") = RecCredsta!cpf
    RecAux("Campo04") = RecCredsta!telefone
    RecAux("Campo05") = RecCredsta!endereco
    RecAux("Campo06") = RecCredsta!bairro
    RecAux("Campo07") = RecCredsta!cidade & "/" & RecCredsta!Estado
    RecAux("Campo08") = RecCredsta!cep
    RecAux("Campo09") = RecCli!nome
    RecAux("Campo10") = FormataData(RecCli!dtnasc)
    RecAux("Campo11") = RecCli!cpf
    RecAux("Campo12") = RecCli!telefone
    RecAux("Campo13") = RecCli!endereco
    RecAux("Campo14") = RecCli!bairro
    RecAux("Campo15") = RecCli!cidade & "/" & RecCli!Estado
    RecAux("Campo16") = RecCli!cep
    RecAux("Campo17") = VLStrNomeMed
    RecAux("Campo18") = VLStrCRMMed
    RecAux("Campo19") = VLStrCPFMed
    RecAux("Campo20") = FormataData(RecCred!dtcred)
    RecAux("Campo21") = RecCred!tipocred
    RecAux("Campo22") = FormataMoeda(RecCred!valorvenda)
    If RecCred!juros = "" Then
        RecAux("Campo23") = ""
    Else
        RecAux("Campo23") = FormataNum(RecCred!juros) & "%"
    End If
    RecAux("Campo24") = FormataMoeda(RecCred!valortotal)
    RecAux("Campo25") = FormataNum(RecCred!parcela)
    RecAux("Campo26") = RecCred!tipoentr
    If RecCred!valorentr = "" Then
        RecAux("Campo27") = ""
    Else
        RecAux("Campo27") = FormataMoeda(RecCred!valorentr)
    End If
    
    parctemp = 1
    
    Do While parctemp <= RecCredParc.RecordCount
        
        If parctemp = 1 Then
            RecAux("Campo28") = FormataData(RecCredParc!vencimento)
            RecAux("Campo29") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 2 Then
            RecAux("Campo30") = FormataData(RecCredParc!vencimento)
            RecAux("Campo31") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 3 Then
            RecAux("Campo32") = FormataData(RecCredParc!vencimento)
            RecAux("Campo33") = FormataMoeda(RecCredParc!valor)
                
        ElseIf parctemp = 4 Then
            RecAux("Campo34") = FormataData(RecCredParc!vencimento)
            RecAux("Campo35") = FormataMoeda(RecCredParc!valor)
                
        ElseIf parctemp = 5 Then
            RecAux("Campo36") = FormataData(RecCredParc!vencimento)
            RecAux("Campo37") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 6 Then
            RecAux("Campo38") = FormataData(RecCredParc!vencimento)
            RecAux("Campo39") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 7 Then
            RecAux("Campo40") = FormataData(RecCredParc!vencimento)
            RecAux("Campo41") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 8 Then
            RecAux("Campo42") = FormataData(RecCredParc!vencimento)
            RecAux("Campo43") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 9 Then
            RecAux("Campo44") = FormataData(RecCredParc!vencimento)
            RecAux("Campo45") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 10 Then
            RecAux("Campo46") = FormataData(RecCredParc!vencimento)
            RecAux("Campo47") = FormataMoeda(RecCredParc!valor)
        End If
        
        parctemp = parctemp + 1
        RecCredParc.MoveNext
    Loop
    
    RecAux.Update
    
    Desconecta
        
    rptPropCredito.Show
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaParcelas()
    '=== Parcela de cheque ===
    CboPrazoChqParc.AddItem ("00")
    CboPrazoChqParc.AddItem ("01")
    CboPrazoChqParc.AddItem ("02")
    CboPrazoChqParc.AddItem ("03")
    CboPrazoChqParc.AddItem ("04")
    CboPrazoChqParc.AddItem ("05")
    CboPrazoChqParc.AddItem ("06")
    CboPrazoChqParc.AddItem ("07")
    CboPrazoChqParc.AddItem ("08")
    CboPrazoChqParc.AddItem ("09")
    CboPrazoChqParc.AddItem ("10")
    CboPrazoChqParc.AddItem ("11")
    CboPrazoChqParc.AddItem ("12")
    
    CboPrazoChqParc.Text = "00"
    
    '=== Parcela de carnê ===
    CboPrazoCarParc.AddItem ("00")
    CboPrazoCarParc.AddItem ("01")
    CboPrazoCarParc.AddItem ("02")
    CboPrazoCarParc.AddItem ("03")
    CboPrazoCarParc.AddItem ("04")
    CboPrazoCarParc.AddItem ("05")
    CboPrazoCarParc.AddItem ("06")
    CboPrazoCarParc.AddItem ("07")
    CboPrazoCarParc.AddItem ("08")
    CboPrazoCarParc.AddItem ("09")
    CboPrazoCarParc.AddItem ("10")
    CboPrazoCarParc.AddItem ("11")
    CboPrazoCarParc.AddItem ("12")
    
    CboPrazoCarParc.Text = "00"
    
End Sub
