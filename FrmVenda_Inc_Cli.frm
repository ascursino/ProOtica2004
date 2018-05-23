VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmVenda_Inc_Cli 
   Caption         =   "Inclusão de Venda - Cliente"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
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
   Icon            =   "FrmVenda_Inc_Cli.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6345
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
      TabIndex        =   5
      Top             =   3960
      Width           =   6135
      Begin VB.CommandButton CmdVendRap 
         Caption         =   "&Venda rápida"
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
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Fazer venda rápida"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "&Incluir cliente"
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
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "Incluir cliente na venda"
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Inc_Cli.frx":0CCA
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
         Left            =   4800
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
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
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6135
      Begin FPSpread.vaSpread GridCliente 
         Height          =   3495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5895
         _Version        =   393216
         _ExtentX        =   10398
         _ExtentY        =   6165
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   2
         MaxRows         =   1
         OperationMode   =   2
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmVenda_Inc_Cli.frx":0EFE
         UserResize      =   1
      End
   End
End
Attribute VB_Name = "FrmVenda_Inc_Cli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    VGStrForm = ""
    VGIntCodCli = 0
    VGStrNomeCli = ""
    
    Unload Me
   
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdIncluir_Click()
    Unload Me
    FrmCliente_Inc.Show
End Sub

Private Sub CmdVendRap_Click()
    VGStrVendaRapida = "sim"
    Unload Me
    FrmVenda_Inc.Show
End Sub

Private Sub Form_Resize()
  FrmVenda_Inc_Cli.Left = (MDIPrincipal.Width / 2) - (FrmVenda_Inc_Cli.Width / 2)
  FrmVenda_Inc_Cli.Top = (MDIPrincipal.Height / 3) - (FrmVenda_Inc_Cli.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 5310
    Width = 6465
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    If VGStrForm = "receita" Then
        Me.Caption = "Inclusão de Receita - Cliente"
    End If
    
    Call MontaGridCliente
    
End Sub

Private Sub GridCliente_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridCliente.Row = Row
    GridCliente.Col = 2
    VGIntCodCli = GridCliente.Text
    
    GridCliente.Row = Row
    GridCliente.Col = 1
    VGStrNomeCli = GridCliente.Text
    
    Unload Me
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
    
    If VGStrForm = "receita" Then
        VGStrForm = ""
        FrmReceita_Inc.Show
    Else
        FrmVenda_Inc.Show
    End If
End Sub

Sub MontaGridCliente()
    
    Dim VLIntLinha As Long
    Dim RecPesq As New ADODB.Recordset
    
    Conecta
    
    StrSql = "Select CodCli,Nome from tb_Cliente"
    RecPesq.Open StrSql, vgCon, 1, 3

    If RecPesq.EOF Then
        GridCliente.Refresh
        GridCliente.MaxRows = 0
    Else
    
        VLIntLinha = 1
        GridCliente.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridCliente.Row = VLIntLinha
            GridCliente.Lock = True
            
            'Cliente
            GridCliente.Col = 1
            GridCliente.Text = VerificaNulo(RecPesq.Fields.Item(1).Value)
            GridCliente.Lock = True
            
            'CodCli
            GridCliente.Col = 2
            GridCliente.Text = Val(RecPesq.Fields.Item(0).Value)
            GridCliente.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridCliente.MaxRows = GridCliente.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         GridCliente.MaxRows = GridCliente.MaxRows - 1
    End If
    
    Desconecta
    
End Sub

