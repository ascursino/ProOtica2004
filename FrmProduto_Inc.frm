VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmProduto_Inc 
   Caption         =   "Inclusão de Produtos"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
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
   Icon            =   "FrmProduto_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6720
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
      TabIndex        =   29
      Top             =   3960
      Width           =   6495
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1320
         OleObjectBlob   =   "FrmProduto_Inc.frx":0CCA
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
         Left            =   5160
         TabIndex        =   27
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
         Left            =   3840
         TabIndex        =   26
         ToolTipText     =   "Efetuar inclusão"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   6495
      Begin VB.Frame FraArmacao 
         Caption         =   "Armação"
         Height          =   2775
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   6255
         Begin VB.ComboBox CboMoedaArm 
            Height          =   315
            ItemData        =   "FrmProduto_Inc.frx":0EFE
            Left            =   4680
            List            =   "FrmProduto_Inc.frx":0F08
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Moeda"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox TxtPrecoArm 
            Height          =   285
            Left            =   1440
            TabIndex        =   12
            ToolTipText     =   "Preço do fabricante"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox CboGrifArm 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Nome da griffe"
            Top             =   840
            Width           =   4215
         End
         Begin VB.CommandButton CmdIncluirGriffe 
            Caption         =   "+"
            Height          =   255
            Left            =   5760
            TabIndex        =   6
            ToolTipText     =   "Adicionar griffe"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox TxtCorArm 
            Height          =   285
            Left            =   3600
            TabIndex        =   8
            ToolTipText     =   "Cor da armação"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtNumArm 
            Height          =   285
            Left            =   5400
            TabIndex        =   9
            ToolTipText     =   "Número da armação"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox TxtPonteArm 
            Height          =   285
            Left            =   4680
            TabIndex        =   11
            ToolTipText     =   "Tamanho da ponte"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox TxtAroArm 
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            ToolTipText     =   "Tamanho do aro"
            Top             =   1800
            Width           =   735
         End
         Begin VB.TextBox TxtModeloArm 
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            ToolTipText     =   "Modelo da armação"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.ComboBox CboFornArm 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Fornecedor de armações"
            Top             =   360
            Width           =   4215
         End
         Begin VB.CommandButton CmdIncluirFornArm 
            Caption         =   "+"
            Height          =   255
            Left            =   5760
            TabIndex        =   4
            ToolTipText     =   "Adicionar fornecedor"
            Top             =   360
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":0F19
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":0F87
            TabIndex        =   32
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmProduto_Inc.frx":0FED
            TabIndex        =   33
            Top             =   1320
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   4560
            OleObjectBlob   =   "FrmProduto_Inc.frx":104D
            TabIndex        =   34
            Top             =   1320
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":10B3
            TabIndex        =   35
            Top             =   1320
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":1119
            TabIndex        =   36
            Top             =   1800
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmProduto_Inc.frx":1189
            TabIndex        =   37
            Top             =   1800
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":11FD
            TabIndex        =   38
            Top             =   2280
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmProduto_Inc.frx":1271
            TabIndex        =   50
            Top             =   2280
            Width           =   1215
         End
      End
      Begin VB.Frame FraLenteC 
         Caption         =   "Lente de contato"
         Height          =   2775
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   6255
         Begin VB.ComboBox CboMoedaLenteC 
            Height          =   315
            ItemData        =   "FrmProduto_Inc.frx":12D5
            Left            =   4200
            List            =   "FrmProduto_Inc.frx":12DF
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "Moeda"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox TxtPrecoLenteC 
            Height          =   285
            Left            =   1320
            TabIndex        =   18
            ToolTipText     =   "Preço do fabricante"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton CmdIncluirFornLC 
            Caption         =   "+"
            Height          =   255
            Left            =   5760
            TabIndex        =   15
            ToolTipText     =   "Adicionar fornecedor"
            Top             =   360
            Width           =   375
         End
         Begin VB.ComboBox CboFornLenteC 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Fornecedor de lentes de contato"
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox TxtChaveLenteC 
            Height          =   285
            Left            =   1320
            TabIndex        =   17
            ToolTipText     =   "Chave para lentes de contato"
            Top             =   1320
            Width           =   4335
         End
         Begin VB.ComboBox CboTipoLenteC 
            Height          =   315
            Left            =   1320
            TabIndex        =   16
            ToolTipText     =   "Tipo de lentes de contato"
            Top             =   840
            Width           =   4335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":12F0
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":135E
            TabIndex        =   41
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":13C0
            TabIndex        =   42
            Top             =   1320
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":1424
            TabIndex        =   43
            Top             =   1800
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   3360
            OleObjectBlob   =   "FrmProduto_Inc.frx":1498
            TabIndex        =   51
            Top             =   1800
            Width           =   735
         End
      End
      Begin VB.Frame FraLente 
         Caption         =   "Lente"
         Height          =   2775
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Visible         =   0   'False
         Width           =   6255
         Begin VB.ComboBox CboMoedaLente 
            Height          =   315
            ItemData        =   "FrmProduto_Inc.frx":14FC
            Left            =   4200
            List            =   "FrmProduto_Inc.frx":1506
            Style           =   2  'Dropdown List
            TabIndex        =   25
            ToolTipText     =   "Moeda"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox TxtPrecoLente 
            Height          =   285
            Left            =   1320
            TabIndex        =   24
            ToolTipText     =   "Preço do fabricante"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.ComboBox CboTipoLente 
            Height          =   315
            Left            =   1320
            TabIndex        =   22
            ToolTipText     =   "Tipo de lentes"
            Top             =   840
            Width           =   4335
         End
         Begin VB.TextBox TxtChaveLente 
            Height          =   285
            Left            =   1320
            TabIndex        =   23
            ToolTipText     =   "Chave para lentes"
            Top             =   1320
            Width           =   4335
         End
         Begin VB.ComboBox CboFornLente 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   "Fornecedor de lentes"
            Top             =   360
            Width           =   4335
         End
         Begin VB.CommandButton CmdIncluirFornL 
            Caption         =   "+"
            Height          =   255
            Left            =   5760
            TabIndex        =   21
            ToolTipText     =   "Adicionar fornecedor"
            Top             =   360
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":1517
            TabIndex        =   45
            Top             =   360
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":1585
            TabIndex        =   46
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":15E7
            TabIndex        =   47
            Top             =   1320
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Inc.frx":164B
            TabIndex        =   48
            Top             =   1800
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   3360
            OleObjectBlob   =   "FrmProduto_Inc.frx":16BF
            TabIndex        =   52
            Top             =   1800
            Width           =   735
         End
      End
      Begin VB.OptionButton OptLenteC 
         Caption         =   "Lente de contato"
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         ToolTipText     =   "Lente de contato"
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton OptLente 
         Caption         =   "Lente"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         ToolTipText     =   "Lente"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptArmacao 
         Caption         =   "Armação"
         Height          =   255
         Left            =   1800
         TabIndex        =   0
         ToolTipText     =   "Armação"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmProduto_Inc.frx":1723
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmProduto_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdIncluirGriffeArm_Click()
    VGStrEstoqueIncExtra = "Armação"
    FrmEstoque_Inc_Extra.Show
End Sub

Private Sub CmdIncluirLabArm_Click()
    VGStrEstoqueIncExtra = "Armação"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirLabLente_Click()
    VGStrEstoqueIncExtra = "Lente"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirLabLenteC_Click()
    VGStrEstoqueIncExtra = "Lente de contato"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirModeloLente_Click()
    VGStrEstoqueIncExtra = "ModeloLente"
    FrmEstoque_Inc_Extra.Show
End Sub

Private Sub CmdIncluirModeloLenteC_Click()
    VGStrEstoqueIncExtra = "ModeloLenteC"
    FrmEstoque_Inc_Extra.Show
End Sub

Private Sub CmdIncluirTipoLente_Click()
    VGStrEstoqueIncExtra = "TipoLente"
    FrmEstoque_Inc_Extra.Show
End Sub

Private Sub CmdIncluirTipoLenteC_Click()
    VGStrEstoqueIncExtra = "TipoLenteC"
    FrmEstoque_Inc_Extra.Show
End Sub

Private Sub CmdIncluirFornL_Click()
    VGStrIncluirProd = "prod"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirFornLC_Click()
    VGStrIncluirProd = "prod"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirGriffe_Click()
    VGStrIncluirProd = "prod"
    FrmEstoque_Inc_Extra.Show
End Sub

Private Sub CmdIncluirFornArm_Click()
    VGStrIncluirProd = "prod"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdOK_Click()
    Dim RecLente As New ADODB.Recordset
    Dim RecLenteC As New ADODB.Recordset
    Dim RecArm As New ADODB.Recordset
    
    Conecta
    
    If CboFornLente.Text <> "" Or CboTipoLente.Text <> "" Or TxtChaveLente.Text <> "" Or TxtPrecoLente.Text <> "" Then
        StrSql = "Select * from tb_Produto"
        RecLente.Open StrSql, vgCon, 1, 3
        
        RecLente.AddNew
        If CboFornLente.Text <> "" Then
            RecLente("CodForn") = Trim(Mid(CboFornLente.Text, Len(CboFornLente.Text) - 10))
        Else
            RecLente("CodForn") = 0
        End If
        RecLente("TipoProd") = "Lente"
        RecLente("Tipo") = CboTipoLente.Text
        RecLente("Chave") = TxtChaveLente.Text
        RecLente("PrecoFabric") = CCur(TxtPrecoLente.Text)
        RecLente("Moeda") = CboMoedaLente.Text
        RecLente.Update
    End If
    
    If CboFornLenteC.Text <> "" Or CboTipoLenteC.Text <> "" Or TxtChaveLenteC.Text <> "" Or TxtPrecoLenteC.Text <> "" Then
        StrSql = "Select * from tb_Produto"
        RecLenteC.Open StrSql, vgCon, 1, 3
        
        RecLenteC.AddNew
        If CboFornLenteC.Text <> "" Then
            RecLenteC("CodForn") = Trim(Mid(CboFornLenteC.Text, Len(CboFornLenteC.Text) - 10))
        Else
            RecLenteC("CodForn") = 0
        End If
        RecLenteC("TipoProd") = "Lente de contato"
        RecLenteC("Tipo") = CboTipoLenteC.Text
        RecLenteC("Chave") = TxtChaveLenteC.Text
        RecLenteC("PrecoFabric") = CCur(TxtPrecoLenteC.Text)
        RecLenteC("Moeda") = CboMoedaLenteC.Text
        RecLenteC.Update
    End If
    
    If CboFornArm.Text <> "" Or CboGrifArm.Text <> "" Or TxtModeloArm.Text <> "" Or TxtCorArm.Text <> "" Or TxtNumArm.Text <> "" Or TxtAroArm.Text <> "" Or TxtPonteArm.Text <> "" Or TxtPrecoArm.Text <> "" Then
        StrSql = "Select * from tb_Produto"
        RecArm.Open StrSql, vgCon, 1, 3
        
        RecArm.AddNew
        If CboFornArm.Text <> "" Then
            RecArm("CodForn") = Trim(Mid(CboFornArm.Text, Len(CboFornArm.Text) - 10))
        Else
            RecArm("CodForn") = 0
        End If
        If CboGrifArm.Text <> "" Then
            RecArm("CodGriffe") = Trim(Mid(CboGrifArm.Text, Len(CboGrifArm.Text) - 10))
        Else
            RecArm("CodGriffe") = 0
        End If
        RecArm("TipoProd") = "Armação"
        RecArm("Cor") = TxtCorArm.Text
        RecArm("Numero") = TxtNumArm.Text
        RecArm("Modelo") = TxtModeloArm.Text
        RecArm("TamAro") = TxtAroArm.Text
        RecArm("TamPonte") = TxtPonteArm.Text
        RecArm("PrecoFabric") = CCur(TxtPrecoArm.Text)
        RecArm("Moeda") = CboMoedaArm.Text
        RecArm.Update
    End If
    
    Desconecta
    
    VPStrResponse = MsgBox("Deseja incluir informações do(s) produto(s) no estoque agora?", vbYesNo, "Pró Ótica 2004 - Informação")
    If VPStrResponse = vbYes Then
        CboFornLente.ListIndex = 0
        CboTipoLente.Text = ""
        TxtChaveLente.Text = ""
        TxtPrecoLente.Text = ""
        CboFornLenteC.ListIndex = 0
        CboTipoLenteC.Text = ""
        TxtChaveLenteC.Text = ""
        TxtPrecoLenteC.Text = ""
        CboFornArm.ListIndex = 0
        CboGrifArm.ListIndex = 0
        TxtCorArm.Text = ""
        TxtNumArm.Text = ""
        TxtModeloArm.Text = ""
        TxtAroArm.Text = ""
        TxtPonteArm.Text = ""
        TxtPrecoArm.Text = ""

        FrmEstoque_Inc_Alt.Show
    Else
        CboFornLente.ListIndex = 0
        CboTipoLente.Text = ""
        TxtChaveLente.Text = ""
        TxtPrecoLente.Text = ""
        CboFornLenteC.ListIndex = 0
        CboTipoLenteC.Text = ""
        TxtChaveLenteC.Text = ""
        TxtPrecoLenteC.Text = ""
        CboFornArm.ListIndex = 0
        CboGrifArm.ListIndex = 0
        TxtCorArm.Text = ""
        TxtNumArm.Text = ""
        TxtModeloArm.Text = ""
        TxtAroArm.Text = ""
        TxtPonteArm.Text = ""
        TxtPrecoArm.Text = ""
    End If
End Sub

Private Sub Form_Resize()
  FrmProduto_Inc.Left = (MDIPrincipal.Width / 2) - (FrmProduto_Inc.Width / 2)
  FrmProduto_Inc.Top = (MDIPrincipal.Height / 3) - (FrmProduto_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 5310
    Width = 6840
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Call MontaCboForn
    Call MontaCboGrif
    Call MontaCboTipo
    
    CboMoedaArm.Text = "Real"
    CboMoedaLenteC.Text = "Real"
    CboMoedaLente.Text = "Real"
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaCboForn()
    Dim RecCbo As New ADODB.Recordset
    
    CboFornLente.Clear
    CboFornLenteC.Clear
    CboFornArm.Clear
    
    Conecta
    
    StrSql = "Select CodForn,Nome from tb_Fornecedor"
    RecCbo.Open StrSql, vgCon, 1, 3
    
    CboFornLente.AddItem ("")
    CboFornLenteC.AddItem ("")
    CboFornArm.AddItem ("")
    Do While Not RecCbo.EOF
        CboFornLente.AddItem (RecCbo.Fields.Item(1).Value & "                                                                                                 " & RecCbo.Fields.Item(0).Value)
        CboFornLenteC.AddItem (RecCbo.Fields.Item(1).Value & "                                                                                                 " & RecCbo.Fields.Item(0).Value)
        CboFornArm.AddItem (RecCbo.Fields.Item(1).Value & "                                                                                                 " & RecCbo.Fields.Item(0).Value)
        RecCbo.MoveNext
    Loop
    
    RecCbo.Close
    
    Desconecta

End Sub

Sub MontaCboTipo()
    Dim RecCbo As New ADODB.Recordset
    
    CboTipoLente.Clear
    CboTipoLenteC.Clear
    
    Conecta
    
    StrSql = "Select distinct Tipo from tb_Produto"
    RecCbo.Open StrSql, vgCon, 1, 3
    
    CboTipoLente.AddItem ("")
    CboTipoLenteC.AddItem ("")
    Do While Not RecCbo.EOF
        If RecCbo.Fields.Item(0).Value <> "" And IsNull(RecCbo.Fields.Item(0).Value) = False Then
            CboTipoLente.AddItem (RecCbo.Fields.Item(0).Value)
            CboTipoLenteC.AddItem (RecCbo.Fields.Item(0).Value)
        End If
        RecCbo.MoveNext
    Loop
    
    RecCbo.Close
    
    Desconecta

End Sub

Sub MontaCboGrif()
    Dim RecCbo As New ADODB.Recordset
    
    CboGrifArm.Clear
    
    Conecta
    
    StrSql = "Select CodGriffe,Nome from tb_Griffe"
    RecCbo.Open StrSql, vgCon, 1, 3
    
    CboGrifArm.AddItem ("")
    Do While Not RecCbo.EOF
        CboGrifArm.AddItem (RecCbo.Fields.Item(1).Value & "                                                                                                 " & RecCbo.Fields.Item(0).Value)
        RecCbo.MoveNext
    Loop
    Desconecta

End Sub

Private Sub OptArmacao_Click()
    FraArmacao.Visible = True
    FraLente.Visible = False
    FraLenteC.Visible = False
    
    CboFornArm.SetFocus
End Sub

Private Sub OptLente_Click()
    FraArmacao.Visible = False
    FraLente.Visible = True
    FraLenteC.Visible = False
    
    CboFornLente.SetFocus
End Sub

Private Sub OptLenteC_Click()
    FraArmacao.Visible = False
    FraLente.Visible = False
    FraLenteC.Visible = True
    
    CboFornLenteC.SetFocus
End Sub

Private Sub TxtPrecoArm_LostFocus()
    If TxtPrecoArm.Text <> "" Then
       TxtPrecoArm.Text = FormataMoeda(TxtPrecoArm.Text)
    End If
End Sub

Private Sub TxtPrecoLente_LostFocus()
    If TxtPrecoLente.Text <> "" Then
       TxtPrecoLente.Text = FormataMoeda(TxtPrecoLente.Text)
    End If
End Sub

Private Sub TxtPrecoLenteC_LostFocus()
    If TxtPrecoLenteC.Text <> "" Then
       TxtPrecoLenteC.Text = FormataMoeda(TxtPrecoLenteC.Text)
    End If
End Sub
