VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmProduto_Alt 
   Caption         =   "Alteração de Produtos"
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
   Icon            =   "FrmProduto_Alt.frx":0000
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
      TabIndex        =   26
      Top             =   3960
      Width           =   6495
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1320
         OleObjectBlob   =   "FrmProduto_Alt.frx":0CCA
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
         TabIndex        =   24
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
         TabIndex        =   23
         ToolTipText     =   "Efetuar alteração"
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
      TabIndex        =   25
      Top             =   120
      Width           =   6495
      Begin VB.Frame FraArmacao 
         Caption         =   "Armação"
         Height          =   2775
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   6255
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
            ToolTipText     =   "Nome da griffe do produto"
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
            OleObjectBlob   =   "FrmProduto_Alt.frx":0EFE
            TabIndex        =   28
            Top             =   360
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":0F6C
            TabIndex        =   29
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmProduto_Alt.frx":0FD2
            TabIndex        =   30
            Top             =   1320
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   4560
            OleObjectBlob   =   "FrmProduto_Alt.frx":1032
            TabIndex        =   31
            Top             =   1320
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":1098
            TabIndex        =   32
            Top             =   1320
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":10FE
            TabIndex        =   33
            Top             =   1800
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "FrmProduto_Alt.frx":116E
            TabIndex        =   34
            Top             =   1800
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":11E2
            TabIndex        =   35
            Top             =   2280
            Width           =   1215
         End
      End
      Begin VB.Frame FraLenteC 
         Caption         =   "Lente de contato"
         Height          =   2775
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox TxtPrecoLenteC 
            Height          =   285
            Left            =   1320
            TabIndex        =   17
            ToolTipText     =   "Preço do fabricante"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton CmdIncluirFornLC 
            Caption         =   "+"
            Height          =   255
            Left            =   5760
            TabIndex        =   14
            ToolTipText     =   "Adicionar fornecedor"
            Top             =   360
            Width           =   375
         End
         Begin VB.ComboBox CboFornLenteC 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Fornecedor de lentes de contato"
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox TxtChaveLenteC 
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            ToolTipText     =   "Chave para lentes de contato"
            Top             =   1320
            Width           =   4335
         End
         Begin VB.ComboBox CboTipoLenteC 
            Height          =   315
            Left            =   1320
            TabIndex        =   15
            ToolTipText     =   "Tipo de lentes de contato"
            Top             =   840
            Width           =   4335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":1256
            TabIndex        =   37
            Top             =   360
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":12C4
            TabIndex        =   38
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":1326
            TabIndex        =   39
            Top             =   1320
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":138A
            TabIndex        =   40
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame FraLente 
         Caption         =   "Lente"
         Height          =   2775
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox TxtPrecoLente 
            Height          =   285
            Left            =   1320
            TabIndex        =   22
            ToolTipText     =   "Preço do fabricante"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.ComboBox CboTipoLente 
            Height          =   315
            Left            =   1320
            TabIndex        =   20
            ToolTipText     =   "Tipo de lentes"
            Top             =   840
            Width           =   4335
         End
         Begin VB.TextBox TxtChaveLente 
            Height          =   285
            Left            =   1320
            TabIndex        =   21
            ToolTipText     =   "Chave para lentes"
            Top             =   1320
            Width           =   4335
         End
         Begin VB.ComboBox CboFornLente 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Fornecedor de lentes"
            Top             =   360
            Width           =   4335
         End
         Begin VB.CommandButton CmdIncluirFornL 
            Caption         =   "+"
            Height          =   255
            Left            =   5760
            TabIndex        =   19
            ToolTipText     =   "Adicionar fornecedor"
            Top             =   360
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":13FE
            TabIndex        =   42
            Top             =   360
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":146C
            TabIndex        =   43
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":14CE
            TabIndex        =   44
            Top             =   1320
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmProduto_Alt.frx":1532
            TabIndex        =   45
            Top             =   1800
            Width           =   1215
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
         OleObjectBlob   =   "FrmProduto_Alt.frx":15A6
         TabIndex        =   46
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmProduto_Alt"
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

Private Sub CmdIncluirFornL_Click()
    VGStrIncluirProd = "forn"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirFornLC_Click()
    VGStrIncluirProd = "forn"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdIncluirGriffe_Click()
    VGStrIncluirProd = "prod"
    FrmEstoque_Inc_Extra.Show
End Sub

Private Sub CmdIncluirFornArm_Click()
    VGStrIncluirProd = "forn"
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdOK_Click()
    Dim RecLente As New ADODB.Recordset
    Dim RecLenteC As New ADODB.Recordset
    Dim RecArm As New ADODB.Recordset
    
    Conecta
    
    If CboFornLente.Text <> "" Or CboTipoLente.Text <> "" Or TxtChaveLente.Text <> "" Or TxtPrecoLente.Text <> "" Then
        StrSql = "Select * from tb_Produto where CodProd=" & VGIntCodProd
        RecLente.Open StrSql, vgCon, 1, 3
        
        If CboFornLente.Text <> "" Then
            RecLente("CodForn") = Trim(Mid(CboFornLente.Text, Len(CboFornLente.Text) - 10))
        Else
            RecLente("CodForn") = 0
        End If
        RecLente("TipoProd") = "Lente"
        RecLente("Tipo") = CboTipoLente.Text
        RecLente("Chave") = TxtChaveLente.Text
        RecLente("PrecoFabric") = TxtPrecoLente.Text
        RecLente.Update
    End If
    
    If CboFornLenteC.Text <> "" Or CboTipoLenteC.Text <> "" Or TxtChaveLenteC.Text <> "" Or TxtPrecoLenteC.Text <> "" Then
        StrSql = "Select * from tb_Produto where CodProd=" & VGIntCodProd
        RecLenteC.Open StrSql, vgCon, 1, 3
        
        If CboFornLenteC.Text <> "" Then
            RecLenteC("CodForn") = Trim(Mid(CboFornLenteC.Text, Len(CboFornLenteC.Text) - 10))
        Else
            RecLenteC("CodForn") = 0
        End If
        RecLenteC("TipoProd") = "Lente de contato"
        RecLenteC("Tipo") = CboTipoLenteC.Text
        RecLenteC("Chave") = TxtChaveLenteC.Text
        RecLenteC("PrecoFabric") = TxtPrecoLenteC.Text
        RecLenteC.Update
    End If
    
    If CboFornArm.Text <> "" Or CboGrifArm.Text <> "" Or TxtModeloArm.Text <> "" Or TxtCorArm.Text <> "" Or TxtNumArm.Text <> "" Or TxtAroArm.Text <> "" Or TxtPonteArm.Text <> "" Or TxtPrecoArm.Text <> "" Then
        StrSql = "Select * from tb_Produto where CodProd=" & VGIntCodProd
        RecArm.Open StrSql, vgCon, 1, 3
        
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
        RecArm("PrecoFabric") = TxtPrecoArm.Text
        RecArm.Update
    End If
    
    Desconecta
    
    VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Ótica 2004 - Informação")
    
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
    
End Sub

Private Sub Form_Resize()
  FrmProduto_Alt.Left = (MDIPrincipal.Width / 2) - (FrmProduto_Alt.Width / 2)
  FrmProduto_Alt.Top = (MDIPrincipal.Height / 3) - (FrmProduto_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 5310
    Width = 6840
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Call MontaCbos
    
    Dim RecProd As New ADODB.Recordset
    Dim RecForn As New ADODB.Recordset
    Dim RecGrif As New ADODB.Recordset
    
    Conecta
    
    StrSql = "Select * from tb_Produto where CodProd=" & VGIntCodProd
    RecProd.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select CodForn,Nome from tb_Fornecedor where CodForn=" & RecProd.Fields.Item(1).Value
    RecForn.Open StrSql, vgCon, 1, 3
    
    If RecProd.Fields.Item(3).Value = "Armação" Then
        StrSql = "Select CodGriffe,Nome from tb_Griffe where CodGriffe=" & RecProd.Fields.Item(2).Value
        RecGrif.Open StrSql, vgCon, 1, 3
        
        OptArmacao.Value = True
        
        CboFornArm.Text = RecForn.Fields.Item(1).Value & "                                                                                                 " & RecForn.Fields.Item(0).Value
        CboGrifArm.Text = RecGrif.Fields.Item(1).Value & "                                                                                                 " & RecGrif.Fields.Item(0).Value
        TxtCorArm.Text = RecProd.Fields.Item(4).Value
        TxtNumArm.Text = RecProd.Fields.Item(5).Value
        TxtModeloArm.Text = RecProd.Fields.Item(6).Value
        TxtAroArm.Text = RecProd.Fields.Item(7).Value
        TxtPonteArm.Text = RecProd.Fields.Item(8).Value
        TxtPrecoArm.Text = RecProd.Fields.Item(11).Value
    
    ElseIf RecProd.Fields.Item(3).Value = "Lente" Then
        OptLente.Value = True
        
        CboFornLente.Text = RecForn.Fields.Item(1).Value & "                                                                                                 " & RecForn.Fields.Item(0).Value
        CboTipoLente.Text = RecProd.Fields.Item(9).Value
        TxtChaveLente.Text = RecProd.Fields.Item(10).Value
        TxtPrecoLente.Text = RecProd.Fields.Item(11).Value
    
    ElseIf RecProd.Fields.Item(3).Value = "Lente de contato" Then
        OptLenteC.Value = True
        
        CboFornLenteC.Text = RecForn.Fields.Item(1).Value & "                                                                                                 " & RecForn.Fields.Item(0).Value
        CboTipoLenteC.Text = RecProd.Fields.Item(9).Value
        TxtChaveLenteC.Text = RecProd.Fields.Item(10).Value
        TxtPrecoLenteC.Text = RecProd.Fields.Item(11).Value
    End If
    
    Desconecta
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaCbos()
    Dim RecCbo As New ADODB.Recordset
    
    CboFornLente.Clear
    CboTipoLente.Clear
    CboFornLenteC.Clear
    CboTipoLenteC.Clear
    CboFornArm.Clear
    CboGrifArm.Clear
    
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
End Sub

Private Sub OptLente_Click()
    FraArmacao.Visible = False
    FraLente.Visible = True
    FraLenteC.Visible = False
End Sub

Private Sub OptLenteC_Click()
    FraArmacao.Visible = False
    FraLente.Visible = False
    FraLenteC.Visible = True
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
