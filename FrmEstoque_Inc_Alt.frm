VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmEstoque_Inc_Alt 
   Caption         =   "Inclusão/Alteração de Produtos no Estoque"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
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
   Icon            =   "FrmEstoque_Inc_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8760
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
      TabIndex        =   11
      Top             =   4440
      Width           =   8535
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2280
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":0CCA
         Top             =   240
      End
      Begin VB.CommandButton CmdAlterar 
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
         Left            =   5880
         TabIndex        =   7
         ToolTipText     =   "Alterar informações do produto em estoque"
         Top             =   240
         Width           =   1095
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
         Left            =   7200
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdIncluir 
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
         Left            =   4560
         TabIndex        =   6
         ToolTipText     =   "Incluir informações do produto em estoque"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   8535
      Begin VB.TextBox TxtPrecoVenda 
         Height          =   285
         Left            =   6960
         TabIndex        =   5
         ToolTipText     =   "Preço de venda do produto"
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox TxtPrecoFabric 
         Height          =   285
         Left            =   6960
         TabIndex        =   3
         ToolTipText     =   "Preço do fabricante do produto"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox TxtQtdeMin 
         Height          =   285
         Left            =   6960
         TabIndex        =   1
         ToolTipText     =   "Quantide mínima recomendada do produto no estoque"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TxtProd 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   5640
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "FrmEstoque_Inc_Alt.frx":0EFE
         ToolTipText     =   "Descrição do produto"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox TxtQtdeEst 
         Height          =   285
         Left            =   6960
         TabIndex        =   2
         ToolTipText     =   "Quantidade do produto em estoque"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox CboMult 
         Height          =   315
         ItemData        =   "FrmEstoque_Inc_Alt.frx":0F0A
         Left            =   6960
         List            =   "FrmEstoque_Inc_Alt.frx":0F0C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Multiplicação para gerar preço de venda"
         Top             =   3240
         Width           =   1455
      End
      Begin FPSpread.vaSpread GridProduto 
         Height          =   3855
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5415
         _Version        =   393216
         _ExtentX        =   9551
         _ExtentY        =   6800
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
         SpreadDesigner  =   "FrmEstoque_Inc_Alt.frx":0F0E
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":125F
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":12C7
         TabIndex        =   13
         Top             =   2760
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":133B
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":13AB
         TabIndex        =   15
         Top             =   2280
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":141D
         TabIndex        =   16
         Top             =   3240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   5640
         OleObjectBlob   =   "FrmEstoque_Inc_Alt.frx":148D
         TabIndex        =   17
         Top             =   3720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmEstoque_Inc_Alt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public VPStrMult As String

Private Sub CboMult_Click()
    If VPStrMult = "" Then
        If TxtPrecoFabric.Text <> "" And CboMult.Text <> "" Then
            TxtPrecoVenda.Text = FormataMoeda(Mid(TxtPrecoFabric.Text, 3) * Val(CboMult.Text))
        ElseIf TxtPrecoFabric.Text <> "" And CboMult.Text = "" Then
            TxtPrecoVenda.Text = FormataMoeda(TxtPrecoFabric.Text)
        End If
    End If
    VPStrMult = ""
End Sub

Private Sub CmdAlterar_Click()
    
    Conecta
    
    Dim RecEst As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    
    StrSql = "Select QtdeMin,QtdeProd,Multiplicar,PrecoVenda from tb_Estoque where CodProd=" & VGIntCodProd
    RecEst.Open StrSql, vgCon, 1, 3
    
    If TxtQtdeMin.Text = "" Then
        RecEst("QtdeMin") = 0
    Else
        RecEst("QtdeMin") = TxtQtdeMin.Text
    End If
    
    If TxtQtdeEst.Text = "" Then
        RecEst("QtdeProd") = 0
    Else
        RecEst("QtdeProd") = TxtQtdeEst.Text
    End If
    
    If CboMult.Text = "" Then
        RecEst("Multiplicar") = 1
    Else
        RecEst("Multiplicar") = CboMult.Text
    End If
    
    If TxtPrecoVenda.Text <> "" Then
        RecEst("PrecoVenda") = TxtPrecoVenda.Text
    End If
    RecEst.Update
    
    If TxtPrecoFabric.Text <> "" Then
        StrSql = "Select PrecoFabric from tb_Produto where CodProd=" & VGIntCodProd
        RecProd.Open StrSql, vgCon, 1, 3
    
        RecProd("PrecoFabric") = TxtPrecoFabric.Text
        RecProd.Update
        
        RecProd.Close
    End If
    
    Desconecta
    
    VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Ótica 2004 - Informação")
    
    TxtProd.Text = ""
    TxtQtdeMin.Text = ""
    TxtQtdeEst.Text = ""
    TxtPrecoFabric.Text = ""
    CboMult.ListIndex = 0
    TxtPrecoVenda.Text = ""
    
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdIncluir_Click()
    Dim RecProd As New ADODB.Recordset
    
    If TxtQtdeMin.Text = "" Or TxtQtdeEst.Text = "" Or TxtPrecoVenda.Text = "" Then
        VPStrBox = MsgBox("Não pode conter campos em branco.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        Conecta
        
        StrSql = "Select * from tb_Estoque"
        RecProd.Open StrSql, vgCon, 1, 3
        
        RecProd.AddNew
        RecProd("CodProd") = VGIntCodProd
        
        If TxtQtdeMin.Text = "" Then
            RecProd("QtdeMin") = 0
        Else
            RecProd("QtdeMin") = TxtQtdeMin.Text
        End If
        
        If TxtQtdeEst.Text = "" Then
            RecProd("QtdeProd") = 0
        Else
            RecProd("QtdeProd") = TxtQtdeEst.Text
        End If
        
        If CboMult.Text = "" Then
            RecProd("Multiplicar") = 0
        Else
            RecProd("Multiplicar") = CboMult.Text
        End If
        
        RecProd("PrecoVenda") = CCur(TxtPrecoVenda.Text)
        RecProd.Update
        
        Desconecta
        
        VPStrBox = MsgBox("Estoque cadastrado.", vbInformation, "Pró Ótica 2004 - Informação")
        
        TxtProd.Text = ""
        TxtQtdeMin.Text = ""
        TxtQtdeEst.Text = ""
        TxtPrecoFabric.Text = ""
        CboMult.ListIndex = 0
        TxtPrecoVenda.Text = ""
    End If
End Sub

Private Sub Form_Resize()
  FrmEstoque_Inc_Alt.Left = (MDIPrincipal.Width / 2) - (FrmEstoque_Inc_Alt.Width / 2)
  FrmEstoque_Inc_Alt.Top = (MDIPrincipal.Height / 3) - (FrmEstoque_Inc_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 5790
    Width = 8880
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    CmdIncluir.Enabled = False
    CmdAlterar.Enabled = False
    TxtProd.Text = ""
    
    Call MontaCboMult
    
    Call MontaGridProduto
    
End Sub

Sub MontaGridProduto()
    Dim VLIntCodProd As Long
    Dim VLIntLinha As Long
    Dim RecProd As New ADODB.Recordset
    Dim RecGrif As New ADODB.Recordset
    Dim Griffe As String
     
    Conecta
    
    StrSql = "Select CodProd,CodGriffe,TipoProd,Cor,Numero,Modelo,TamAro,TamPonte,Tipo,Chave from tb_Produto order by CodProd desc"
    RecProd.Open StrSql, vgCon, 1, 3
    
    If RecProd.EOF Then
        LblNumTotalProd.Caption = "Nenhum produto encontrado."
        
        GridProduto.MaxRows = 0
        
        CmdIncluir.Enabled = False
        CmdAlterar.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridProduto.MaxRows = VLIntLinha
         
        Do While Not RecProd.EOF
                 
            GridProduto.Row = VLIntLinha
            GridProduto.Lock = True
            
            'Produto
            If RecProd!CodGriffe <> 0 And RecProd!CodGriffe <> "" And IsNull(RecProd!CodGriffe) = False Then
                StrSql = "Select Nome From tb_Griffe where CodGriffe=" & RecProd!CodGriffe
                RecGrif.Open StrSql, vgCon, 1, 3
                
                If Not RecGrif.EOF Then
                    Griffe = RecGrif!nome
                Else
                    Griffe = ""
                End If
                
                RecGrif.Close
                
            Else
                Griffe = ""
            End If
            
            GridProduto.Col = 1
            If Griffe = "" Then
            'mostra dados para lentes
                GridProduto.Text = RecProd!tipoprod & " - " & VerificaNulo(RecProd!tipo) & "/" & VerificaNulo(RecProd!chave)
            Else
            'mostra dados para armação
                GridProduto.Text = RecProd!tipoprod & " - " & Griffe & "/" & VerificaNulo(RecProd!cor) & "/" & VerificaNulo(RecProd!Numero) & "/" & VerificaNulo(RecProd!modelo) & "/" & VerificaNulo(RecProd!TamAro) & "/" & VerificaNulo(RecProd!TamPonte)
            End If
            GridProduto.Lock = True
            
            'CodProd
            GridProduto.Col = 2
            GridProduto.Text = Val(RecProd.Fields.Item(0).Value)
            GridProduto.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridProduto.MaxRows = GridProduto.MaxRows + 1
            RecProd.MoveNext
         Loop
         
         GridProduto.MaxRows = GridProduto.MaxRows - 1
    End If
    
    Desconecta
    
    VPStrMult = "primeiro"
    
End Sub

Private Sub GridProduto_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim RecProd As New ADODB.Recordset
    Dim RecPreco As New ADODB.Recordset
    Dim Produto As String
    
    GridProduto.Row = Row
    GridProduto.Col = 2
    VGIntCodProd = GridProduto.Text
    
    GridProduto.Row = Row
    GridProduto.Col = 1
    Produto = GridProduto.Text
    
    Conecta
    
    StrSql = "Select E.QtdeMin,E.QtdeProd,E.Multiplicar,E.PrecoVenda,P.PrecoFabric from tb_Estoque as E, tb_Produto as P where E.CodProd=P.CodProd and E.CodProd=" & VGIntCodProd
    RecProd.Open StrSql, vgCon, 1, 3
    
    If Not RecProd.EOF Then
        TxtProd.Text = Produto
        TxtQtdeMin.Text = RecProd.Fields.Item(0).Value
        TxtQtdeEst.Text = RecProd.Fields.Item(1).Value
        TxtPrecoFabric.Text = FormataMoeda(RecProd.Fields.Item(4).Value)
        
        If RecProd.Fields.Item(2).Value <> "" And RecProd.Fields.Item(2).Value <> 0 And IsNull(RecProd.Fields.Item(2).Value) = False Then
            CboMult.Text = RecProd.Fields.Item(2).Value
        End If
        
        TxtPrecoVenda.Text = FormataMoeda(RecProd.Fields.Item(3).Value)
        
        CmdIncluir.Enabled = False
        CmdAlterar.Enabled = True
        
    Else
        
        StrSql = "Select PrecoFabric from tb_Produto where CodProd=" & VGIntCodProd
        RecPreco.Open StrSql, vgCon, 1, 3
        
        CmdIncluir.Enabled = True
        CmdAlterar.Enabled = False
                
        TxtProd.Text = Produto
        TxtQtdeMin.Text = ""
        TxtQtdeEst.Text = ""
        TxtPrecoFabric.Text = FormataMoeda(RecPreco.Fields.Item(0).Value)
        CboMult.ListIndex = 0
        TxtPrecoVenda.Text = ""
        
        RecPreco.Close
        
        VPStrBox = MsgBox("Não existe informação de estoque" & Chr(13) & "cadastrado para este produto.", vbInformation, "Pró Ótica 2004 - Informação")
        
        TxtQtdeMin.SetFocus
    End If
    
    Desconecta
    
End Sub

Sub MontaCboMult()
    CboMult.AddItem ("")
    CboMult.AddItem ("1")
    CboMult.AddItem ("2")
    CboMult.AddItem ("3")
    CboMult.AddItem ("4")
    CboMult.AddItem ("5")
    CboMult.AddItem ("6")
    CboMult.AddItem ("7")
    CboMult.AddItem ("8")
    CboMult.AddItem ("9")
    CboMult.AddItem ("10")
    CboMult.AddItem ("11")
    CboMult.AddItem ("12")
    CboMult.AddItem ("13")
    CboMult.AddItem ("14")
    CboMult.AddItem ("15")
End Sub

Private Sub TxtPrecoFabric_LostFocus()
    If TxtPrecoFabric <> "" Then
       TxtPrecoFabric.Text = FormataMoeda(TxtPrecoFabric.Text)
    End If
End Sub

Private Sub TxtPrecoVenda_LostFocus()
    If TxtPrecoVenda <> "" Then
       TxtPrecoVenda.Text = FormataMoeda(TxtPrecoVenda.Text)
    End If
End Sub
