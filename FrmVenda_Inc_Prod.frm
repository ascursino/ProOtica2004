VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmVenda_Inc_Prod 
   Caption         =   "Inclusão de Venda - Produto"
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
   Icon            =   "FrmVenda_Inc_Prod.frx":0000
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
      TabIndex        =   3
      Top             =   3960
      Width           =   6135
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Inc_Prod.frx":0CCA
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
         TabIndex        =   1
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
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin FPSpread.vaSpread GridProduto 
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
         MaxCols         =   3
         MaxRows         =   1
         OperationMode   =   2
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmVenda_Inc_Prod.frx":0EFE
         UserResize      =   1
      End
   End
End
Attribute VB_Name = "FrmVenda_Inc_Prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    VGIntCodProd = 0
    VGStrDescrProd = ""
    
    Unload Me
   
    FrmVenda_Inc.Enabled = True
End Sub

Private Sub Form_Resize()
  FrmVenda_Inc_Prod.Left = (MDIPrincipal.Width / 2) - (FrmVenda_Inc_Prod.Width / 2)
  FrmVenda_Inc_Prod.Top = (MDIPrincipal.Height / 3) - (FrmVenda_Inc_Prod.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 5310
    Width = 6465
    Top = 480
    Left = 4650
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmVenda_Inc.Enabled = False
    
    Call MontaGridProduto
    
End Sub

Private Sub GridProduto_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridProduto.Row = Row
    GridProduto.Col = 3
    If GridProduto.Text <> "" And GridProduto.Text <> "CodProd" Then
        VGIntCodProd = GridProduto.Text
        
        GridProduto.Row = Row
        GridProduto.Col = 2
        VGStrTipoProd = GridProduto.Text
        
        GridProduto.Row = Row
        GridProduto.Col = 1
        VGStrDescrProd = GridProduto.Text
        
        Unload Me
        FrmVenda_Inc.Enabled = True
        
        FrmVenda_Inc.TxtDescrProd.Text = VGStrDescrProd
        FrmVenda_Inc.LblTipoProd.Caption = VGStrTipoProd
        FrmVenda_Inc.TxtPrecoUnit.Text = ""
        FrmVenda_Inc.TxtQtdeProd.Text = ""
        FrmVenda_Inc.TxtValorTotal.Text = ""
        FrmVenda_Inc.LblQtdeProdEst.Visible = False
        FrmVenda_Inc.TxtDescrProd.SetFocus
    End If
End Sub

Sub MontaGridProduto()
    
    Dim VLIntLinha As Long
    Dim RecPesq As New ADODB.Recordset
    Dim RecGriff As New ADODB.Recordset
    Dim Griffe As String
    
    Conecta
    
    StrSql = "Select * from tb_Produto"
    RecPesq.Open StrSql, vgCon, 1, 3
    
    If RecPesq.EOF Then
        GridProduto.Refresh
        GridProduto.MaxRows = 0
    Else
    
        VLIntLinha = 1
        GridProduto.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridProduto.Row = VLIntLinha
            GridProduto.Lock = True
            
            StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecPesq!CodGriffe
            RecGriff.Open StrSql, vgCon, 1, 3
            
            If Not RecGriff.EOF Then
                Griffe = RecGriff!nome
            Else
                Griffe = ""
            End If
            
            RecGriff.Close
            
            'Produto
            GridProduto.Col = 1
            If RecPesq.Fields.Item(3).Value = "Armação" Then
                GridProduto.Text = RecPesq.Fields.Item(0).Value & "/" & Griffe & "/" & RecPesq.Fields.Item(2).Value & "/" & RecPesq.Fields.Item(4).Value & "/" & RecPesq.Fields.Item(5).Value & "/" & RecPesq.Fields.Item(6).Value & "/" & RecPesq.Fields.Item(7).Value & "/" & RecPesq.Fields.Item(8).Value
            ElseIf InStr(RecPesq.Fields.Item(3).Value, "Lente") <> 0 Then
                GridProduto.Text = RecPesq.Fields.Item(0).Value & "/" & RecPesq.Fields.Item(9).Value & "/" & RecPesq.Fields.Item(10).Value
            Else
                GridProduto.Text = RecPesq.Fields.Item(0).Value & "/" & RecPesq.Fields.Item(3).Value
            End If
            GridProduto.Lock = True
            
            'TipoProd
            GridProduto.Col = 2
            GridProduto.Text = VerificaNulo(RecPesq.Fields.Item(3).Value)
            GridProduto.Lock = True
            
            'CodProd
            GridProduto.Col = 3
            GridProduto.Text = Val(RecPesq.Fields.Item(0).Value)
            GridProduto.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridProduto.MaxRows = GridProduto.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         GridProduto.MaxRows = GridProduto.MaxRows - 1
    End If
    
    Desconecta
    
End Sub

