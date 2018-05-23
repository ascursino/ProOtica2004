VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmVendedor_Cons 
   Caption         =   "Consulta de Vendedor"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmVendedor_Cons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraBotaoVendedor 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   6495
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
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcluir 
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
         TabIndex        =   6
         ToolTipText     =   "Excluir vendedor"
         Top             =   240
         Width           =   1095
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
         Left            =   1560
         TabIndex        =   5
         ToolTipText     =   "Alterar vendedor"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdImprimir 
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
         Left            =   3960
         TabIndex        =   7
         ToolTipText     =   "Imprimir consulta de vendedor"
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
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "Incluir vendedor"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FraConsultarVendedor 
      Caption         =   "Consulta de Vendedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6495
      Begin VB.TextBox TxtTel 
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
         MaxLength       =   14
         TabIndex        =   1
         ToolTipText     =   "Número do telefone do vendedor"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtVend 
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
         ToolTipText     =   "Nome do vendedor"
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton CmdPesqVend 
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
         Left            =   4560
         TabIndex        =   2
         ToolTipText     =   "Pesquisar vendedor"
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   5880
         OleObjectBlob   =   "FrmVendedor_Cons.frx":0CCA
         Top             =   240
      End
      Begin FPSpread.vaSpread GridVendedor 
         Height          =   3495
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   6015
         _Version        =   393216
         _ExtentX        =   10610
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
         SpreadDesigner  =   "FrmVendedor_Cons.frx":0EFE
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalVend 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "FrmVendedor_Cons.frx":12A0
         TabIndex        =   11
         Top             =   1320
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmVendedor_Cons.frx":132E
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmVendedor_Cons.frx":1398
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmVendedor_Cons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrResponse As String
Public RecPesq As New ADODB.Recordset

Private Sub CmdAlterar_Click()
    FrmVendedor_Alt.Show
End Sub

Private Sub CmdExcluir_Click()
    VPStrResponse = MsgBox("Deseja excluir este vendedor?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Vendedor WHERE CodVendedor=" & VGIntCodVend)
        Desconecta
        
        FrmVendedor_Cons.CmdPesqVend.Value = True
    End If
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    
    Dim vend As String
    Dim tel As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridVendedor.MaxRows
        
        GridVendedor.Col = 1
        GridVendedor.Row = VLStrLinha
        vend = GridVendedor.Text
        
        GridVendedor.Col = 2
        GridVendedor.Row = VLStrLinha
        tel = GridVendedor.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02) " & _
        "VALUES ('" & vend & "','" & tel & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptVendedor.Show

End Sub

Private Sub CmdIncluir_Click()
    FrmVendedor_Inc.Show
End Sub

Private Sub CmdPesqVend_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Vendedor where 0=0"
            
    '====== PESQUISAR POR NOME DO VENDEDOR ==========
    If TxtVend.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtVend.Text & "%'"
        VLStrOrder = VLStrOrder + "Nome,"
    End If
            
    '====== PESQUISAR POR TELEFONE ==========
    If TxtTel.Text <> "" Then
        StrSql = StrSql + " and Telefone like '%" & TxtTel.Text & "%'"
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
            
    Call MontaGridVendedor
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Resize()
  FrmVendedor_Cons.Left = (MDIPrincipal.Width / 2) - (FrmVendedor_Cons.Width / 2)
  FrmVendedor_Cons.Top = (MDIPrincipal.Height / 3) - (FrmVendedor_Cons.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 6720
    Width = 6825
    Top = 1320
    Left = 4365
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdImprimir.Enabled = False
    
    MDIPrincipal.Enabled = False

End Sub

Sub MontaGridVendedor()
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalVend.Caption = "Nenhum vendedor encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridVendedor.Refresh
        GridVendedor.MaxRows = 0
        
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdImprimir.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridVendedor.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridVendedor.Row = VLIntLinha
            GridVendedor.Lock = True
            
            'Vendedor
            GridVendedor.Col = 1
            GridVendedor.Text = VerificaNulo(RecPesq.Fields.Item(1).Value)
            GridVendedor.Lock = True
            
            'Telefone
            GridVendedor.Col = 2
            GridVendedor.Text = VerificaNulo(RecPesq.Fields.Item(2).Value)
            GridVendedor.Lock = True
            
            'CodVendedor
            GridVendedor.Col = 3
            GridVendedor.Text = Val(RecPesq.Fields.Item(0).Value)
            GridVendedor.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridVendedor.MaxRows = GridVendedor.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE VENDEDORES PESQUISADOS =========
         GridVendedor.MaxRows = GridVendedor.MaxRows - 1
         
         If GridVendedor.MaxRows = 1 Then
            LblNumTotalVend.Caption = FormataNum(GridVendedor.MaxRows) & " vendedor encontrado."
         Else
            LblNumTotalVend.Caption = FormataNum(GridVendedor.MaxRows) & " vendedores encontrados."
         End If
         '================================================
         
         CmdImprimir.Enabled = True
    End If

End Sub

Private Sub GridVendedor_Click(ByVal Col As Long, ByVal Row As Long)
    GridVendedor.Row = Row
    GridVendedor.Col = 3
    If GridVendedor.Text <> "" And GridVendedor.Text <> "CodVendedor" Then
        VGIntCodVend = GridVendedor.Text
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
        CmdImprimir.Enabled = True
    Else
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdImprimir.Enabled = False
    End If
    
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub
