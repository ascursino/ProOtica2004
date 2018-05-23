VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmCaixa_APagar_Cons 
   Caption         =   "Consulta de Contas A Pagar"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000007&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCaixa_APagar_Cons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9360
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de Contas A Pagar"
      Height          =   5655
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9135
      Begin VB.OptionButton OptPagoTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         ToolTipText     =   "Todas as contas"
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton OptPagoNao 
         Caption         =   "Em aberto"
         Height          =   195
         Left            =   4560
         TabIndex        =   2
         ToolTipText     =   "Contas a pagar em aberto"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptPagoSim 
         Caption         =   "Pago"
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         ToolTipText     =   "Contas pagas"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtDescr 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   "Descrição da conta a pagar"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox TxtDtVenc2 
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         ToolTipText     =   "Maior data de vencimento da conta a pagar"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenc1 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Menor data de vencimento da conta a pagar"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton CmdPesqPag 
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
         Left            =   6960
         TabIndex        =   5
         ToolTipText     =   "Pesquisar contas a pagar"
         Top             =   480
         Width           =   1335
      End
      Begin FPSpread.vaSpread GridAPagar 
         Height          =   3735
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   8655
         _Version        =   393216
         _ExtentX        =   15266
         _ExtentY        =   6588
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
         MaxCols         =   6
         MaxRows         =   1
         OperationMode   =   2
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmCaixa_APagar_Cons.frx":0CCA
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalPag 
         Height          =   255
         Left            =   5760
         OleObjectBlob   =   "FrmCaixa_APagar_Cons.frx":113A
         TabIndex        =   15
         Top             =   1440
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_APagar_Cons.frx":11CA
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "FrmCaixa_APagar_Cons.frx":1238
         TabIndex        =   17
         Top             =   480
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_APagar_Cons.frx":1292
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame FraBotaoRec 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   9135
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmCaixa_APagar_Cons.frx":12FE
         Top             =   120
      End
      Begin VB.CommandButton CmdBaixar 
         Caption         =   "&Baixar"
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
         Left            =   6480
         TabIndex        =   11
         ToolTipText     =   "Baixa em conta a pagar"
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
         Left            =   7800
         TabIndex        =   13
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
         Left            =   3840
         TabIndex        =   9
         ToolTipText     =   "Excluir conta a pagar"
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
         Left            =   2520
         TabIndex        =   8
         ToolTipText     =   "Alterar conta a pagar"
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
         Left            =   1200
         TabIndex        =   7
         ToolTipText     =   "Incluir conta a pagar"
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
         Left            =   5160
         TabIndex        =   10
         ToolTipText     =   "Imprimir consulta de contas a pagar"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCaixa_APagar_Cons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public RecPesq As New ADODB.Recordset

Private Sub CmdAlterar_Click()
    FrmCaixa_APagar_Alt.Show
End Sub

Private Sub CmdExcluir_Click()
    
    If VGStrStatusPagto = "Pago" Then
        
        VPStrResponse = MsgBox("Este pagamento já foi efetuado, sua exclusão" & Chr(13) & "apagará todas as informações deste pagamento." & Chr(13) & "Deseja continuar?", vbYesNo, "Pró Ótica 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            vgCon.Execute ("DELETE FROM tb_ContaPagar_Pagto WHERE CodCPag=" & VGIntCodPagar)
            vgCon.Execute ("DELETE FROM tb_ContaPagar WHERE CodCPag=" & VGIntCodPagar)
            Desconecta
            
            CmdPesqPag.Value = True
            
            CmdAlterar.Enabled = False
            CmdExcluir.Enabled = False
            CmdBaixar.Enabled = False
            
        End If
    Else
        VPStrResponse = MsgBox("Deseja excluir este pagamento?", vbYesNo, "Pró Ótica 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            vgCon.Execute ("DELETE FROM tb_ContaPagar WHERE CodCPag=" & VGIntCodPagar)
            Desconecta
            
            CmdPesqPag.Value = True
            
            CmdAlterar.Enabled = False
            CmdExcluir.Enabled = False
            CmdBaixar.Enabled = False
            
        End If
    End If
    
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    
    Dim desc As String
    Dim tipo As String
    Dim venc As String
    Dim valor As String
    Dim status As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridAPagar.MaxRows
        
        GridAPagar.Col = 1
        GridAPagar.Row = VLStrLinha
        desc = GridAPagar.Text
        
        GridAPagar.Col = 2
        GridAPagar.Row = VLStrLinha
        tipo = GridAPagar.Text
        
        GridAPagar.Col = 3
        GridAPagar.Row = VLStrLinha
        venc = GridAPagar.Text
        
        GridAPagar.Col = 4
        GridAPagar.Row = VLStrLinha
        valor = GridAPagar.Text
        
        GridAPagar.Col = 5
        GridAPagar.Row = VLStrLinha
        status = GridAPagar.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & desc & "','" & tipo & "','" & venc & "','" & valor & "','" & status & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCaixa_APagar.Show

End Sub

Private Sub CmdIncluir_Click()
    FrmCaixa_APagar_Inc.Show
End Sub

Private Sub CmdBaixar_Click()
    If VGStrStatusPagto = "Pago" Then
        FrmCaixa_APagar_Baixado.Show
    Else
        FrmCaixa_APagar_Baixa.Show
    End If
End Sub

Private Sub CmdPesqPag_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    If OptPagoSim.Value = True Then
        StrSql = "Select * from tb_ContaPagar where Pago='sim'"
    ElseIf OptPagoNao.Value = True Then
        StrSql = "Select * from tb_ContaPagar where Pago='não'"
    ElseIf OptPagoTodos.Value = True Then
        StrSql = "Select * from tb_ContaPagar where 0=0"
    End If
    
    '====== PESQUISAR POR DATA DO VENCIMENTO ==========
    If (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and Vencimento >=#" & FormataDataUS(TxtDtVenc1.Text) & "# and Vencimento <= #" & FormataDataUS(TxtDtVenc2.Text) & "#"
    
    ElseIf (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text = "" Or TxtDtVenc2.Text = "__/__/____") Then
        StrSql = StrSql + " and Vencimento =#" & FormataDataUS(TxtDtVenc1.Text) & "#"
    
    ElseIf (TxtDtVenc1.Text = "" Or TxtDtVenc1.Text = "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and Vencimento =#" & FormataDataUS(TxtDtVenc2.Text) & "#"
    
    End If
            
    '====== PESQUISAR POR DESCRIÇÃO ==========
    If TxtDescr.Text <> "" Then
        StrSql = StrSql + " and Descricao like '%" & TxtDescr.Text & "%'"
    End If
            
    '====== ORDENAR PESQUISA ======================
        StrSql = StrSql + " order by Vencimento desc"
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridAPagar
        
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmCaixa_APagar_Cons.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_APagar_Cons.Width / 2)
  FrmCaixa_APagar_Cons.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_APagar_Cons.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 7215
    Width = 9480
    Top = 1020
    Left = 1890
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    TxtDtVenc1.Text = FormataData(Date)
    TxtDtVenc2.Text = FormataData(Date)
    
    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    CmdImprimir.Enabled = False
    CmdBaixar.Enabled = False
    
End Sub

Private Sub GridAPagar_Click(ByVal Col As Long, ByVal Row As Long)
    GridAPagar.Row = Row
    GridAPagar.Col = 6
    If GridAPagar.Text <> "CodCPag" And GridAPagar.Text <> "" Then
        VGIntCodPagar = GridAPagar.Text
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
        CmdBaixar.Enabled = True
    Else
        VGIntCodPagar = 0
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdBaixar.Enabled = False
    End If
    
    GridAPagar.Row = Row
    GridAPagar.Col = 5
    If GridAPagar.Text <> "Status" And GridAPagar.Text <> "" Then
        VGStrStatusPagto = GridAPagar.Text
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
        CmdBaixar.Enabled = True
    Else
        VGStrStatusPagto = ""
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdBaixar.Enabled = False
    End If
    
End Sub

Private Sub TxtDtVenc1_GotFocus()
    If TxtDtVenc1.Text = "__/__/____" Then
        TxtDtVenc1.Text = ""
    End If
End Sub

Private Sub TxtDtVenc1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
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
    If TxtDtVenc2.Text = "__/__/____" Then
        TxtDtVenc2.Text = ""
    End If
End Sub

Private Sub TxtDtVenc2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
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

Sub MontaGridAPagar()
    Dim VLIntCodValor As Double
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalPag.Caption = "Nenhum pagamento encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridAPagar.Refresh
        GridAPagar.MaxRows = 0
        
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
        CmdImprimir.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridAPagar.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridAPagar.Row = VLIntLinha
            GridAPagar.Lock = True
            
            'Descrição
            GridAPagar.Col = 1
            GridAPagar.TypeMaxEditLen = 255
            GridAPagar.Text = VerificaNulo(RecPesq.Fields.Item(4).Value)
            GridAPagar.Lock = True
            
            'Tipo
            GridAPagar.Col = 2
            GridAPagar.Text = VerificaNulo(RecPesq.Fields.Item(1).Value)
            GridAPagar.Lock = True
            
            'Vencimento
            GridAPagar.Col = 3
            GridAPagar.Text = FormataData(RecPesq.Fields.Item(2).Value)
            GridAPagar.Lock = True
            
            'Valor
            GridAPagar.Col = 4
            GridAPagar.Text = FormataMoeda(RecPesq.Fields.Item(3).Value)
            If RecPesq.Fields.Item(5).Value = "não" Then
                VLIntValor = VLIntValor + CCur(GridAPagar.Text)
            End If
            GridAPagar.Lock = True
            
            'Status
            GridAPagar.Col = 5
            If RecPesq.Fields.Item(5).Value = "sim" Then
                GridAPagar.Text = "Pago"
            ElseIf RecPesq.Fields.Item(5).Value = "não" Then
                GridAPagar.Text = "Em aberto"
            End If
            GridAPagar.Lock = True
            
            'CodCPag
            GridAPagar.Col = 6
            GridAPagar.Text = Val(RecPesq.Fields.Item(0).Value)
            GridAPagar.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridAPagar.MaxRows = GridAPagar.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         GridAPagar.Row = GridAPagar.MaxRows
         GridAPagar.Col = 1
         GridAPagar.Lock = True
         GridAPagar.Col = 2
         GridAPagar.Lock = True
         GridAPagar.Col = 3
         GridAPagar.Lock = True
         GridAPagar.Col = 4
         GridAPagar.Lock = True
         GridAPagar.Col = 5
         GridAPagar.Lock = True
         GridAPagar.Col = 6
         GridAPagar.Lock = True
         
         
         GridAPagar.MaxRows = GridAPagar.MaxRows + 1
         GridAPagar.Row = GridAPagar.MaxRows
         
         GridAPagar.Col = 1
         GridAPagar.Text = "TOTAL À PAGAR:"
         GridAPagar.Font.Bold = True
         GridAPagar.Lock = True
         GridAPagar.Col = 2
         GridAPagar.Text = FormataMoeda(VLIntValor)
         GridAPagar.Font.Bold = True
         GridAPagar.Lock = True
         GridAPagar.Col = 3
         GridAPagar.Lock = True
         GridAPagar.Col = 4
         GridAPagar.Lock = True
         GridAPagar.Col = 5
         GridAPagar.Lock = True
         GridAPagar.Col = 6
         GridAPagar.Lock = True
         
         '===== CONTAGEM DE PAGAMENTOS PESQUISADOS =========
         If (GridAPagar.MaxRows - 2) = 1 Then
            LblNumTotalPag.Caption = FormataNum((GridAPagar.MaxRows - 2)) & " pagamento encontrado."
         Else
            LblNumTotalPag.Caption = FormataNum((GridAPagar.MaxRows - 2)) & " pagamentos encontrados."
         End If
         '================================================
         
         CmdImprimir.Enabled = True
    End If

End Sub

