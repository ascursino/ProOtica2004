VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmCaixa_AReceber_Cons 
   Caption         =   "Consulta de Contas A Receber"
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
   Icon            =   "FrmCaixa_AReceber_Cons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9360
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de Contas A Receber"
      Height          =   5655
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9135
      Begin VB.TextBox TxtCli 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         ToolTipText     =   "Nome do cliente da conta"
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton OptRecebTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   4560
         TabIndex        =   5
         ToolTipText     =   "Todas as contas"
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton OptRecebNao 
         Caption         =   "A receber"
         Height          =   195
         Left            =   4560
         TabIndex        =   3
         ToolTipText     =   "Contas a receber"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptRecebSim 
         Caption         =   "Recebido"
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         ToolTipText     =   "Contas recebidas"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenc2 
         Height          =   285
         Left            =   3000
         TabIndex        =   1
         ToolTipText     =   "Maior data de vencimento da conta"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenc1 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Menor data de vencimento da conta"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton CmdPesqReceb 
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
         TabIndex        =   6
         ToolTipText     =   "Pesquisar contas a receber"
         Top             =   480
         Width           =   1335
      End
      Begin FPSpread.vaSpread GridAReceber 
         Height          =   3735
         Left            =   240
         TabIndex        =   7
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
         SpreadDesigner  =   "FrmCaixa_AReceber_Cons.frx":0CCA
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalReceb 
         Height          =   255
         Left            =   5760
         OleObjectBlob   =   "FrmCaixa_AReceber_Cons.frx":113C
         TabIndex        =   12
         Top             =   1440
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_AReceber_Cons.frx":11D0
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "FrmCaixa_AReceber_Cons.frx":123E
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_AReceber_Cons.frx":1298
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame FraBotaoRec 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   5880
      Width           =   9135
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmCaixa_AReceber_Cons.frx":1300
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
         Left            =   7800
         TabIndex        =   9
         ToolTipText     =   "Fechar"
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
         Left            =   6480
         TabIndex        =   8
         ToolTipText     =   "Imprimir consulta de contas a receber"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCaixa_AReceber_Cons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public RecPesq As New ADODB.Recordset

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
    
    Do While VLStrLinha <= GridAReceber.MaxRows
        
        GridAReceber.Col = 1
        GridAReceber.Row = VLStrLinha
        desc = GridAReceber.Text
        
        GridAReceber.Col = 2
        GridAReceber.Row = VLStrLinha
        tipo = GridAReceber.Text
        
        GridAReceber.Col = 3
        GridAReceber.Row = VLStrLinha
        venc = GridAReceber.Text
        
        GridAReceber.Col = 4
        GridAReceber.Row = VLStrLinha
        valor = GridAReceber.Text
        
        GridAReceber.Col = 5
        GridAReceber.Row = VLStrLinha
        status = GridAReceber.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & desc & "','" & tipo & "','" & venc & "','" & valor & "','" & status & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCaixa_AReceber.Show

End Sub

Private Sub CmdPesqReceb_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    If OptRecebSim.Value = True Then
        StrSql = "Select * from tb_Crediario_Parcela as P,tb_Crediario as CR,tb_Cliente as C where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and Quitado='sim'"
    ElseIf OptRecebNao.Value = True Then
        StrSql = "Select * from tb_Crediario_Parcela as P,tb_Crediario as CR,tb_Cliente as C where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and Quitado='não'"
    ElseIf OptRecebTodos.Value = True Then
        StrSql = "Select * from tb_Crediario_Parcela as P,tb_Crediario as CR,tb_Cliente as C where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred"
    End If
    
    '====== PESQUISAR POR DATA DO VENCIMENTO ==========
    If (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and P.Vencimento >=#" & FormataDataUS(TxtDtVenc1.Text) & "# and P.Vencimento <= #" & FormataDataUS(TxtDtVenc2.Text) & "#"
    
    ElseIf (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text = "" Or TxtDtVenc2.Text = "__/__/____") Then
        StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc1.Text) & "#"
    
    ElseIf (TxtDtVenc1.Text = "" Or TxtDtVenc1.Text = "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc2.Text) & "#"
    
    End If
            
    '====== PESQUISAR POR CLIENTE ==========
    If TxtCli.Text <> "" Then
        StrSql = StrSql + " and C.Nome like '%" & TxtCli.Text & "%'"
    End If
    
    '====== ORDENAR PESQUISA ======================
        StrSql = StrSql + " order by P.Vencimento desc"
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridAReceber
        
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmCaixa_AReceber_Cons.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_AReceber_Cons.Width / 2)
  FrmCaixa_AReceber_Cons.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_AReceber_Cons.Height / 3)
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
    
    CmdImprimir.Enabled = False
    
End Sub

Private Sub GridAReceber_Click(ByVal Col As Long, ByVal Row As Long)
    GridAReceber.Row = Row
    GridAReceber.Col = 6
    If GridAReceber.Text <> "CodParc" And GridAReceber.Text <> "" Then
        VGIntCodReceber = GridAReceber.Text
    Else
        VGIntCodReceber = 0
    End If
    
    GridAReceber.Row = Row
    GridAReceber.Col = 5
    If GridAReceber.Text <> "Status" And GridAReceber.Text <> "" Then
        VGStrStatusReceb = GridAReceber.Text
    Else
        VGStrStatusReceb = ""
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

Sub MontaGridAReceber()
    Dim RecCred As New ADODB.Recordset
    Dim VLIntCodValor As Double
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalReceb.Caption = "Nenhum recebimento encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Ótica 2004 - Informação")
        GridAReceber.Refresh
        GridAReceber.MaxRows = 0
        
        CmdImprimir.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridAReceber.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
            StrSql = "SELECT C.Nome,CR.TipoCred FROM tb_Cliente as C, tb_Crediario as CR " & _
                     "WHERE C.CodCli=CR.CodCli AND CR.CodCred=" & RecPesq.Fields.Item(1).Value
            RecCred.Open StrSql, vgCon, 1, 3
            
            GridAReceber.Row = VLIntLinha
            GridAReceber.Lock = True
            
            'Descrição
            GridAReceber.Col = 1
            GridAReceber.TypeMaxEditLen = 255
            GridAReceber.Text = "Parcela de crediário - Cliente: " & VerificaNulo(RecCred.Fields.Item(0).Value)
            GridAReceber.Lock = True
            
            'Tipo
            GridAReceber.Col = 2
            GridAReceber.Text = VerificaNulo(RecCred.Fields.Item(1).Value)
            GridAReceber.Lock = True
            
            'Vencimento
            GridAReceber.Col = 3
            GridAReceber.Text = FormataData(RecPesq.Fields.Item(3).Value)
            GridAReceber.Lock = True
            
            'Valor
            GridAReceber.Col = 4
            GridAReceber.Text = FormataMoeda(RecPesq.Fields.Item(4).Value)
            If RecPesq.Fields.Item(5).Value = "não" Then
                VLIntValor = VLIntValor + CCur(GridAReceber.Text)
            End If
            GridAReceber.Lock = True
            
            'Status
            GridAReceber.Col = 5
            If RecPesq.Fields.Item(5).Value = "sim" Then
                GridAReceber.Text = "Recebido"
            ElseIf RecPesq.Fields.Item(5).Value = "não" Then
                GridAReceber.Text = "A receber"
            End If
            GridAReceber.Lock = True
            
            'CodParc
            GridAReceber.Col = 6
            GridAReceber.Text = Val(RecPesq.Fields.Item(0).Value)
            GridAReceber.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridAReceber.MaxRows = GridAReceber.MaxRows + 1
            RecCred.Close
            RecPesq.MoveNext
         Loop
         
         GridAReceber.Row = GridAReceber.MaxRows
         GridAReceber.Col = 1
         GridAReceber.Lock = True
         GridAReceber.Col = 2
         GridAReceber.Lock = True
         GridAReceber.Col = 3
         GridAReceber.Lock = True
         GridAReceber.Col = 4
         GridAReceber.Lock = True
         GridAReceber.Col = 5
         GridAReceber.Lock = True
         GridAReceber.Col = 6
         GridAReceber.Lock = True
         
         
         GridAReceber.MaxRows = GridAReceber.MaxRows + 1
         GridAReceber.Row = GridAReceber.MaxRows
         
         GridAReceber.Col = 1
         GridAReceber.Text = "TOTAL À RECEBER:"
         GridAReceber.Font.Bold = True
         GridAReceber.Lock = True
         GridAReceber.Col = 2
         GridAReceber.Text = FormataMoeda(VLIntValor)
         GridAReceber.Font.Bold = True
         GridAReceber.Lock = True
         GridAReceber.Col = 3
         GridAReceber.Lock = True
         GridAReceber.Col = 4
         GridAReceber.Lock = True
         GridAReceber.Col = 5
         GridAReceber.Lock = True
         GridAReceber.Col = 6
         GridAReceber.Lock = True
         
         '===== CONTAGEM DE RECEBIMENTOS PESQUISADOS =========
         If (GridAReceber.MaxRows - 2) = 1 Then
            LblNumTotalReceb.Caption = FormataNum((GridAReceber.MaxRows - 2)) & " recebimento encontrado."
         Else
            LblNumTotalReceb.Caption = FormataNum((GridAReceber.MaxRows - 2)) & " recebimentos encontrados."
         End If
         '================================================
         
         CmdImprimir.Enabled = True
    End If

End Sub

