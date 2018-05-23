VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCaixa_APagar_Baixa 
   Caption         =   "Baixa em Contas a Pagar"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
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
   Icon            =   "FrmCaixa_APagar_Baixa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   5865
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
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   5655
      Begin VB.OptionButton OptPagtoDin 
         Caption         =   "Dinheiro"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         ToolTipText     =   "Pagamento em dinheiro"
         Top             =   1320
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptPagtoChq 
         Caption         =   "Cheque"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         ToolTipText     =   "Pagamento em cheque"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame FraPagtoChq 
         Height          =   855
         Left            =   1800
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox TxtDigito 
            Height          =   285
            Left            =   2520
            MaxLength       =   1
            TabIndex        =   8
            ToolTipText     =   "Dígito do número do cheque"
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox TxtCheque 
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   7
            ToolTipText     =   "Número do cheque"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TxtBanco 
            Height          =   285
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   6
            ToolTipText     =   "Número do banco do cheque"
            Top             =   120
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0CCA
            TabIndex        =   14
            Top             =   120
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0D34
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox TxtDtPagto 
         Height          =   285
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Data do pagamento da parcela"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtValorPago 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Valor pago da parcela"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Data de vencimento da parcela"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtValor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Valor da parcela"
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0DA0
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0E04
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0E72
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0EDA
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0F4C
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
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
      TabIndex        =   11
      Top             =   2760
      Width           =   5655
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmCaixa_APagar_Baixa.frx":0FBA
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
         Left            =   4320
         TabIndex        =   10
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
         Left            =   3000
         TabIndex        =   9
         ToolTipText     =   "Efetuar a quitação da parcela"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCaixa_APagar_Baixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    FrmCaixa_APagar_Cons.Enabled = True
End Sub

Private Sub CmdOK_Click()
    Conecta
    
    Dim RecBx As New ADODB.Recordset
    Dim RecPag As New ADODB.Recordset
    Dim RecDescr As New ADODB.Recordset
    Dim RecCxa As New ADODB.Recordset
    Dim TipoPagto As String
    
    If OptPagtoDin.Value = True Then
        TipoPagto = "Dinheiro"
    ElseIf OptPagtoChq.Value = True Then
        TipoPagto = "Cheque"
    End If
    
    StrSql = "SELECT * FROM tb_ContaPagar_Pagto"
    RecBx.Open StrSql, vgCon, 1, 3
    
    RecBx.AddNew
    RecBx("CodCPag") = VGIntCodPagar
    RecBx("DtPagto") = FormataDataUS(TxtDtPagto.Text)
    RecBx("ValorPago") = Mid(TxtValorPago.Text, 4)
    RecBx("TipoPagto") = TipoPagto
    
    If TipoPagto = "Cheque" Then
        RecBx("NumBanco") = TxtBanco.Text
        RecBx("NumCheque") = TxtCheque.Text & "-" & TxtDigito.Text
    Else
        RecBx("NumBanco") = 0
        RecBx("NumCheque") = ""
    End If
    RecBx.Update
    
    'atualiza informações de pagamento na tabela de contas a pagar
    StrSql = "SELECT * FROM tb_ContaPagar where CodCPag=" & VGIntCodPagar
    RecPag.Open StrSql, vgCon, 1, 3
    
    RecPag("Pago") = "sim"
    RecPag.Update
    
    VPStrResponse = MsgBox("Envia dados para o lançamento de caixa?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        'envia dados para o lançamento de caixa
        StrSql = "SELECT * FROM tb_Caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        StrSql = "SELECT Descricao FROM tb_ContaPagar where CodCPag=" & VGIntCodPagar
        RecDescr.Open StrSql, vgCon, 1, 3
    
        RecCxa.AddNew
        RecCxa("CodVenda") = 0
        RecCxa("DtMov") = FormataDataUS(TxtDtPagto.Text)
        RecCxa("TipoMov") = "Baixa em pagamento"
        RecCxa("Valor") = Mid(TxtValorPago.Text, 4)
        RecCxa("TipoValor") = "debito"
        RecCxa("Descricao") = RecDescr.Fields.Item(0).Value
        RecCxa("TipoPagto") = TipoPagto
        RecCxa.Update
    End If
    
    Desconecta
    
    VPStrBox = MsgBox("Baixa efetuada.", vbInformation, "Pró Ótica 2004 - Informação")
    
    'atualiza a grid de contas pagar
    FrmCaixa_APagar_Cons.CmdPesqPag.Value = True
    
    Unload Me
    
    FrmCaixa_APagar_Cons.Enabled = True

End Sub

Private Sub Form_Resize()
  FrmCaixa_APagar_Baixa.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_APagar_Baixa.Width / 2)
  FrmCaixa_APagar_Baixa.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_APagar_Baixa.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4095
    Width = 5985
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmCaixa_APagar_Cons.Enabled = False
    
    Conecta
    
    Dim RecBx As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_ContaPagar WHERE CodCPag=" & VGIntCodPagar
    RecBx.Open StrSql, vgCon, 1, 3
    
    TxtValor.Text = FormataMoeda(RecBx.Fields.Item(3).Value)
    TxtDtVenc.Text = FormataData(RecBx.Fields.Item(2).Value)
    TxtValorPago.Text = FormataMoeda(RecBx.Fields.Item(3).Value)
    TxtDtPagto.Text = FormataData(Date)
    
    Desconecta
    
End Sub

Private Sub OptPagtoChq_Click()
    FraPagtoChq.Visible = True
End Sub

Private Sub OptPagtoDin_Click()
    FraPagtoChq.Visible = False
End Sub

Private Sub TxtDtPagto_GotFocus()
    If TxtDtPagto.Text = "__/__/____" Then
        TxtDtPagto.Text = ""
    End If
End Sub

Private Sub TxtDtPagto_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtPagto_LostFocus()
    Dim VLStrData As String
    
    If TxtDtPagto.Text <> "" Then
        VLStrData = VerificaData(TxtDtPagto.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtPagto.SetFocus
        Else
            TxtDtPagto.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtPagto.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenc_GotFocus()
    If TxtDtVenc.Text = "__/__/____" Then
        TxtDtVenc.Text = ""
    End If
End Sub

Private Sub TxtDtVenc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc_LostFocus()
    Dim VLStrData As String
    
    If TxtDtVenc.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenc.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtVenc.SetFocus
        Else
            TxtDtVenc.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtVenc.Text = "__/__/____"
    End If
End Sub

Private Sub TxtValor_LostFocus()
    TxtValor.Text = FormataMoeda(TxtValor.Text)
End Sub

Private Sub TxtValorPago_LostFocus()
    TxtValorPago.Text = FormataMoeda(TxtValorPago.Text)
End Sub
