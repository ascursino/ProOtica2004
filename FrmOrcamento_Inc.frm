VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmOrcamento_Inc 
   Caption         =   "Inclusão de Orçamento"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
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
   Icon            =   "FrmOrcamento_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   8280
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
      Height          =   7095
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   8055
      Begin VB.TextBox TxtTotalPrazo 
         Height          =   285
         Left            =   5280
         MaxLength       =   13
         TabIndex        =   11
         ToolTipText     =   "Total da venda a prazo"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox TxtValorParc 
         Height          =   285
         Left            =   5280
         MaxLength       =   13
         TabIndex        =   14
         ToolTipText     =   "Valor das parcelas"
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox TxtEntrada 
         Height          =   285
         Left            =   5280
         MaxLength       =   13
         TabIndex        =   13
         ToolTipText     =   "Valor da entrada"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox TxtTotalVista 
         Height          =   285
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   10
         ToolTipText     =   "Total da venda à vista"
         Top             =   3480
         Width           =   1575
      End
      Begin VB.ComboBox CboVendedor 
         Height          =   315
         ItemData        =   "FrmOrcamento_Inc.frx":0CCA
         Left            =   1200
         List            =   "FrmOrcamento_Inc.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   16
         ToolTipText     =   "Nome do vendedor"
         Top             =   5520
         Width           =   6615
      End
      Begin VB.ComboBox CboQtdeParc 
         Height          =   315
         ItemData        =   "FrmOrcamento_Inc.frx":0CCE
         Left            =   5280
         List            =   "FrmOrcamento_Inc.frx":0CD0
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Quantidade de parcelas"
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Outro"
         Height          =   1215
         Left            =   4080
         TabIndex        =   31
         Top             =   2040
         Width           =   3855
         Begin VB.TextBox TxtValorOutro 
            Height          =   285
            Left            =   1080
            MaxLength       =   13
            TabIndex        =   9
            ToolTipText     =   "Valor do produto extra"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox TxtDescrOutro 
            Height          =   285
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   8
            ToolTipText     =   "Descrição de produto extra"
            Top             =   360
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0CD2
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0D3E
            TabIndex        =   33
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Lente de contato"
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   3855
         Begin VB.TextBox TxtValorLenteC 
            Height          =   285
            Left            =   1080
            MaxLength       =   13
            TabIndex        =   7
            ToolTipText     =   "Valor da lente de contato"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox TxtDescrLenteC 
            Height          =   285
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   6
            ToolTipText     =   "Descrição da lente de contato"
            Top             =   360
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0DA2
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0E0E
            TabIndex        =   30
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Lente"
         Height          =   1215
         Left            =   4080
         TabIndex        =   25
         Top             =   720
         Width           =   3855
         Begin VB.TextBox TxtValorLente 
            Height          =   285
            Left            =   1080
            MaxLength       =   13
            TabIndex        =   5
            ToolTipText     =   "Valor da lente"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox TxtDescrLente 
            Height          =   285
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   4
            ToolTipText     =   "Descrição da lente"
            Top             =   360
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0E72
            TabIndex        =   26
            Top             =   360
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0EDE
            TabIndex        =   27
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Observação sobre o cliente e/ou orçamento"
         Top             =   6240
         Width           =   7695
      End
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   6240
         MaxLength       =   14
         TabIndex        =   1
         ToolTipText     =   "Número do telefone do cliente"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Armação"
         Height          =   1215
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   3855
         Begin VB.TextBox TxtValorArm 
            Height          =   285
            Left            =   1080
            MaxLength       =   13
            TabIndex        =   3
            ToolTipText     =   "Valor da armação"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox TxtDescrArm 
            Height          =   285
            Left            =   1080
            MaxLength       =   200
            TabIndex        =   2
            ToolTipText     =   "Descrição  da armação"
            Top             =   360
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0F42
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmOrcamento_Inc.frx":0FAE
            TabIndex        =   24
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.TextBox TxtCli 
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   0
         ToolTipText     =   "Nome do cliente"
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox TxtValidade 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   15
         ToolTipText     =   "Data da validade do orçamento"
         Top             =   5040
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":1012
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":107A
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":10E4
         TabIndex        =   36
         Top             =   3480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":1158
         TabIndex        =   37
         Top             =   3960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":11D2
         TabIndex        =   38
         Top             =   4440
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":123A
         TabIndex        =   39
         Top             =   4920
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":12B8
         TabIndex        =   40
         Top             =   3480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":132C
         TabIndex        =   41
         Top             =   5040
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":139A
         TabIndex        =   42
         Top             =   5520
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":1404
         TabIndex        =   43
         Top             =   6000
         Width           =   1215
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
      TabIndex        =   20
      Top             =   7200
      Width           =   8055
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2760
         OleObjectBlob   =   "FrmOrcamento_Inc.frx":1472
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
         Left            =   6720
         TabIndex        =   19
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
         Left            =   5520
         TabIndex        =   18
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmOrcamento_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public VPIntValorArm As Currency
Public VPIntValorLente As Currency
Public VPIntValorLenteC As Currency
Public VPIntValorOutro As Currency
Public VPIntEntrada As Currency

Private Sub CboQtdeParc_Click()
    Dim VLIntRestante As Currency
    
    If CboQtdeParc.Text <> "" And CboQtdeParc.Text <> "00" Then
        TxtEntrada.Text = FormataMoeda((CCur(TxtTotalPrazo.Text) * 20) / 100)
        VLIntRestante = CCur(TxtTotalPrazo.Text) - CCur(TxtEntrada.Text)
        
        TxtValorParc.Text = FormataMoeda(VLIntRestante / Int(CboQtdeParc.Text))
    End If
End Sub

Private Sub CboQtdeParc_LostFocus()
    Dim VLIntRestante As Currency
    
    If CboQtdeParc.Text <> "" And CboQtdeParc.Text <> "00" Then
        TxtEntrada.Text = FormataMoeda((CCur(TxtTotalPrazo.Text) * 20) / 100)
        VLIntRestante = CCur(TxtTotalPrazo.Text) - CCur(TxtEntrada.Text)
        
        TxtValorParc.Text = FormataMoeda(VLIntRestante / Int(CboQtdeParc.Text))
    End If

End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdOK_Click()
    
    Conecta
    
    Dim RecOrc As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_Orcamento"
    RecOrc.Open StrSql, vgCon, 1, 3
        
    RecOrc.AddNew
    RecOrc("CodVendedor") = Trim(Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10))
    RecOrc("DtOrc") = FormataDataUS(Date)
    RecOrc("Nome") = TxtCli.Text
    RecOrc("Telefone") = TxtTel.Text
    RecOrc("DescrArm") = TxtDescrArm.Text
    RecOrc("ValorArm") = Mid(TxtValorArm.Text, 4)
    RecOrc("DescrLente") = TxtDescrLente.Text
    RecOrc("ValorLente") = Mid(TxtValorLente.Text, 4)
    RecOrc("DescrLenteC") = TxtDescrLenteC.Text
    RecOrc("ValorLenteC") = Mid(TxtValorLenteC.Text, 4)
    RecOrc("DescrOutro") = TxtDescrOutro.Text
    RecOrc("ValorOutro") = Mid(TxtValorOutro.Text, 4)
    RecOrc("TotalVenda") = Mid(TxtTotalVista.Text, 4)
    If CboQtdeParc.Text = "" Then
        RecOrc("Parcela") = 0
    Else
        RecOrc("Parcela") = CboQtdeParc.Text
    End If
    RecOrc("Juros") = "0"
    RecOrc("Entrada") = Mid(TxtEntrada.Text, 4)
    RecOrc("ValorParc") = Mid(TxtValorParc.Text, 4)
    RecOrc("ValorPrazo") = Mid(TxtTotalPrazo.Text, 4)
    RecOrc("Validade") = FormataDataUS(TxtValidade.Text)
    RecOrc("Obs") = TxtObs.Text
    RecOrc.Update
        
    RecOrc.Close
    
    Desconecta
    
    FrmPrincipal.CmdPesqOrc.Value = True
    
    VPStrResponse = MsgBox("Orçamento cadastrado." & Chr(13) & "Deseja imprimir agora?", vbYesNo, "Pró Ótica 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18) " & _
        "VALUES ('" & FormataData(Date) & "','" & Trim(Mid(CboVendedor.Text, 1, Len(CboVendedor.Text) - 10)) & "'," & _
        "'" & TxtDescrArm.Text & "','" & TxtValorArm.Text & "','" & TxtDescrLente.Text & "'," & _
        "'" & TxtValorLente.Text & "','" & TxtDescrLenteC.Text & "','" & TxtValorLenteC.Text & "'," & _
        "'" & TxtDescrOutro.Text & "','" & TxtValorOutro.Text & "','" & TxtTotalVista.Text & "'," & _
        "'" & CboQtdeParc.Text & "','" & TxtEntrada.Text & "','" & TxtValorParc.Text & "'," & _
        "'" & TxtTotalPrazo.Text & "','" & TxtValidade.Text & "','" & TxtObs.Text & "','" & TxtCli.Text & "')"
        Desconecta
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
        
        rptOrcamento.Show
    Else
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
    End If
        
End Sub

Private Sub Form_Resize()
  FrmOrcamento_Inc.Left = (MDIPrincipal.Width / 2) - (FrmOrcamento_Inc.Width / 2)
  FrmOrcamento_Inc.Top = (MDIPrincipal.Height / 3) - (FrmOrcamento_Inc.Height / 3)
End Sub

Private Sub Form_Load()
   Height = 8535
    Width = 8400
    Top = 345
    Left = 2775
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    Call MontaCbos
    
    TxtValidade.Text = "__/__/____"
End Sub

Sub MontaCbos()
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
    
    CboQtdeParc.AddItem ("00")
    CboQtdeParc.AddItem ("01")
    CboQtdeParc.AddItem ("02")
    CboQtdeParc.AddItem ("03")
    CboQtdeParc.AddItem ("04")
    CboQtdeParc.AddItem ("05")
    CboQtdeParc.AddItem ("06")
    CboQtdeParc.AddItem ("07")
    CboQtdeParc.AddItem ("08")
    CboQtdeParc.AddItem ("09")
    CboQtdeParc.AddItem ("10")
End Sub

Private Sub TxtEntrada_LostFocus()
    If TxtEntrada.Text <> "" Then
        If TxtTotalPrazo.Text <> "" And CboQtdeParc.Text <> "" Then
            VLIntRestante = CCur(TxtTotalPrazo.Text) - CCur(TxtEntrada.Text)
            TxtValorParc.Text = FormataMoeda(VLIntRestante / Int(CboQtdeParc.Text))
            TxtEntrada.Text = FormataMoeda(TxtEntrada.Text)
        End If
    Else
        If TxtTotalPrazo.Text <> "" And CboQtdeParc.Text <> "" Then
            TxtValorParc.Text = FormataMoeda(CCur(TxtTotalPrazo.Text) / Int(CboQtdeParc.Text))
        End If
    End If
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTotalPrazo_GotFocus()
    If TxtTotalVista.Text <> "" Then
        TxtTotalPrazo.Text = TxtTotalVista.Text
    End If
End Sub

Private Sub TxtTotalPrazo_LostFocus()
    If TxtTotalPrazo.Text <> "" Then
        TxtTotalPrazo.Text = FormataMoeda(TxtTotalPrazo.Text)
    End If
End Sub

Private Sub TxtTotalVista_LostFocus()
    If TxtTotalVista.Text <> "" Then
        TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
    End If
End Sub

Private Sub TxtValidade_GotFocus()
    If TxtValidade.Text = "__/__/____" Then
        TxtValidade.Text = ""
    End If
End Sub

Private Sub TxtValidade_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtValidade_LostFocus()
    Dim VLStrData As String
    
    If TxtValidade.Text <> "" Then
        VLStrData = VerificaData(TxtValidade.Text)
        
        If VGStrDataErro = "sim" Then
            TxtValidade.SetFocus
        Else
            TxtValidade.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtValidade.Text = "__/__/____"
    End If
End Sub

Private Sub TxtValorArm_GotFocus()
    If TxtValorArm.Text <> "" Then
        VPIntValorArm = TxtValorArm.Text
    Else
        VPIntValorArm = "0,00"
    End If
End Sub

Private Sub TxtValorArm_LostFocus()
    If TxtValorArm.Text <> "" Then
        If VPIntValorArm = "0,00" Then
            If TxtTotalVista.Text = "" Then
                TxtTotalVista.Text = "0,00"
            End If
        
            TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorArm.Text)
            
            TxtValorArm.Text = FormataMoeda(TxtValorArm.Text)
            TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
        Else
            If VPIntValorArm <> TxtValorArm.Text Then
                If TxtTotalVista.Text = "" Then
                    TxtTotalVista.Text = "0,00"
                End If
            
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorArm)
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorArm.Text)
                
                TxtValorArm.Text = FormataMoeda(TxtValorArm.Text)
                TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
            Else
                TxtValorArm.Text = FormataMoeda(TxtValorArm.Text)
            End If
        End If
    Else
        If TxtTotalVista.Text = "" Then
            TxtTotalVista.Text = "0,00"
        End If
        
        TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorArm)
        
        TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
    End If
End Sub

Private Sub TxtValorLente_GotFocus()
    If TxtValorLente.Text <> "" Then
        VPIntValorLente = TxtValorLente.Text
    Else
        VPIntValorLente = "0,00"
    End If
End Sub

Private Sub TxtValorLente_LostFocus()
    If TxtValorLente.Text <> "" Then
        If VPIntValorLente = "0,00" Then
            If TxtTotalVista.Text = "" Then
                TxtTotalVista.Text = "0,00"
            End If
        
            TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorLente.Text)
            
            TxtValorLente.Text = FormataMoeda(TxtValorLente.Text)
            TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
        Else
            If VPIntValorLente <> TxtValorLente.Text Then
                If TxtTotalVista.Text = "" Then
                    TxtTotalVista.Text = "0,00"
                End If
            
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorLente)
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorLente.Text)
                
                TxtValorLente.Text = FormataMoeda(TxtValorLente.Text)
                TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
            Else
                TxtValorLente.Text = FormataMoeda(TxtValorLente.Text)
            End If
        End If
    Else
        If TxtTotalVista.Text = "" Then
            TxtTotalVista.Text = "0,00"
        End If
        
        TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorLente)
        
        TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
    End If
End Sub

Private Sub TxtValorLenteC_GotFocus()
    If TxtValorLenteC.Text <> "" Then
        VPIntValorLenteC = TxtValorLenteC.Text
    Else
        VPIntValorLenteC = "0,00"
    End If
End Sub

Private Sub TxtValorLenteC_LostFocus()
    If TxtValorLenteC.Text <> "" Then
        If VPIntValorLenteC = "0,00" Then
            If TxtTotalVista.Text = "" Then
                TxtTotalVista.Text = "0,00"
            End If
        
            TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorLenteC.Text)
            
            TxtValorLenteC.Text = FormataMoeda(TxtValorLenteC.Text)
            TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
        Else
            If VPIntValorLenteC <> TxtValorLenteC.Text Then
                If TxtTotalVista.Text = "" Then
                    TxtTotalVista.Text = "0,00"
                End If
            
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorLenteC)
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorLenteC.Text)
                
                TxtValorLenteC.Text = FormataMoeda(TxtValorLenteC.Text)
                TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
            Else
                TxtValorLenteC.Text = FormataMoeda(TxtValorLenteC.Text)
            End If
        End If
    Else
        If TxtTotalVista.Text = "" Then
            TxtTotalVista.Text = "0,00"
        End If
        
        TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorLenteC)
        
        TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
    End If
End Sub

Private Sub TxtValorOutro_GotFocus()
    If TxtValorOutro.Text <> "" Then
        VPIntValorOutro = TxtValorOutro.Text
    Else
        VPIntValorOutro = "0,00"
    End If
End Sub

Private Sub TxtValorOutro_LostFocus()
    If TxtValorOutro.Text <> "" Then
        If VPIntValorOutro = "0,00" Then
            If TxtTotalVista.Text = "" Then
                TxtTotalVista.Text = "0,00"
            End If
        
            TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorOutro.Text)
            
            TxtValorOutro.Text = FormataMoeda(TxtValorOutro.Text)
            TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
        Else
            If VPIntValorOutro <> TxtValorOutro.Text Then
                If TxtTotalVista.Text = "" Then
                    TxtTotalVista.Text = "0,00"
                End If
            
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorOutro)
                TxtTotalVista.Text = CCur(TxtTotalVista.Text) + CCur(TxtValorOutro.Text)
                
                TxtValorOutro.Text = FormataMoeda(TxtValorOutro.Text)
                TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
            Else
                TxtValorOutro.Text = FormataMoeda(TxtValorOutro.Text)
            End If
        End If
    Else
        If TxtTotalVista.Text = "" Then
            TxtTotalVista.Text = "0,00"
        End If
        
        TxtTotalVista.Text = CCur(TxtTotalVista.Text) - CCur(VPIntValorOutro)
        
        TxtTotalVista.Text = FormataMoeda(TxtTotalVista.Text)
    End If
End Sub

Private Sub TxtValorParc_LostFocus()
    If TxtValorParc.Text <> "" Then
        If TxtTotalPrazo.Text <> "" And CboQtdeParc.Text <> "" Then
            TxtEntrada.Text = FormataMoeda(CCur(TxtTotalPrazo.Text) - (CCur(TxtValorParc.Text) * Int(CboQtdeParc.Text)))
            TxtValorParc.Text = FormataMoeda(TxtValorParc.Text)
        End If
    End If
End Sub
