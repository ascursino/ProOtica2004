VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmAssinaturaOrc 
   Caption         =   "Personalização de orçamento e proposta de crédito"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
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
   Icon            =   "FrmAssinaturaOrc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6480
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
      TabIndex        =   9
      Top             =   2400
      Width           =   6255
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
         Left            =   4680
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0CCA
         Top             =   120
      End
      Begin VB.CommandButton CmdBranco 
         Caption         =   "&Em branco"
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
         Left            =   3120
         TabIndex        =   6
         ToolTipText     =   "Personalizar orçamento em branco"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Personalizar"
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
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Personalizar carnê"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6255
      Begin VB.TextBox TxtWeb 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   4
         ToolTipText     =   "Endereço do site o email da empresa"
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   0
         ToolTipText     =   "Nome da empresa"
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   3
         ToolTipText     =   "Telefone da empresa"
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Bairro da empresa"
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   1
         ToolTipText     =   "Endereço da empresa"
         Top             =   600
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0EFE
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0F68
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":0FCE
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":1038
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAssinaturaOrc.frx":10A0
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmAssinaturaOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdBranco_Click()
    Unload Me
    If VGStrAssinaturaProposta = "proposta" Or VGStrAssinaturaProposta = "extraproposta" Then
        VGStrAssinaturaProp = "branco"
    Else
        VGStrAssinaturaOrc = "branco"
    End If
    
    MDIPrincipal.Enabled = True
    
    If VGStrAssinaturaProposta = "proposta" Then
        VGStrAssinaturaProposta = ""
        FrmVenda_Inc.MontaImpressaoProposta
    
    ElseIf VGStrAssinaturaProposta = "extraproposta" Then
        VGStrAssinaturaProposta = ""
        Call ImprimirExtraProp
    
    Else
        Call ImprimirOrc
    End If
End Sub

Private Sub CmdFechar_Click()
    VGStrAssinaturaProposta = ""
    Unload Me
    MDIPrincipal.Enabled = True
End Sub

Private Sub CmdOK_Click()
    If TxtNome.Text = "" And TxtEndereco.Text = "" And TxtBairro.Text = "" And TxtTel.Text = "" And TxtWeb.Text = "" Then
        If VGStrAssinaturaProposta = "proposta" Or VGStrAssinaturaProposta = "extraproposta" Then
            VPStrBox = MsgBox("Preencha os campos para personalizar a proposta de crédito." & Chr(13) & "Caso não deseje personalizar escolha o botão 'Em Branco'", vbInformation, "Pró Ótica 2004 - Informação")
        Else
            VPStrBox = MsgBox("Preencha os campos para personalizar o orçamento." & Chr(13) & "Caso não deseje personalizar escolha o botão 'Em Branco'", vbInformation, "Pró Ótica 2004 - Informação")
        End If
    Else
        Conecta
        
        Dim RecAss As New ADODB.Recordset
        
        If VGStrAssinaturaProposta = "proposta" Or VGStrAssinaturaProposta = "extraproposta" Then
            StrSql = "Select * From tb_AssinaturaProp"
        Else
            StrSql = "Select * From tb_AssinaturaOrc"
        End If
        
        RecAss.Open StrSql, vgCon, 1, 3
        
        If RecAss.EOF Then
            RecAss.AddNew
            RecAss("Nome") = TxtNome.Text
            RecAss("Endereco") = TxtEndereco.Text
            RecAss("Bairro") = TxtBairro.Text
            RecAss("Telefone") = TxtTel.Text
            RecAss("Web") = TxtWeb.Text
            RecAss.Update
        Else
            RecAss("Nome") = TxtNome.Text
            RecAss("Endereco") = TxtEndereco.Text
            RecAss("Bairro") = TxtBairro.Text
            RecAss("Telefone") = TxtTel.Text
            RecAss("Web") = TxtWeb.Text
            RecAss.Update
        End If
        
        Desconecta
        
        MDIPrincipal.Enabled = True
        
        Unload Me
        
        If VGStrAssinaturaProposta = "proposta" Then
            VGStrAssinaturaProp = "personalizada"
            VGStrAssinaturaProposta = ""
            FrmVenda_Inc.MontaImpressaoProposta
        
        ElseIf VGStrAssinaturaProposta = "extraproposta" Then
            VGStrAssinaturaProp = "personalizada"
            VGStrAssinaturaProposta = ""
            Call ImprimirExtraProp
        
        Else
            VGStrAssinaturaOrc = "personalizada"
            Call ImprimirOrc
        End If
    End If
End Sub

Private Sub Form_Resize()
  FrmAssinaturaOrc.Left = (MDIPrincipal.Width / 2) - (FrmAssinaturaOrc.Width / 2)
  FrmAssinaturaOrc.Top = (MDIPrincipal.Height / 3) - (FrmAssinaturaOrc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 3735
    Width = 6600
    Top = 1725
    Left = 4965
    
    MDIPrincipal.Enabled = False
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Conecta
    
    Dim RecAss As New ADODB.Recordset
    
    If VGStrAssinaturaProposta = "proposta" Or VGStrAssinaturaProposta = "extraproposta" Then
        Me.Caption = "Personalização da proposta de crédito"
        StrSql = "Select * From tb_AssinaturaProp"
    
    Else
        Me.Caption = "Personalização de orçamento"
        StrSql = "Select * From tb_AssinaturaOrc"
    End If
    
    RecAss.Open StrSql, vgCon, 1, 3
    
    If Not RecAss.EOF Then
        If IsNull(RecAss!nome) = False Then
            TxtNome.Text = RecAss!nome
        End If
        
        If IsNull(RecAss!endereco) = False Then
            TxtEndereco.Text = RecAss!endereco
        End If
        
        If IsNull(RecAss!bairro) = False Then
            TxtBairro.Text = RecAss!bairro
        End If
        
        If IsNull(RecAss!telefone) = False Then
            TxtTel.Text = RecAss!telefone
        End If
        
        If IsNull(RecAss!web) = False Then
            TxtWeb.Text = RecAss!web
        End If
    End If
    
    Desconecta
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Sub ImprimirOrc()
    Screen.MousePointer = vbHourglass

    Dim data As String
    Dim vendedor As String
    Dim Armacao As String
    Dim valorarm As String
    Dim lente As String
    Dim valorlente As String
    Dim lentec As String
    Dim valorlentec As String
    Dim outro As String
    Dim valoroutro As String
    Dim totalvenda As String
    Dim parcelado As String
    Dim entrada As String
    Dim valorparc As String
    Dim valorprazo As String
    Dim validade As String
    Dim obs As String
    Dim cliente As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= FrmPrincipal.GridOrcamento.MaxRows

        FrmPrincipal.GridOrcamento.Col = 1
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        data = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 2
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        vendedor = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 3
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        cliente = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 5
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        Armacao = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 6
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        valorarm = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 7
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        lente = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 8
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        valorlente = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 9
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        lentec = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 10
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        valorlentec = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 11
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        outro = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 12
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        valoroutro = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 13
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        totalvenda = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 14
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        parcelado = Mid(FrmPrincipal.GridOrcamento.Text, 1, 2)

        FrmPrincipal.GridOrcamento.Col = 15
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        entrada = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 16
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        valorparc = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 17
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        valorprazo = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 18
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        validade = FrmPrincipal.GridOrcamento.Text

        FrmPrincipal.GridOrcamento.Col = 19
        FrmPrincipal.GridOrcamento.Row = VLStrLinha
        obs = FrmPrincipal.GridOrcamento.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18) " & _
        "VALUES ('" & data & "','" & vendedor & "','" & Armacao & "','" & valorarm & "','" & lente & "','" & valorlente & "','" & lentec & "','" & valorlentec & "','" & outro & "','" & valoroutro & "','" & totalvenda & "','" & parcelado & "','" & entrada & "','" & valorparc & "','" & valorprazo & "','" & validade & "','" & obs & "','" & cliente & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptOrcamento.Show
End Sub

Sub ImprimirExtraProp()
    Dim RecVenda As New ADODB.Recordset
    Dim RecCred As New ADODB.Recordset
    Dim RecCredParc As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    Dim RecCredsta As New ADODB.Recordset
    Dim RecMed As New ADODB.Recordset
    Dim RecRec As New ADODB.Recordset
    Dim RecAux As New ADODB.Recordset
    Dim RecVerif As New ADODB.Recordset
    Dim VLStrNomeMed As String
    Dim VLStrCRMMed As String
    Dim VLStrCPFMed As String
    Dim parctemp As Integer
    
    Conecta
    
    '=== Pega código da venda =======
    StrSql = "Select CodVenda From tb_Venda where CodCred=" & VGIntPropCodCred
    RecVenda.Open StrSql, vgCon, 1, 3
    
    VGIntCodVendaRel = RecVenda!codvenda
    
    '=== Pega informações do crediário =======
    StrSql = "Select CodCredsta,CodCli,DtCred,TipoCred,ValorVenda,Parcela,Juros,ValorTotal,TipoEntr,ValorEntr " & _
             "From tb_Crediario where CodCred=" & VGIntPropCodCred
    RecCred.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações das parcelas crediário =======
    StrSql = "Select Vencimento,Valor From tb_Crediario_Parcela where CodCred=" & VGIntPropCodCred
    RecCredParc.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do crediarista =======
    StrSql = "Select Nome,Endereco,Bairro,Cep,Cidade,Estado,DtNasc,Telefone,CPF " & _
             "From tb_Crediarista where CodCredsta=" & RecCred!CodCredsta
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do cliente =======
    StrSql = "Select Nome,Endereco,Bairro,Cep,Cidade,Estado,DtNasc,Telefone,CPF " & _
             "From tb_Cliente where CodCli=" & RecCred!CodCli
    RecCli.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do médico =======
    StrSql = "Select CodMed From tb_Receita where CodCli=" & RecCred!CodCli
    RecRec.Open StrSql, vgCon, 1, 3

    If Not RecRec.EOF Then
        StrSql = "Select Nome,CRM,Cpf From tb_Medico where CodMed=" & RecRec!CodMed
        RecMed.Open StrSql, vgCon, 1, 3
        VLStrNomeMed = RecMed!nome
        VLStrCRMMed = RecMed!crm
        VLStrCPFMed = RecMed!cpf
    Else
        VLStrNomeMed = ""
        VLStrCRMMed = ""
        VLStrCPFMed = ""
    End If
    
    '=== Insere informações na tabela auxiliar =======
    StrSql = "Select * From tb_Auxiliar"
    RecAux.Open StrSql, vgCon, 1, 3
    
    RecAux.AddNew
    RecAux("Campo01") = RecCredsta!nome
    RecAux("Campo02") = FormataData(RecCredsta!dtnasc)
    RecAux("Campo03") = RecCredsta!cpf
    RecAux("Campo04") = RecCredsta!telefone
    RecAux("Campo05") = RecCredsta!endereco
    RecAux("Campo06") = RecCredsta!bairro
    RecAux("Campo07") = RecCredsta!cidade & "/" & RecCredsta!Estado
    RecAux("Campo08") = RecCredsta!cep
    RecAux("Campo09") = RecCli!nome
    RecAux("Campo10") = FormataData(RecCli!dtnasc)
    RecAux("Campo11") = RecCli!cpf
    RecAux("Campo12") = RecCli!telefone
    RecAux("Campo13") = RecCli!endereco
    RecAux("Campo14") = RecCli!bairro
    RecAux("Campo15") = RecCli!cidade & "/" & RecCli!Estado
    RecAux("Campo16") = RecCli!cep
    RecAux("Campo17") = VLStrNomeMed
    RecAux("Campo18") = VLStrCRMMed
    RecAux("Campo19") = VLStrCPFMed
    RecAux("Campo20") = FormataData(RecCred!dtcred)
    RecAux("Campo21") = RecCred!tipocred
    RecAux("Campo22") = FormataMoeda(RecCred!valorvenda)
    If RecCred!juros = "" Then
        RecAux("Campo23") = ""
    Else
        RecAux("Campo23") = FormataNum(RecCred!juros) & "%"
    End If
    RecAux("Campo24") = FormataMoeda(RecCred!valortotal)
    RecAux("Campo25") = FormataNum(RecCred!parcela)
    RecAux("Campo26") = RecCred!tipoentr
    If RecCred!valorentr = "" Then
        RecAux("Campo27") = ""
    Else
        RecAux("Campo27") = FormataMoeda(RecCred!valorentr)
    End If
    
    parctemp = 1
    
    Do While parctemp <= RecCredParc.RecordCount
        
        If parctemp = 1 Then
            RecAux("Campo28") = FormataData(RecCredParc!vencimento)
            RecAux("Campo29") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 2 Then
            RecAux("Campo30") = FormataData(RecCredParc!vencimento)
            RecAux("Campo31") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 3 Then
            RecAux("Campo32") = FormataData(RecCredParc!vencimento)
            RecAux("Campo33") = FormataMoeda(RecCredParc!valor)
                
        ElseIf parctemp = 4 Then
            RecAux("Campo34") = FormataData(RecCredParc!vencimento)
            RecAux("Campo35") = FormataMoeda(RecCredParc!valor)
                
        ElseIf parctemp = 5 Then
            RecAux("Campo36") = FormataData(RecCredParc!vencimento)
            RecAux("Campo37") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 6 Then
            RecAux("Campo38") = FormataData(RecCredParc!vencimento)
            RecAux("Campo39") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 7 Then
            RecAux("Campo40") = FormataData(RecCredParc!vencimento)
            RecAux("Campo41") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 8 Then
            RecAux("Campo42") = FormataData(RecCredParc!vencimento)
            RecAux("Campo43") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 9 Then
            RecAux("Campo44") = FormataData(RecCredParc!vencimento)
            RecAux("Campo45") = FormataMoeda(RecCredParc!valor)
        
        ElseIf parctemp = 10 Then
            RecAux("Campo46") = FormataData(RecCredParc!vencimento)
            RecAux("Campo47") = FormataMoeda(RecCredParc!valor)
        End If
        
        parctemp = parctemp + 1
        RecCredParc.MoveNext
    Loop
    
    RecAux.Update
    
    Desconecta
    
    rptPropCredito.Show
End Sub
