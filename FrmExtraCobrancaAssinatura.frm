VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmExtraCobrancaAssinatura 
   Caption         =   "Assinatura para carta de cobrança"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
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
   Icon            =   "FrmExtraCobrancaAssinatura.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5280
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
      Width           =   5055
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
         Left            =   3600
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   480
         OleObjectBlob   =   "FrmExtraCobrancaAssinatura.frx":0CCA
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
         Left            =   2160
         TabIndex        =   6
         ToolTipText     =   "Assinar carta em branco"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "&Assinar"
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
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "Assinar carta personalizada"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5055
      Begin VB.TextBox TxtWeb 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   4
         ToolTipText     =   "Endereço do site ou email da empresa"
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   3
         ToolTipText     =   "Telefone da empresa"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Bairro da empresa"
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1200
         MaxLength       =   60
         TabIndex        =   1
         ToolTipText     =   "Endereço da empresa"
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Nome da empresa"
         Top             =   240
         Width           =   3735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExtraCobrancaAssinatura.frx":0EFE
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExtraCobrancaAssinatura.frx":0F66
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExtraCobrancaAssinatura.frx":0FD0
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExtraCobrancaAssinatura.frx":1036
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExtraCobrancaAssinatura.frx":10A0
         TabIndex        =   14
         Top             =   1680
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmExtraCobrancaAssinatura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdBranco_Click()
    Unload Me
    VGStrAssinatura = "branco"
    MDIPrincipal.Enabled = True
    Call ImprimirCarta
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    MDIPrincipal.Enabled = True
End Sub

Private Sub CmdOK_Click()
    If TxtNome.Text = "" And TxtEndereco.Text = "" And TxtBairro.Text = "" And TxtTel.Text = "" And TxtWeb.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos para colocar uma assinatura personalizada." & Chr(13) & "Caso não deseje assinar a carta escolha o botão 'Em Branco'", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        Conecta
        
        Dim RecAss As New ADODB.Recordset
        
        StrSql = "Select * From tb_AssinaturaCobranca"
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
        
        VGStrAssinatura = "personalizada"
        
        Call ImprimirCarta
    End If
End Sub

Private Sub Form_Resize()
  FrmExtraCobrancaAssinatura.Left = (MDIPrincipal.Width / 2) - (FrmExtraCobrancaAssinatura.Width / 2)
  FrmExtraCobrancaAssinatura.Top = (MDIPrincipal.Height / 3) - (FrmExtraCobrancaAssinatura.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 3750
    Width = 5400
    Top = 1725
    Left = 4965
    
    MDIPrincipal.Enabled = False
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Conecta
    
    Dim RecAss As New ADODB.Recordset
    
    StrSql = "Select * From tb_AssinaturaCobranca"
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

Sub ImprimirCarta()
    If InStr(FrmPrincipal.CboTipoCarta.Text, "simples") <> 0 Then
        rptExtra_CobrancaSimples.Show
       
    ElseIf InStr(FrmPrincipal.CboTipoCarta.Text, "amigável") <> 0 Then
        rptExtra_CobrancaAmigavel.Show
    
    ElseIf InStr(FrmPrincipal.CboTipoCarta.Text, "último") <> 0 Then
        rptExtra_CobrancaUltimoAviso.Show
    
    End If
End Sub
