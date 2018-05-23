VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmFornecedor_Inc 
   Caption         =   "Inclusão de Fornecedor"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
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
   Icon            =   "FrmFornecedor_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6945
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6735
      Begin VB.TextBox TxtCep 
         Height          =   285
         Left            =   5280
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "Cep do fornecedor"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Email do fornecedor"
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox TxtCnpj 
         Height          =   285
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   7
         ToolTipText     =   "Cnpj do fornecedor"
         Top             =   2640
         Width           =   2895
      End
      Begin VB.ComboBox CboTipoForn 
         Height          =   315
         ItemData        =   "FrmFornecedor_Inc.frx":0CCA
         Left            =   1920
         List            =   "FrmFornecedor_Inc.frx":0CDA
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Tipo de fornecedor"
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   1
         ToolTipText     =   "Nome do fornecedor"
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   2
         ToolTipText     =   "Endereço do fornecedor"
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   3
         ToolTipText     =   "Bairro do fornecedor"
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   5
         ToolTipText     =   "Cidade do fornecedor"
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Estado do fornecedor"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   5280
         MaxLength       =   14
         TabIndex        =   8
         ToolTipText     =   "Número do telefone do fornecedor"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox TxtCel 
         Height          =   285
         Left            =   5280
         MaxLength       =   14
         TabIndex        =   10
         ToolTipText     =   "Número do celular do fornecedor"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Observação sobre o fornecedor"
         Top             =   4320
         Width           =   6495
      End
      Begin VB.TextBox TxtContato 
         Height          =   285
         Left            =   1320
         MaxLength       =   200
         TabIndex        =   11
         ToolTipText     =   "Nome da pessoa de contato"
         Top             =   3600
         Width           =   5295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0D08
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0D76
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0DE0
         TabIndex        =   19
         Top             =   1680
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0E46
         TabIndex        =   20
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0EAC
         TabIndex        =   21
         Top             =   2640
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0F0E
         TabIndex        =   22
         Top             =   3120
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0F72
         TabIndex        =   23
         Top             =   3600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":0FDA
         TabIndex        =   24
         Top             =   1680
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":103A
         TabIndex        =   25
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":10A0
         TabIndex        =   26
         Top             =   2640
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":110A
         TabIndex        =   27
         Top             =   3120
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":1172
         TabIndex        =   28
         Top             =   4080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":11E0
         TabIndex        =   29
         Top             =   240
         Width           =   1815
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
      TabIndex        =   15
      Top             =   5400
      Width           =   6735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmFornecedor_Inc.frx":125E
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
         Left            =   5400
         TabIndex        =   14
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
         Left            =   4080
         TabIndex        =   13
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmFornecedor_Inc"
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
    
    If VGStrIncluirProd <> "" Then
        VGStrIncluirProd = ""
        Unload Me
        FrmProduto_Inc.Enabled = True
    End If

End Sub

Private Sub CmdOK_Click()
   
    Conecta
    
    Dim RecForn As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_Fornecedor"
    RecForn.Open StrSql, vgCon, 1, 3
        
    RecForn.AddNew
    RecForn("DtForn") = FormataDataUS(Date)
    RecForn("Tipo") = CboTipoForn.Text
    RecForn("Nome") = TxtNome.Text
    RecForn("Endereco") = TxtEndereco.Text
    RecForn("Bairro") = TxtBairro.Text
    RecForn("Cep") = TxtCep.Text
    RecForn("Cidade") = TxtCidade.Text
    RecForn("Estado") = CboEstado.Text
    RecForn("CNPJ") = TxtCnpj.Text
    RecForn("Email") = TxtEmail.Text
    RecForn("Contato") = TxtContato.Text
    RecForn("Telefone") = TxtTel.Text
    RecForn("Celular") = TxtCel.Text
    RecForn("Obs") = TxtObs.Text
    RecForn.Update
    
    Desconecta
    
    VPStrBox = MsgBox("Fornecedor cadastrado.", vbInformation, "Pró Ótica 2004 - Informação")
    
    TxtNome.Text = ""
    TxtEndereco.Text = ""
    TxtBairro.Text = ""
    TxtCep.Text = ""
    TxtCnpj.Text = ""
    TxtEmail.Text = ""
    TxtContato.Text = ""
    TxtTel.Text = ""
    TxtCel.Text = ""
    TxtObs.Text = ""
    
    TxtNome.SetFocus
    
    If VGStrIncluirProd <> "" Then
        VGStrIncluirProd = ""
        Unload Me
        FrmProduto_Inc.MontaCboForn
        FrmProduto_Inc.Enabled = True
    End If
End Sub

Private Sub Form_Resize()
  FrmFornecedor_Inc.Left = (MDIPrincipal.Width / 2) - (FrmFornecedor_Inc.Width / 2)
  FrmFornecedor_Inc.Top = (MDIPrincipal.Height / 3) - (FrmFornecedor_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 6735
    Width = 7065
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    If VGStrIncluirProd <> "" Then
        FrmProduto_Inc.Enabled = False
    End If
    
    If VGStrEstoqueIncExtra <> "" Then
        CboTipoForn.Text = VGStrEstoqueIncExtra
    End If
    
    Call MontaCbos
    
    TxtCidade.Text = "Rio de Janeiro"
    CboEstado.Text = "RJ"
    
End Sub

Sub MontaCbos()
    
    '===== CboEstado ============
    CboEstado.AddItem ("")
    CboEstado.AddItem ("AC")
    CboEstado.AddItem ("AL")
    CboEstado.AddItem ("AM")
    CboEstado.AddItem ("AP")
    CboEstado.AddItem ("BA")
    CboEstado.AddItem ("CE")
    CboEstado.AddItem ("DF")
    CboEstado.AddItem ("ES")
    CboEstado.AddItem ("GO")
    CboEstado.AddItem ("MA")
    CboEstado.AddItem ("MG")
    CboEstado.AddItem ("MS")
    CboEstado.AddItem ("MT")
    CboEstado.AddItem ("PA")
    CboEstado.AddItem ("PB")
    CboEstado.AddItem ("PE")
    CboEstado.AddItem ("PI")
    CboEstado.AddItem ("PR")
    CboEstado.AddItem ("RJ")
    CboEstado.AddItem ("RN")
    CboEstado.AddItem ("RO")
    CboEstado.AddItem ("RR")
    CboEstado.AddItem ("RS")
    CboEstado.AddItem ("SC")
    CboEstado.AddItem ("SE")
    CboEstado.AddItem ("SP")
    CboEstado.AddItem ("TO")
    '============================
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCep_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCnpj_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtEmail_LostFocus()
    If TxtEmail.Text <> "" Then
        If InStr(TxtEmail.Text, "@") = 0 Then
            VPStrBox = MsgBox("Formato do email está incorreto.", vbCritical, "Pró Ótica 2004 - Erro")
            TxtEmail.SetFocus
        End If
    End If
End Sub

