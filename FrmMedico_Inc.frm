VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmMedico_Inc 
   Caption         =   "Inclusão de Médico"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
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
   Icon            =   "FrmMedico_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   6930
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
      Height          =   6135
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtTel2 
         Height          =   285
         Left            =   2760
         MaxLength       =   14
         TabIndex        =   12
         ToolTipText     =   "Número do telefone da Clínica/Consultório ou médico"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox TxtFax 
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   14
         ToolTipText     =   "Número do fax da Clínica/Consultório ou médico"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox TxtTel1 
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   11
         ToolTipText     =   "Número do telefone da Clínica/Consultório ou médico"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox TxtCpf 
         Height          =   285
         Left            =   5160
         MaxLength       =   11
         TabIndex        =   2
         ToolTipText     =   "CPF do médico"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Email do médico"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox TxtDtNasc 
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "__/__/____"
         ToolTipText     =   "Data de nascimento do médico"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox TxtCep 
         Height          =   285
         Left            =   5280
         MaxLength       =   8
         TabIndex        =   8
         ToolTipText     =   "Cep da Clínica/Consultório do médico"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox TxtCliCons 
         Height          =   285
         Left            =   1920
         MaxLength       =   200
         TabIndex        =   5
         ToolTipText     =   "Clínica/Consultório do médico"
         Top             =   1680
         Width           =   4695
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   0
         ToolTipText     =   "Nome do médico"
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   6
         ToolTipText     =   "Endereço da Clínica/Consultório do médico"
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   7
         ToolTipText     =   "Bairro da Clínica/Consultório do médico"
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   9
         ToolTipText     =   "Cidade da Clínica/Consultório do médico"
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Estado da Clínica/Consultório do médico"
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox TxtCel 
         Height          =   285
         Left            =   5280
         MaxLength       =   14
         TabIndex        =   13
         ToolTipText     =   "Número do celular do médico"
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox TxtObs 
         Height          =   1005
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Observação sobre o médico e/ou a Clínica/Consultório"
         Top             =   4920
         Width           =   6495
      End
      Begin VB.TextBox TxtCrm 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "CRM do médico"
         Top             =   720
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":0CCA
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":0D2C
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":0DAC
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":0E0C
         TabIndex        =   23
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":0E76
         TabIndex        =   24
         Top             =   2640
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":0EDC
         TabIndex        =   25
         Top             =   3120
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmMedico_Inc.frx":0F42
         TabIndex        =   26
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmMedico_Inc.frx":0FA6
         TabIndex        =   27
         Top             =   720
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmMedico_Inc.frx":1006
         TabIndex        =   28
         Top             =   2640
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmMedico_Inc.frx":1066
         TabIndex        =   29
         Top             =   3120
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":10CC
         TabIndex        =   30
         Top             =   3600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   4440
         OleObjectBlob   =   "FrmMedico_Inc.frx":1136
         TabIndex        =   31
         Top             =   3600
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":119E
         TabIndex        =   32
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":1202
         TabIndex        =   33
         Top             =   4560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMedico_Inc.frx":1270
         TabIndex        =   34
         Top             =   4080
         Width           =   855
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
      TabIndex        =   18
      Top             =   6240
      Width           =   6735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2040
         OleObjectBlob   =   "FrmMedico_Inc.frx":12D0
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
         TabIndex        =   17
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
         TabIndex        =   16
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmMedico_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub CmdFechar_Click()
    Unload Me
    
    If VGStrForm = "receita" Then
        FrmReceita_Inc.Enabled = True
    End If
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdOK_Click()
    
    Conecta
    
    Dim RecMed As New ADODB.Recordset
    Dim VLStrTel As String
    
    If TxtTel1.Text <> "" And TxtTel2.Text <> "" Then
        VLStrTel = TxtTel1.Text & "/" & TxtTel2.Text
    ElseIf TxtTel1.Text <> "" And TxtTel2.Text = "" Then
        VLStrTel = TxtTel1.Text
    ElseIf TxtTel1.Text = "" And TxtTel2.Text <> "" Then
        VLStrTel = TxtTel2.Text
    End If
    
    StrSql = "SELECT * FROM tb_Medico"
    RecMed.Open StrSql, vgCon, 1, 3
    
    RecMed.AddNew
    RecMed("Nome") = TxtNome.Text
    RecMed("CliCons") = TxtCliCons.Text
    RecMed("Crm") = TxtCrm.Text
    RecMed("Endereco") = TxtEndereco.Text
    RecMed("Bairro") = TxtBairro.Text
    RecMed("Cep") = TxtCep.Text
    RecMed("Cidade") = TxtCidade.Text
    RecMed("Estado") = CboEstado.Text
    RecMed("DtNasc") = FormataDataUS(TxtDtNasc.Text)
    RecMed("Telefone") = VLStrTel
    RecMed("Celular") = TxtCel.Text
    RecMed("Fax") = TxtFax.Text
    RecMed("Cpf") = TxtCpf.Text
    RecMed("Email") = TxtEmail.Text
    RecMed("Obs") = TxtObs.Text
    RecMed.Update
        
    RecMed.Close
    
    Desconecta
    
    VPStrBox = MsgBox("Médico cadastrado.", vbInformation, "Pró Ótica 2004 - Informação")
    
    Unload Me
    
    If VGStrForm = "receita" Then
        FrmReceita_Inc.MontaCboMedico
        FrmReceita_Inc.Enabled = True
        VGStrForm = ""
    Else
        FrmPrincipal.CmdPesqMed.Value = True
    End If
    
    'MDIPrincipal.Enabled = True
    'MDIPrincipal.WindowState = 2
End Sub

Private Sub Form_Resize()
  FrmMedico_Inc.Left = (MDIPrincipal.Width / 2) - (FrmMedico_Inc.Width / 2)
  FrmMedico_Inc.Top = (MDIPrincipal.Height / 3) - (FrmMedico_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 7590
    Width = 7050
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Call MontaCboEstado
    
    TxtCidade.Text = "Rio de Janeiro"
    CboEstado.Text = "RJ"

    MDIPrincipal.Enabled = False
    
    If VGStrForm = "receita" Then
        FrmReceita_Inc.Enabled = False
    End If
End Sub

Sub MontaCboEstado()
    
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

Private Sub TxtCpf_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCrm_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, letras minúsculas, letras maiúsculas e / - backspace e enter ===
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii < 97 Or KeyAscii > 122) And KeyAscii <> 45 And KeyAscii <> 47 And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNasc_GotFocus()
    TxtDtNasc.Text = ""
End Sub

Private Sub TxtDtNasc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If TxtDtNasc.Text = "__/__/____" Then
        TxtDtNasc.Text = ""
    End If
    
End Sub

Private Sub TxtDtNasc_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtNasc.Text <> "" Then
        VLStrData = VerificaData(TxtDtNasc.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtNasc.SetFocus
        Else
            TxtDtNasc.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtNasc.Text = "__/__/____"
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

Private Sub TxtTel1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTel2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub
