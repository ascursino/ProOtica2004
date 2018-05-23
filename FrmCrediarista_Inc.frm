VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCrediarista_Inc 
   Caption         =   "Inclusão de Crediarista"
   ClientHeight    =   5640
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
   Icon            =   "FrmCrediarista_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   6945
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
      Height          =   4695
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   6735
      Begin VB.TextBox TxtDtNasc 
         Height          =   285
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "__/__/____"
         ToolTipText     =   "Data do nascimento do crediarista"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TxtCep 
         Height          =   285
         Left            =   5160
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Cep do crediarista"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox TxtCpf 
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   6
         ToolTipText     =   "Cpf do crediarista"
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox TxtEmail 
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Email do crediarista"
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox TxtNome 
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   0
         ToolTipText     =   "Nome do crediarista"
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox TxtEndereco 
         Height          =   285
         Left            =   1080
         MaxLength       =   200
         TabIndex        =   1
         ToolTipText     =   "Endereço do crediarista"
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox TxtBairro 
         Height          =   285
         Left            =   1080
         MaxLength       =   60
         TabIndex        =   2
         ToolTipText     =   "Bairro do crediarista"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox TxtCidade 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   4
         ToolTipText     =   "Cidade do crediarista"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Estado do crediarista"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtTel 
         Height          =   285
         Left            =   1080
         MaxLength       =   14
         TabIndex        =   8
         ToolTipText     =   "Número do telefone do crediarista"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox TxtCel 
         Height          =   285
         Left            =   5160
         MaxLength       =   14
         TabIndex        =   9
         ToolTipText     =   "Número do celular do crediarista"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox TxtObs 
         Height          =   645
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Observação sobre o crediarista"
         Top             =   3840
         Width           =   6495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0CCA
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0D2C
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0D96
         TabIndex        =   18
         Top             =   1200
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0DFC
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0E62
         TabIndex        =   20
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0EC6
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0F26
         TabIndex        =   22
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0F86
         TabIndex        =   23
         Top             =   1680
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":0FEC
         TabIndex        =   24
         Top             =   2640
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":1056
         TabIndex        =   25
         Top             =   2640
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":10BE
         TabIndex        =   26
         Top             =   3120
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":1122
         TabIndex        =   27
         Top             =   3600
         Width           =   1335
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
      TabIndex        =   14
      Top             =   4800
      Width           =   6735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2640
         OleObjectBlob   =   "FrmCrediarista_Inc.frx":1190
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
         TabIndex        =   13
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
         TabIndex        =   12
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCrediarista_Inc"
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
End Sub

Private Sub CmdOK_Click()
    
    Conecta
    
    Dim RecCredsta As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_Crediarista where CPF='" & TxtCpf.Text & "'"
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    If Not RecCredsta.EOF Then
        Desconecta
        VPStrBox = MsgBox("Já existe um crediarista cadastrado com este CPF.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        RecCredsta.AddNew
        RecCredsta("Nome") = TxtNome.Text
        RecCredsta("Endereco") = TxtEndereco.Text
        RecCredsta("Bairro") = TxtBairro.Text
        RecCredsta("Cep") = TxtCep.Text
        RecCredsta("Cidade") = TxtCidade.Text
        RecCredsta("Estado") = CboEstado.Text
        RecCredsta("DtNasc") = FormataDataUS(TxtDtNasc.Text)
        RecCredsta("Telefone") = TxtTel.Text
        RecCredsta("Celular") = TxtCel.Text
        RecCredsta("Cpf") = TxtCpf.Text
        RecCredsta("Email") = TxtEmail.Text
        RecCredsta("Obs") = TxtObs.Text
        RecCredsta.Update
    
        RecCredsta.Close
        
        Desconecta
        
        VPStrBox = MsgBox("Crediarista cadastrado.", vbInformation, "Pró Ótica 2004 - Informação")
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
    
    End If
    
End Sub

Private Sub Form_Resize()
  FrmCrediarista_Inc.Left = (MDIPrincipal.Width / 2) - (FrmCrediarista_Inc.Width / 2)
  FrmCrediarista_Inc.Top = (MDIPrincipal.Height / 3) - (FrmCrediarista_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 6150
    Width = 7065
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
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

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

