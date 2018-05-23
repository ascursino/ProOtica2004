VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCaixa_APagar_Inc 
   Caption         =   "Inclusão de Contas A Pagar"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
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
   Icon            =   "FrmCaixa_APagar_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   3945
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
      TabIndex        =   7
      Top             =   3600
      Width           =   3735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_APagar_Inc.frx":0CCA
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
         Left            =   2400
         TabIndex        =   5
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
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "Efetuar inclusão de conta"
         Top             =   240
         Width           =   1095
      End
   End
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
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3735
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Valor da conta a pagar"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TxtDtVenc 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Data do vencimento da conta"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtDescr 
         Height          =   1365
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Descrição da conta a pagar"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.ComboBox CboTipo 
         Height          =   315
         ItemData        =   "FrmCaixa_APagar_Inc.frx":0EFE
         Left            =   1320
         List            =   "FrmCaixa_APagar_Inc.frx":0F0E
         TabIndex        =   0
         ToolTipText     =   "Tipo de conta a pagar"
         Top             =   240
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Inc.frx":0F59
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Inc.frx":0FC7
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Inc.frx":1035
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Inc.frx":1099
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCaixa_APagar_Inc"
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
    
    If CboTipo.Text = "" Or TxtDtVenc.Text = "" Or TxtValor.Text = "" Or TxtDescr.Text = "" Then
        VPStrBox = MsgBox("Preencha todos os campos.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        Conecta
        
        Dim RecPag As New ADODB.Recordset
        
        StrSql = "Select * From tb_ContaPagar"
        RecPag.Open StrSql, vgCon, 1, 3
        
        RecPag.AddNew
        RecPag("Tipo") = CboTipo.Text
        RecPag("Vencimento") = FormataDataUS(TxtDtVenc.Text)
        RecPag("Valor") = Mid(TxtValor.Text, 4)
        RecPag("Descricao") = TxtDescr.Text
        RecPag("Pago") = "não"
        RecPag.Update
        
        Desconecta
        
        FrmCaixa_APagar_Cons.CmdPesqPag.Value = True
        
        VPStrBox = MsgBox("Pagamento cadastrado.", vbInformation, "Pró Ótica 2004 - Informação")
        
        Call MontaCbos
        
        TxtDtVenc.Text = "__/__/____"
        TxtValor.Text = ""
        TxtDescr.Text = ""
        
        CboTipo.SetFocus
    End If
    
End Sub

Private Sub Form_Resize()
  FrmCaixa_APagar_Inc.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_APagar_Inc.Width / 2)
  FrmCaixa_APagar_Inc.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_APagar_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4920
    Width = 4065
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmCaixa_APagar_Cons.Enabled = False
        
    Call MontaCbos
    
    TxtDtVenc.Text = "__/__/____"
End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecTipo As New ADODB.Recordset
    
    StrSql = "Select distinct Tipo From tb_ContaPagar"
    RecTipo.Open StrSql, vgCon, 1, 3
    
    CboTipo.Clear
    CboTipo.AddItem ("")
    Do While Not RecTipo.EOF
        CboTipo.AddItem (RecTipo.Fields.Item(0).Value)
        RecTipo.MoveNext
    Loop
    
    Desconecta
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
