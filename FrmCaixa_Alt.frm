VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCaixa_Alt 
   Caption         =   "Alteração do Movimento de Caixa"
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
   Icon            =   "FrmCaixa_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   3945
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
      TabIndex        =   9
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox CboTipoMov 
         Height          =   315
         ItemData        =   "FrmCaixa_Alt.frx":0CCA
         Left            =   720
         List            =   "FrmCaixa_Alt.frx":0CCC
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Tipo de movimento de caixa"
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox TxtDesc 
         Height          =   1245
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Descrição do movimento de caixa"
         Top             =   2040
         Width           =   3495
      End
      Begin VB.OptionButton OptCred 
         Caption         =   "Crédito"
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Crédito"
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton OptDeb 
         Caption         =   "Débito"
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         ToolTipText     =   "Débito"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxtValor 
         Height          =   285
         Left            =   720
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Valor do movimento de caixa"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox CboTipoPagto 
         Height          =   315
         ItemData        =   "FrmCaixa_Alt.frx":0CCE
         Left            =   720
         List            =   "FrmCaixa_Alt.frx":0CD0
         Sorted          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Tipo de pagamento para o movimento de caixa"
         Top             =   720
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_Alt.frx":0CD2
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_Alt.frx":0D34
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_Alt.frx":0D98
         TabIndex        =   12
         Top             =   1800
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_Alt.frx":0E04
         TabIndex        =   13
         Top             =   720
         Width           =   615
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
      TabIndex        =   8
      Top             =   3600
      Width           =   3735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   240
         OleObjectBlob   =   "FrmCaixa_Alt.frx":0E68
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
         Left            =   2400
         TabIndex        =   7
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
         ToolTipText     =   "Efetuar alteração"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCaixa_Alt"
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
    
    If CboTipoMov.Text = "" Or CboTipoPagto.Text = "" Or TxtValor.Text = "" Or TxtDesc.Text = "" Or (OptCred.Value = False And OptDeb.Value = False) Then
        VPStrBox = MsgBox("Preencha todos os campos.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        Conecta
        
        Dim RecCx As New ADODB.Recordset
        
        StrSql = "Select * From tb_Caixa where CodCx=" & VGIntCodCx
        RecCx.Open StrSql, vgCon, 1, 3
        
        RecCx("TipoMov") = CboTipoMov.Text
        RecCx("Valor") = Mid(TxtValor.Text, 4)
        
        If OptCred.Value = True Then
            RecCx("TipoValor") = "credito"
            
        ElseIf OptDeb.Value = True Then
            RecCx("TipoValor") = "debito"
        End If
        
        RecCx("Descricao") = TxtDesc.Text
        RecCx("TipoPagto") = CboTipoPagto.Text
        RecCx.Update
        
        Desconecta
        
        FrmPrincipal.CmdPesqCx.Value = True
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Ótica 2004 - Informação")
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
        
    End If
    
End Sub

Private Sub Form_Resize()
  FrmCaixa_Alt.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_Alt.Width / 2)
  FrmCaixa_Alt.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4920
    Width = 4065
    Top = 1365
    Left = 3795
        
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (FrmCaixa_Alt.hwnd)

    MDIPrincipal.Enabled = False
    
    Call MontaCbos
    
    Conecta
    
    Dim RecCx As New ADODB.Recordset
    
    StrSql = "Select * From tb_Caixa where CodCx=" & VGIntCodCx
    RecCx.Open StrSql, vgCon, 1, 3
    
    CboTipoMov.Text = RecCx.Fields.Item(3).Value
    CboTipoPagto.Text = RecCx.Fields.Item(7).Value
    TxtValor.Text = FormataMoeda(RecCx.Fields.Item(4).Value)
    
    If RecCx.Fields.Item(5).Value = "credito" Then
        OptCred.Value = True
        OptDeb.Value = False
        
    ElseIf RecCx.Fields.Item(5).Value = "debito" Then
        OptCred.Value = False
        OptDeb.Value = True
        
    End If
    TxtDesc.Text = RecCx.Fields.Item(6).Value
    
    Desconecta

End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecTipoMov As New ADODB.Recordset
    Dim RecTipoPagto As New ADODB.Recordset
    
    StrSql = "Select distinct TipoMov From tb_Caixa"
    RecTipoMov.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct TipoPagto From tb_Caixa"
    RecTipoPagto.Open StrSql, vgCon, 1, 3
        
    CboTipoMov.Clear
    CboTipoMov.AddItem ("")
    Do While Not RecTipoMov.EOF
        CboTipoMov.AddItem (RecTipoMov.Fields.Item(0).Value)
        RecTipoMov.MoveNext
    Loop
    
    CboTipoPagto.Clear
    CboTipoPagto.AddItem ("")
    Do While Not RecTipoPagto.EOF
        CboTipoPagto.AddItem (RecTipoPagto.Fields.Item(0).Value)
        RecTipoPagto.MoveNext
    Loop
    
    Desconecta
End Sub

Private Sub TxtValor_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtValor_LostFocus()
    If TxtValor.Text <> "" Then
        TxtValor.Text = FormataMoeda(TxtValor.Text)
    End If
End Sub
