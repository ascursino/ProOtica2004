VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmCrediarista_Lista 
   Caption         =   "Lista de Crediarista"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
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
   Icon            =   "FrmCrediarista_Lista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6330
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "FrmCrediarista_Lista.frx":0CCA
      Top             =   1680
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
      TabIndex        =   4
      Top             =   3000
      Width           =   6135
      Begin VB.CheckBox ChkProp 
         Caption         =   "O próprio"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Escolher o próprio cliente para ser o crediarista"
         Top             =   240
         Width           =   1215
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
         Left            =   4800
         TabIndex        =   2
         ToolTipText     =   "Fechar"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      Begin FPSpread.vaSpread GridCred 
         Height          =   2415
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5895
         _Version        =   393216
         _ExtentX        =   10398
         _ExtentY        =   4260
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   2
         MaxRows         =   1
         OperationMode   =   2
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmCrediarista_Lista.frx":0EFE
      End
   End
End
Attribute VB_Name = "FrmCrediarista_Lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String

Private Sub ChkProp_Click()
    If ChkProp.Value = 1 Then
        Conecta
        
        Dim RecCli As New ADODB.Recordset
        Dim RecCredsta As New ADODB.Recordset
        
        StrSql = "SELECT * FROM tb_Cliente where CodCli=" & VGIntCodCli
        RecCli.Open StrSql, vgCon, 1, 3
        
        StrSql = "SELECT * FROM tb_Crediarista where Nome='" & RecCli!nome & "' and DtNasc=#" & FormataDataUS(RecCli!dtnasc) & "#"
        RecCredsta.Open StrSql, vgCon, 1, 3
        
        If Not RecCredsta.EOF Then
            VGIntCodCredstaVenda = RecCredsta!CodCredsta
            
            FrmVenda_Inc.Enabled = True
            FrmVenda_Inc.LblCredstaCarne.Caption = "Crediarista: " & RecCredsta!nome
            FrmVenda_Inc.LblCredstaCheque.Caption = "Crediarista: " & RecCredsta!nome
        
        Else
            RecCredsta.AddNew
            RecCredsta("Nome") = RecCli!nome
            RecCredsta("Endereco") = RecCli!endereco
            RecCredsta("Bairro") = RecCli!bairro
            RecCredsta("Cep") = RecCli!cep
            RecCredsta("Cidade") = RecCli!cidade
            RecCredsta("Estado") = RecCli!Estado
            RecCredsta("DtNasc") = FormataDataUS(RecCli!dtnasc)
            RecCredsta("Telefone") = RecCli!telefone
            RecCredsta("Celular") = RecCli!celular
            RecCredsta("Cpf") = RecCli!cpf
            RecCredsta("Email") = RecCli!email
            RecCredsta.Update
            
            RecCredsta.Close
            
            StrSql = "SELECT Max(CodCredsta) as CodCredsta FROM tb_Crediarista"
            RecCredsta.Open StrSql, vgCon, 1, 3
            
            VGIntCodCredstaVenda = RecCredsta!CodCredsta
            
            FrmVenda_Inc.Enabled = True
            FrmVenda_Inc.LblCredstaCarne.Caption = "Crediarista: " & RecCli!nome
            FrmVenda_Inc.LblCredstaCheque.Caption = "Crediarista: " & RecCli!nome
        End If
        
        Desconecta
        
        Unload Me
    End If
End Sub

Private Sub CmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
  FrmCrediarista_Lista.Left = (MDIPrincipal.Width / 2) - (FrmCrediarista_Lista.Width / 2)
  FrmCrediarista_Lista.Top = (MDIPrincipal.Height / 3) - (FrmCrediarista_Lista.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4320
    Width = 6450
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    'FrmVenda_Inc.Enabled = False
        
    If VGStrCredLista = "venda" Then
        ChkProp.Visible = True
    Else
        ChkProp.Visible = False
    End If
        
    Conecta
    
    Dim RecCredsta As New ADODB.Recordset
    Dim VLIntLinha As Long
    
    StrSql = "SELECT CodCredsta,Nome FROM tb_Crediarista order by Nome"
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    If Not RecCredsta.EOF Then
        
        VLIntLinha = 1
        GridCred.MaxRows = VLIntLinha
         
        Do While Not RecCredsta.EOF
                 
            GridCred.Row = VLIntLinha
            GridCred.Lock = True
            
            'Crediarista
            GridCred.Col = 1
            GridCred.Text = VerificaNulo(RecCredsta.Fields.Item(1).Value)
            GridCred.Lock = True
            
            'CodCredsta
            GridCred.Col = 2
            GridCred.Text = Val(RecCredsta.Fields.Item(0).Value)
            GridCred.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridCred.MaxRows = GridCred.MaxRows + 1
            RecCredsta.MoveNext
         Loop
         
         GridCred.MaxRows = GridCred.MaxRows - 1
         
    End If
    
    Desconecta
    
End Sub

Private Sub GridCred_DblClick(ByVal Col As Long, ByVal Row As Long)
    'pegar nome e código do crediarista
    GridCred.Row = Row
    GridCred.Col = 1
    If GridCred.Text <> "" And GridCred.Text <> "Crediarista" Then
        VGStrNomeCredsta = GridCred.Text
    
        GridCred.Row = Row
        GridCred.Col = 2
        If GridCred.Text <> "" Then
            VGIntCodCredstaVenda = GridCred.Text
        Else
            VGIntCodCredstaVenda = 0
        End If
        
        Unload Me
        
        If VGStrCredLista = "venda" Then
            VGStrCredLista = ""
            FrmVenda_Inc.Enabled = True
            FrmVenda_Inc.LblCredstaCarne.Caption = "Crediarista: " & VGStrNomeCredsta
            FrmVenda_Inc.LblCredstaCheque.Caption = "Crediarista: " & VGStrNomeCredsta
        Else
            FrmCrediarista_Alt.Show
        End If
    End If
End Sub
