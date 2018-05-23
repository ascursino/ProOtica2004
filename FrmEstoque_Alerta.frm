VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmEstoque_Alerta 
   Caption         =   "Alerta de Estoque"
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
   Icon            =   "FrmEstoque_Alerta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6330
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
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
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
         Left            =   3480
         TabIndex        =   1
         ToolTipText     =   "Imprimir alerta de estoque"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   360
         OleObjectBlob   =   "FrmEstoque_Alerta.frx":0CCA
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
      Begin FPSpread.vaSpread GridAlerta 
         Height          =   2055
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   5895
         _Version        =   393216
         _ExtentX        =   10398
         _ExtentY        =   3625
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
         MaxCols         =   3
         MaxRows         =   1
         OperationMode   =   2
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmEstoque_Alerta.frx":0EFE
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmEstoque_Alerta.frx":124D
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "FrmEstoque_Alerta"
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

Private Sub Form_Resize()
  FrmEstoque_Alerta.Left = (MDIPrincipal.Width / 2) - (FrmEstoque_Alerta.Width / 2)
  FrmEstoque_Alerta.Top = (MDIPrincipal.Height / 3) - (FrmEstoque_Alerta.Height / 3)
End Sub



Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass

    Dim prod As String
    Dim qtdeprod As String
    Dim qtdemin As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridAlerta.MaxRows

        GridAlerta.Col = 1
        GridAlerta.Row = VLStrLinha
        prod = GridAlerta.Text

        GridAlerta.Col = 2
        GridAlerta.Row = VLStrLinha
        qtdeprod = GridAlerta.Text

        GridAlerta.Col = 3
        GridAlerta.Row = VLStrLinha
        qtdemin = GridAlerta.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03) " & _
        "VALUES ('" & prod & "','" & qtdeprod & "','" & qtdemin & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptEstoqueAlerta.Show

End Sub

Private Sub Form_Load()
    Height = 4320
    Width = 6450
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
        
    Conecta
    
    Dim RecAlert As New ADODB.Recordset
    Dim RecGrif As New ADODB.Recordset
    Dim Griffe As String
    
    StrSql = "SELECT E.QtdeMin,E.QtdeProd,P.CodProd,P.CodGriffe,P.TipoProd,P.Cor,P.Numero,P.Modelo,P.TamAro,P.TamPonte,P.Tipo,P.Chave " & _
             "FROM tb_Estoque as E,tb_Produto as P " & _
             "WHERE E.CodProd=P.CodProd AND E.QtdeProd <= E.QtdeMin " & _
             "ORDER BY P.TipoProd"
    RecAlert.Open StrSql, vgCon, 1, 3

    If Not RecAlert.EOF Then
        'Dim VLIntCodEst As Long
        Dim VLIntLinha As Long
    
        VLIntLinha = 1
        GridAlerta.MaxRows = VLIntLinha

        Do While Not RecAlert.EOF
            GridAlerta.Row = VLIntLinha
            GridAlerta.Lock = True

            StrSql = "Select Nome From tb_Griffe where CodGriffe=" & RecAlert!CodGriffe
            RecGrif.Open StrSql, vgCon, 1, 3

            If Not RecGrif.EOF Then
                Griffe = RecGrif!nome
            Else
                Griffe = ""
            End If
            
            RecGrif.Close
            
            'Produto
            GridAlerta.Col = 1
            If Griffe = "" Then
                'mostra dados para lentes
                GridAlerta.Text = RecAlert!tipoprod & " - " & VerificaNulo(RecAlert!tipo) & "/" & VerificaNulo(RecAlert!chave)
            Else
                'mostra dados para armação
                GridAlerta.Text = RecAlert!tipoprod & " - " & Griffe & "/" & VerificaNulo(RecAlert!cor) & "/" & VerificaNulo(RecAlert!Numero) & "/" & VerificaNulo(RecAlert!modelo) & "/" & VerificaNulo(RecAlert!TamAro) & "/" & VerificaNulo(RecAlert!TamPonte)
            End If

            'Qtde em estoque
            GridAlerta.Col = 2
            GridAlerta.Text = VerificaNulo(RecAlert!qtdeprod)
            GridAlerta.Lock = True

            'Qtde mínima
            GridAlerta.Col = 3
            GridAlerta.Text = VerificaNulo(RecAlert!qtdemin)
            GridAlerta.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridAlerta.MaxRows = GridAlerta.MaxRows + 1
            RecAlert.MoveNext
        Loop
        
        GridAlerta.MaxRows = GridAlerta.MaxRows - 1
        
        CmdImprimir.Enabled = True
        
    Else
        CmdImprimir.Enabled = False
    End If
    
    Desconecta

End Sub


