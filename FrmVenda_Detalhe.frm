VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmVenda_Detalhe 
   Caption         =   "Detalhe da Venda"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
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
   Icon            =   "FrmVenda_Detalhe.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   6825
   Begin VB.Frame FraProd 
      Caption         =   "Armação"
      Height          =   1575
      Left            =   240
      TabIndex        =   46
      Top             =   960
      Width           =   6375
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel95 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0CCA
         TabIndex        =   47
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel98 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0D3A
         TabIndex        =   48
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel99 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0DA8
         TabIndex        =   49
         Top             =   1080
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblGriffeTipo 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0E1C
         TabIndex        =   50
         Top             =   360
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorProd 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0E9A
         TabIndex        =   51
         Top             =   1080
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblQtde 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0F18
         TabIndex        =   52
         Top             =   720
         Width           =   3975
      End
   End
   Begin MSComctlLib.TabStrip TabDetalhe 
      Height          =   2055
      Left            =   120
      TabIndex        =   45
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produto 01"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produto 02"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produto 03"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produto 04"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produto 05"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Produto 06"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FraVista 
      Caption         =   "Finalização da venda / À vista"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   6615
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel85 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0F92
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel86 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":0FF6
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel87 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1060
         TabIndex        =   5
         Top             =   1200
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel88 
         Height          =   255
         Left            =   2880
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":10C4
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorVista 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1136
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDescVista 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":11AC
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalVista 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":120A
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoPagtoVista 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1282
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBancoVista 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":12F8
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblChequeVista 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1364
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame FraPrazoCheque 
      Caption         =   "Finalização da venda / A prazo - Cheque"
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   6615
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel72 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":13E4
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel73 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1448
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel74 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":14B2
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel75 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1516
         TabIndex        =   17
         Top             =   1080
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorCheque 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":157A
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblParcCheque 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":15F0
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblJurosCheque 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1654
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalCheque 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":16B2
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":172A
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorParcCheque 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1792
         TabIndex        =   25
         Top             =   1440
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEntradaCheque 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":181E
         TabIndex        =   26
         Top             =   360
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBancoCheque 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":18AC
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorEntradaCheque 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1918
         TabIndex        =   28
         Top             =   1200
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblChequeCheque 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":19A0
         TabIndex        =   29
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame FraPrazoCarne 
      Caption         =   "Finalização da venda / A prazo - Carnê"
      Height          =   1815
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   6615
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1A20
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1A84
         TabIndex        =   32
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1AEE
         TabIndex        =   33
         Top             =   840
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1B52
         TabIndex        =   34
         Top             =   1080
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorCarne 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1BB6
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblParcCarne 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1C2C
         TabIndex        =   36
         Top             =   600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblJurosCarne 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1C90
         TabIndex        =   37
         Top             =   840
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTotalCarne 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1CEE
         TabIndex        =   38
         Top             =   1080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1D66
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorParcCarne 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1DCE
         TabIndex        =   40
         Top             =   1440
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEntradaCarne 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1E5A
         TabIndex        =   41
         Top             =   360
         Width           =   2775
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBancoCarne 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1EE8
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorEntradaCarne 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1F54
         TabIndex        =   43
         Top             =   1200
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblChequeCarne 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":1FDC
         TabIndex        =   44
         Top             =   840
         Width           =   2415
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
      TabIndex        =   1
      Top             =   4800
      Width           =   6615
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1440
         OleObjectBlob   =   "FrmVenda_Detalhe.frx":205C
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
         Left            =   5280
         TabIndex        =   0
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmVenda_Detalhe.frx":2290
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblVendedor 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "FrmVenda_Detalhe.frx":22FA
      TabIndex        =   23
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FrmVenda_Detalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrProd01 As String
Public VPStrQtde01 As String
Public VPStrVenda01 As String
Public VPStrTipoProd01 As String

Public VPStrProd02 As String
Public VPStrQtde02 As String
Public VPStrVenda02 As String
Public VPStrTipoProd02 As String

Public VPStrProd03 As String
Public VPStrQtde03 As String
Public VPStrVenda03 As String
Public VPStrTipoProd03 As String

Public VPStrProd04 As String
Public VPStrQtde04 As String
Public VPStrVenda04 As String
Public VPStrTipoProd04 As String

Public VPStrProd05 As String
Public VPStrQtde05 As String
Public VPStrVenda05 As String
Public VPStrTipoProd05 As String

Public VPStrProd06 As String
Public VPStrQtde06 As String
Public VPStrVenda06 As String
Public VPStrTipoProd06 As String

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub Form_Resize()
  FrmVenda_Detalhe.Left = (MDIPrincipal.Width / 2) - (FrmVenda_Detalhe.Width / 2)
  FrmVenda_Detalhe.Top = (MDIPrincipal.Height / 3) - (FrmVenda_Detalhe.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 6135
    Width = 6945
    Top = 1095
    Left = 4155
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
        
    Conecta
    
    Dim RecVenda As New ADODB.Recordset
    Dim RecVendedor As New ADODB.Recordset
    Dim RecCred As New ADODB.Recordset
    Dim RecCredParc As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    Dim RecGrif As New ADODB.Recordset
    
    StrSql = "Select * from tb_Venda where CodVenda=" & VGIntCodVenda
    RecVenda.Open StrSql, vgCon, 1, 3
   
    '==== Primeiro produto =====
    StrSql = "Select CodGriffe,TipoProd,Tipo from tb_Produto where CodProd=" & RecVenda.Fields.Item(5).Value
    RecProd.Open StrSql, vgCon, 1, 3
    
    If Not RecProd.EOF Then
        If RecProd.Fields.Item(0).Value <> "" And RecProd.Fields.Item(0).Value <> 0 And IsNull(RecProd.Fields.Item(0).Value) = False Then
            StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecProd.Fields.Item(0).Value
            RecGrif.Open StrSql, vgCon, 1, 3
            
            VPStrProd01 = RecGrif.Fields.Item(0).Value
            RecGrif.Close
        Else
            VPStrProd01 = RecProd.Fields.Item(2).Value
        End If
        
        VPStrQtde01 = FormataNum(RecVenda.Fields.Item(17).Value)
        VPStrVenda01 = FormataMoeda(RecVenda.Fields.Item(23).Value)
        VPStrTipoProd01 = VerificaNulo(RecProd.Fields.Item(1).Value)
    Else
        VPStrProd01 = ""
        VPStrQtde01 = ""
        VPStrVenda01 = ""
        VPStrTipoProd01 = ""
    End If
    '==========================
    
    RecProd.Close
    
    '==== Segundo produto =====
    StrSql = "Select CodGriffe,TipoProd,Tipo from tb_Produto where CodProd=" & RecVenda.Fields.Item(6).Value
    RecProd.Open StrSql, vgCon, 1, 3
    
    If Not RecProd.EOF Then
        If RecProd.Fields.Item(0).Value <> "" And RecProd.Fields.Item(0).Value <> 0 And IsNull(RecProd.Fields.Item(0).Value) = False Then
            StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecProd.Fields.Item(0).Value
            RecGrif.Open StrSql, vgCon, 1, 3
            
            VPStrProd02 = RecGrif.Fields.Item(0).Value
            RecGrif.Close
        Else
            VPStrProd02 = RecProd.Fields.Item(2).Value
        End If
        
        VPStrQtde02 = FormataNum(RecVenda.Fields.Item(18).Value)
        VPStrVenda02 = FormataMoeda(RecVenda.Fields.Item(24).Value)
        VPStrTipoProd02 = VerificaNulo(RecProd.Fields.Item(1).Value)
    Else
        VPStrProd02 = ""
        VPStrQtde02 = ""
        VPStrVenda02 = ""
        VPStrTipoProd02 = ""
    End If
    '==========================
    
    RecProd.Close
    
    '==== Terceiro produto =====
    StrSql = "Select CodGriffe,TipoProd,Tipo from tb_Produto where CodProd=" & RecVenda.Fields.Item(7).Value
    RecProd.Open StrSql, vgCon, 1, 3
    
    If Not RecProd.EOF Then
        If RecProd.Fields.Item(0).Value <> "" And RecProd.Fields.Item(0).Value <> 0 And IsNull(RecProd.Fields.Item(0).Value) = False Then
            StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecProd.Fields.Item(0).Value
            RecGrif.Open StrSql, vgCon, 1, 3
            
            VPStrProd03 = RecGrif.Fields.Item(0).Value
            RecGrif.Close
        Else
            VPStrProd03 = RecProd.Fields.Item(2).Value
        End If
        
        VPStrQtde03 = FormataNum(RecVenda.Fields.Item(19).Value)
        VPStrVenda03 = FormataMoeda(RecVenda.Fields.Item(25).Value)
        VPStrTipoProd03 = VerificaNulo(RecProd.Fields.Item(1).Value)
    Else
        VPStrProd03 = ""
        VPStrQtde03 = ""
        VPStrVenda03 = ""
        VPStrTipoProd03 = ""
    End If
    '==========================
    
    RecProd.Close
    
    '==== Quarto produto =====
    StrSql = "Select CodGriffe,TipoProd,Tipo from tb_Produto where CodProd=" & RecVenda.Fields.Item(8).Value
    RecProd.Open StrSql, vgCon, 1, 3
    
    If Not RecProd.EOF Then
        If RecProd.Fields.Item(0).Value <> "" And RecProd.Fields.Item(0).Value <> 0 And IsNull(RecProd.Fields.Item(0).Value) = False Then
            StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecProd.Fields.Item(0).Value
            RecGrif.Open StrSql, vgCon, 1, 3
            
            VPStrProd04 = RecGrif.Fields.Item(0).Value
            RecGrif.Close
        Else
            VPStrProd04 = RecProd.Fields.Item(2).Value
        End If
        
        VPStrQtde04 = FormataNum(RecVenda.Fields.Item(20).Value)
        VPStrVenda04 = FormataMoeda(RecVenda.Fields.Item(26).Value)
        VPStrTipoProd04 = VerificaNulo(RecProd.Fields.Item(1).Value)
    Else
        VPStrProd04 = ""
        VPStrQtde04 = ""
        VPStrVenda04 = ""
        VPStrTipoProd04 = ""
    End If
    '==========================
    
    RecProd.Close
    
    '==== Quinto produto =====
    StrSql = "Select CodGriffe,TipoProd,Tipo from tb_Produto where CodProd=" & RecVenda.Fields.Item(9).Value
    RecProd.Open StrSql, vgCon, 1, 3
    
    If Not RecProd.EOF Then
        If RecProd.Fields.Item(0).Value <> "" And RecProd.Fields.Item(0).Value <> 0 And IsNull(RecProd.Fields.Item(0).Value) = False Then
            StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecProd.Fields.Item(0).Value
            RecGrif.Open StrSql, vgCon, 1, 3
            
            VPStrProd05 = RecGrif.Fields.Item(0).Value
            RecGrif.Close
        Else
            VPStrProd05 = RecProd.Fields.Item(2).Value
        End If
        
        VPStrQtde05 = FormataNum(RecVenda.Fields.Item(21).Value)
        VPStrVenda05 = FormataMoeda(RecVenda.Fields.Item(27).Value)
        VPStrTipoProd05 = VerificaNulo(RecProd.Fields.Item(1).Value)
    Else
        VPStrProd05 = ""
        VPStrQtde05 = ""
        VPStrVenda05 = ""
        VPStrTipoProd05 = ""
    End If
    '==========================
    
    RecProd.Close
    
    '==== Sexto produto =====
    StrSql = "Select CodGriffe,TipoProd,Tipo from tb_Produto where CodProd=" & RecVenda.Fields.Item(10).Value
    RecProd.Open StrSql, vgCon, 1, 3
    
    If Not RecProd.EOF Then
        If RecProd.Fields.Item(0).Value <> "" And RecProd.Fields.Item(0).Value <> 0 And IsNull(RecProd.Fields.Item(0).Value) = False Then
            StrSql = "Select Nome from tb_Griffe where CodGriffe=" & RecProd.Fields.Item(0).Value
            RecGrif.Open StrSql, vgCon, 1, 3
            
            VPStrProd06 = RecGrif.Fields.Item(0).Value
            RecGrif.Close
        Else
            VPStrProd06 = RecProd.Fields.Item(2).Value
        End If
        
        VPStrQtde06 = FormataNum(RecVenda.Fields.Item(22).Value)
        VPStrVenda06 = FormataMoeda(RecVenda.Fields.Item(28).Value)
        VPStrTipoProd06 = VerificaNulo(RecProd.Fields.Item(1).Value)
    Else
        VPStrProd06 = ""
        VPStrQtde06 = ""
        VPStrVenda06 = ""
        VPStrTipoProd06 = ""
    End If
    '==========================
    
    RecProd.Close
    
    
    StrSql = "Select Nome from tb_Vendedor where CodVendedor=" & RecVenda.Fields.Item(1).Value
    RecVendedor.Open StrSql, vgCon, 1, 3
    
    LblVendedor.Caption = VerificaNulo(RecVendedor.Fields.Item(0).Value)
    
    If RecVenda.Fields.Item(29).Value = "A prazo - Cheque" Then
        StrSql = "Select ValorVenda,Parcela,Juros,ValorTotal,TipoEntr,ValorEntr,NumBanco,NumCheque from tb_Crediario where CodCred=" & RecVenda.Fields.Item(3).Value
        RecCred.Open StrSql, vgCon, 1, 3
        
        StrSql = "Select Valor from tb_Crediario_Parcela where CodCred=" & RecVenda.Fields.Item(3).Value
        RecCredParc.Open StrSql, vgCon, 1, 3

        LblValorCheque.Caption = FormataMoeda(RecCred.Fields.Item(0).Value)
        LblParcCheque.Caption = FormataNum(RecCred.Fields.Item(1).Value)
        LblJurosCheque.Caption = FormataNum(RecCred.Fields.Item(2).Value) & "%"
        LblTotalCheque.Caption = FormataMoeda(RecCred.Fields.Item(3).Value)
        LblEntradaCheque.Caption = VerificaNulo(RecCred.Fields.Item(4).Value)
        
        If RecCred.Fields.Item(4).Value = "Cheque" Then
            LblBancoCheque.Caption = "Banco: " & VerificaNulo(RecCred.Fields.Item(6).Value)
            LblChequeCheque.Caption = "Cheque: " & VerificaNulo(RecCred.Fields.Item(7).Value)
            LblValorEntradaCheque.Caption = "Valor entrada: " & FormataMoeda(RecCred.Fields.Item(5).Value)
            LblBancoCheque.Visible = True
            LblChequeCheque.Visible = True
            LblValorEntradaCheque.Visible = True
            
        ElseIf RecCred.Fields.Item(4).Value = "Dinheiro" Then
            LblValorEntradaCheque.Caption = "Valor entrada: " & FormataMoeda(RecCred.Fields.Item(5).Value)
            LblBancoCheque.Visible = False
            LblChequeCheque.Visible = False
            LblValorEntradaCheque.Visible = True
        Else
            LblBancoCheque.Visible = False
            LblChequeCheque.Visible = False
            LblValorEntradaCheque.Visible = False
        End If
        
        LblValorParcCheque.Caption = FormataNum(RecCred.Fields.Item(1).Value) & " parcela(s) de " & FormataMoeda(RecCredParc.Fields.Item(0).Value)
        
        FraPrazoCheque.Visible = True
        FraPrazoCarne.Visible = False
        FraVista.Visible = False
        
    ElseIf RecVenda.Fields.Item(29).Value = "A prazo - Carnê" Then
        StrSql = "Select ValorVenda,Parcela,Juros,ValorTotal,TipoEntr,ValorEntr,NumBanco,NumCheque from tb_Crediario where CodCred=" & RecVenda.Fields.Item(3).Value
        RecCred.Open StrSql, vgCon, 1, 3
        
        StrSql = "Select Valor from tb_Crediario_Parcela where CodCred=" & RecVenda.Fields.Item(3).Value
        RecCredParc.Open StrSql, vgCon, 1, 3

        LblValorCarne.Caption = FormataMoeda(RecCred.Fields.Item(0).Value)
        LblParcCarne.Caption = FormataNum(RecCred.Fields.Item(1).Value)
        LblJurosCarne.Caption = FormataNum(RecCred.Fields.Item(2).Value) & "%"
        LblTotalCarne.Caption = FormataMoeda(RecCred.Fields.Item(3).Value)
        LblEntradaCarne.Caption = VerificaNulo(RecCred.Fields.Item(4).Value)
        
        If RecCred.Fields.Item(4).Value = "Cheque" Then
            LblBancoCarne.Caption = "Banco: " & VerificaNulo(RecCred.Fields.Item(6).Value)
            LblChequeCarne.Caption = "Cheque: " & VerificaNulo(RecCred.Fields.Item(7).Value)
            LblValorEntradaCarne.Caption = "Valor entrada: " & FormataMoeda(RecCred.Fields.Item(5).Value)
            LblBancoCarne.Visible = True
            LblChequeCarne.Visible = True
            LblValorEntradaCarne.Visible = True
            
        ElseIf RecCred.Fields.Item(4).Value = "Dinheiro" Then
            LblValorEntradaCarne.Caption = "Valor entrada: " & FormataMoeda(RecCred.Fields.Item(5).Value)
            LblBancoCarne.Visible = False
            LblChequeCarne.Visible = False
            LblValorEntradaCarne.Visible = True
        Else
            LblBancoCarne.Visible = False
            LblChequeCarne.Visible = False
            LblValorEntradaCarne.Visible = False
        End If
        
        LblValorParcCarne.Caption = FormataNum(RecCred.Fields.Item(1).Value) & " parcela(s) de " & FormataMoeda(RecCredParc.Fields.Item(0).Value)
        
        FraPrazoCheque.Visible = False
        FraPrazoCarne.Visible = True
        FraVista.Visible = False
    
    ElseIf RecVenda.Fields.Item(29).Value = "À vista" Then
        LblValorVista.Caption = FormataMoeda(RecVenda.Fields.Item(30).Value)
        
        If RecVenda.Fields.Item(31).Value <> "" And IsNull(RecVenda.Fields.Item(31).Value) = False Then
            LblDescVista.Caption = VerificaNulo(RecVenda.Fields.Item(31).Value) & "%"
        Else
            LblDescVista.Caption = ""
        End If
        
        LblTotalVista.Caption = FormataMoeda(RecVenda.Fields.Item(32).Value)
        LblTipoPagtoVista.Caption = VerificaNulo(RecVenda.Fields.Item(33).Value)
        
        If RecVenda.Fields.Item(33).Value = "Dinheiro" Then
            LblBancoVista.Visible = False
            LblChequeVista.Visible = False
        Else
            LblBancoVista.Caption = "Banco: " & VerificaNulo(RecVenda.Fields.Item(34).Value)
            LblChequeVista.Caption = "Cheque: " & VerificaNulo(RecVenda.Fields.Item(35).Value)
            LblBancoVista.Visible = True
            LblChequeVista.Visible = True
        End If
        
        FraPrazoCheque.Visible = False
        FraPrazoCarne.Visible = False
        FraVista.Visible = True
    
    End If
   
    Desconecta
    
    FraProd.Caption = VPStrTipoProd01
    LblGriffeTipo.Caption = VPStrProd01
    LblQtde.Caption = VPStrQtde01
    LblValorProd.Caption = VPStrVenda01
    
End Sub

Private Sub TabDetalhe_Click()
    If TabDetalhe.Tabs.Item(1).Selected = True Then
        '=== PRODUTO 01 ===
        If VPStrTipoProd01 <> "" Then
            FraProd.Caption = VPStrTipoProd01
        Else
            FraProd.Caption = "Sem produto"
        End If
        
        LblGriffeTipo.Caption = VPStrProd01
        LblQtde.Caption = VPStrQtde01
        LblValorProd.Caption = VPStrVenda01
    
    ElseIf TabDetalhe.Tabs.Item(2).Selected = True Then
        '=== PRODUTO 02 ===
        If VPStrTipoProd02 <> "" Then
            FraProd.Caption = VPStrTipoProd02
        Else
            FraProd.Caption = "Sem produto"
        End If
        
        LblGriffeTipo.Caption = VPStrProd02
        LblQtde.Caption = VPStrQtde02
        LblValorProd.Caption = VPStrVenda02
        
    ElseIf TabDetalhe.Tabs.Item(3).Selected = True Then
        '=== PRODUTO 03 ===
        If VPStrTipoProd03 <> "" Then
            FraProd.Caption = VPStrTipoProd03
        Else
            FraProd.Caption = "Sem produto"
        End If
        
        LblGriffeTipo.Caption = VPStrProd03
        LblQtde.Caption = VPStrQtde03
        LblValorProd.Caption = VPStrVenda03
        
    ElseIf TabDetalhe.Tabs.Item(4).Selected = True Then
        '=== PRODUTO 04 ===
        If VPStrTipoProd04 <> "" Then
            FraProd.Caption = VPStrTipoProd04
        Else
            FraProd.Caption = "Sem produto"
        End If

        
        LblGriffeTipo.Caption = VPStrProd04
        LblQtde.Caption = VPStrQtde04
        LblValorProd.Caption = VPStrVenda04
        
    ElseIf TabDetalhe.Tabs.Item(5).Selected = True Then
        '=== PRODUTO 05 ===
        If VPStrTipoProd05 <> "" Then
            FraProd.Caption = VPStrTipoProd05
        Else
            FraProd.Caption = "Sem produto"
        End If
        
        LblGriffeTipo.Caption = VPStrProd05
        LblQtde.Caption = VPStrQtde05
        LblValorProd.Caption = VPStrVenda05
        
    ElseIf TabDetalhe.Tabs.Item(6).Selected = True Then
        '=== PRODUTO 06 ===
        If VPStrTipoProd06 <> "" Then
            FraProd.Caption = VPStrTipoProd06
        Else
            FraProd.Caption = "Sem produto"
        End If
        
        LblGriffeTipo.Caption = VPStrProd06
        LblQtde.Caption = VPStrQtde06
        LblValorProd.Caption = VPStrVenda06
    End If
End Sub
