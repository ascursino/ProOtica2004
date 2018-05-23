VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCaixa_APagar_Baixado 
   Caption         =   "Pagamentos baixados"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
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
   Icon            =   "FrmCaixa_APagar_Baixado.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4185
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
      TabIndex        =   2
      Top             =   2880
      Width           =   3975
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   480
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":0CCA
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
         Left            =   2640
         TabIndex        =   0
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
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":0EFE
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":0F62
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":0FD0
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":103E
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":10A6
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValor 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":1118
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorPago 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":1180
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtVenc 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":11E8
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtPagto 
         Height          =   255
         Left            =   1080
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":1254
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoPagto 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":12C0
         TabIndex        =   13
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Frame FraPagtoChq 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
         Visible         =   0   'False
         Width           =   2655
         Begin ACTIVESKINLibCtl.SkinLabel LblBanco 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":1336
            TabIndex        =   14
            Top             =   120
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblCheque 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":1394
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":1402
            TabIndex        =   16
            Top             =   120
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCaixa_APagar_Baixado.frx":146C
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "FrmCaixa_APagar_Baixado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    Unload Me
    
    FrmCaixa_APagar_Cons.Enabled = True
End Sub

Private Sub Form_Resize()
  FrmCaixa_APagar_Baixado.Left = (MDIPrincipal.Width / 2) - (FrmCaixa_APagar_Baixado.Width / 2)
  FrmCaixa_APagar_Baixado.Top = (MDIPrincipal.Height / 3) - (FrmCaixa_APagar_Baixado.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4200
    Width = 4305
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Conecta
    
    Dim RecPag As New ADODB.Recordset
    
    StrSql = "SELECT P.Vencimento,P.Valor,PT.DtPagto,PT.ValorPago,PT.TipoPagto,PT.NumBanco,PT.NumCheque " & _
             "FROM tb_ContaPagar as P, tb_ContaPagar_Pagto as PT " & _
             "WHERE PT.CodCPag=P.CodCPag AND P.CodCPag=" & VGIntCodPagar
    RecPag.Open StrSql, vgCon, 1, 3
    
    LblValor.Caption = FormataMoeda(RecPag.Fields.Item(1).Value)
    LblDtVenc.Caption = FormataData(RecPag.Fields.Item(0).Value)
    LblValorPago.Caption = FormataMoeda(RecPag.Fields.Item(3).Value)
    LblDtPagto.Caption = FormataData(RecPag.Fields.Item(2).Value)
    LblTipoPagto.Caption = RecPag.Fields.Item(4).Value
    
    If RecPag.Fields.Item(4).Value = "Cheque" Then
        LblBanco.Caption = RecPag.Fields.Item(5).Value
        LblCheque.Caption = RecPag.Fields.Item(6).Value
        FraPagtoChq.Visible = True
    Else
        FraPagtoChq.Visible = False
    End If
    
    Desconecta
    
    FrmCaixa_APagar_Cons.Enabled = False
    
End Sub
