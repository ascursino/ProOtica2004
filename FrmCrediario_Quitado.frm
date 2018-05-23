VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCrediario_Quitado 
   Caption         =   "Parcelas do Crediário Quitadas"
   ClientHeight    =   4920
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
   Icon            =   "FrmCrediario_Quitado.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
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
      Top             =   4080
      Width           =   3975
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   480
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":0CCA
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
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":0EFE
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":0F78
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":0FE6
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":1054
         TabIndex        =   7
         Top             =   1800
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":10B8
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":1122
         TabIndex        =   9
         Top             =   2520
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":1190
         TabIndex        =   10
         Top             =   2880
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorParc 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":1202
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblValorPago 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":126A
         TabIndex        =   12
         Top             =   1440
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtVenc 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":12D2
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblJuros 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":133E
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDesc 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":139C
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtQuit 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":13FA
         TabIndex        =   16
         Top             =   2520
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTipoPagto 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":1466
         TabIndex        =   17
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Frame FraPagtoChq 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1200
         TabIndex        =   3
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
         Begin ACTIVESKINLibCtl.SkinLabel LblBanco 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "FrmCrediario_Quitado.frx":14DC
            TabIndex        =   18
            Top             =   120
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblCheque 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "FrmCrediario_Quitado.frx":153A
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCrediario_Quitado.frx":15A8
            TabIndex        =   22
            Top             =   120
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmCrediario_Quitado.frx":1612
            TabIndex        =   23
            Top             =   480
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":167E
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumParc 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "FrmCrediario_Quitado.frx":16E6
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmCrediario_Quitado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub Form_Resize()
  FrmCrediario_Quitado.Left = (MDIPrincipal.Width / 2) - (FrmCrediario_Quitado.Width / 2)
  FrmCrediario_Quitado.Top = (MDIPrincipal.Height / 3) - (FrmCrediario_Quitado.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 5430
    Width = 4305
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Conecta
    
    Dim RecParc As New ADODB.Recordset
    
    StrSql = "SELECT P.NumParc,P.Vencimento,P.Valor,PQ.DtPagto,PQ.Juros,PQ.Desconto," & _
             "PQ.ValorPago,PQ.TipoPagto,PQ.NumBanco,PQ.NumCheque " & _
             "FROM tb_Crediario_Parcela as P, tb_Crediario_Parcela_Quitacao as PQ " & _
             "WHERE PQ.CodParc=P.CodParc AND P.CodParc=" & VGIntCodParc
    RecParc.Open StrSql, vgCon, 1, 3
    
    LblNumParc.Caption = FormataNum(RecParc.Fields.Item(0).Value)
    LblValorParc.Caption = FormataMoeda(RecParc.Fields.Item(2).Value)
    LblDtVenc.Caption = FormataData(RecParc.Fields.Item(1).Value)
    LblValorPago.Caption = FormataMoeda(RecParc.Fields.Item(6).Value)
    
    If RecParc.Fields.Item(4).Value <> "" And IsNull(RecParc.Fields.Item(4).Value) = False Then
        LblJuros.Caption = RecParc.Fields.Item(4).Value & "%"
    Else
        LblJuros.Caption = "0%"
    End If
    
    If RecParc.Fields.Item(5).Value <> "" And IsNull(RecParc.Fields.Item(5).Value) = False Then
        LblDesc.Caption = RecParc.Fields.Item(5).Value & "%"
    Else
        LblDesc.Caption = "0%"
    End If
    
    LblDtQuit.Caption = FormataData(RecParc.Fields.Item(3).Value)
    LblTipoPagto.Caption = RecParc.Fields.Item(7).Value
    
    If RecParc.Fields.Item(7).Value = "Cheque" Then
        LblBanco.Caption = RecParc.Fields.Item(8).Value
        LblCheque.Caption = RecParc.Fields.Item(9).Value
        FraPagtoChq.Visible = True
    Else
        FraPagtoChq.Visible = False
    End If
    
    Desconecta
    
    MDIPrincipal.Enabled = False
    
End Sub


