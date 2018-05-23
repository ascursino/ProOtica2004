VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmEstoque_Inc_Extra 
   Caption         =   "Inclusão de Griffe de Armação"
   ClientHeight    =   2250
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
   Icon            =   "FrmEstoque_Inc_Extra.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   6930
   Begin VB.Frame FraGriffe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.TextBox TxtGriffe 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Nome da griffe do produto"
         Top             =   480
         Width           =   5175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "FrmEstoque_Inc_Extra.frx":0CCA
         TabIndex        =   5
         Top             =   480
         Width           =   735
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
      TabIndex        =   3
      Top             =   1440
      Width           =   6735
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmEstoque_Inc_Extra.frx":0D30
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
         TabIndex        =   2
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
         TabIndex        =   1
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmEstoque_Inc_Extra"
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
        FrmProduto_Inc.Enabled = True
    End If
End Sub

Private Sub CmdOK_Click()
    
    Conecta
    
    Dim RecGrif As New ADODB.Recordset
    
    StrSql = "SELECT * FROM tb_Griffe where Nome='" & TxtGriffe.Text & "'"
    RecGrif.Open StrSql, vgCon, 1, 3
    
    If Not RecGrif.EOF Then
        Desconecta
        VPStrBox = MsgBox("Já existe uma griffe com esse nome.", vbInformation, "Pró Ótica 2004 - Informação")
        TxtGriffe.SetFocus
    Else
        RecGrif.AddNew
        RecGrif("Nome") = TxtGriffe.Text
        RecGrif.Update
        
        Desconecta
        
        If VGStrIncluirProd <> "" Then
            VGStrIncluirProd = ""
            FrmProduto_Inc.MontaCboGrif
        End If
        
        FrmProduto_Inc.Enabled = True
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
        
    End If
    
End Sub

Private Sub Form_Resize()
  FrmEstoque_Inc_Extra.Left = (MDIPrincipal.Width / 2) - (FrmEstoque_Inc_Extra.Width / 2)
  FrmEstoque_Inc_Extra.Top = (MDIPrincipal.Height / 3) - (FrmEstoque_Inc_Extra.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 2760
    Width = 7050
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    
    If VGStrIncluirProd <> "" Then
        FrmProduto_Inc.Enabled = False
    End If
    
End Sub

