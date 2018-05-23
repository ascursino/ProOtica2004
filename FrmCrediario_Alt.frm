VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCrediario_Alt 
   Caption         =   "Alteração de Parcela"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
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
   Icon            =   "FrmCrediario_Alt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   6345
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
      TabIndex        =   6
      Top             =   1800
      Width           =   6135
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   720
         OleObjectBlob   =   "FrmCrediario_Alt.frx":0CCA
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
         TabIndex        =   4
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
         Left            =   3480
         TabIndex        =   3
         ToolTipText     =   "Efetuar alteração"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6135
      Begin VB.TextBox TxtValorParc 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         ToolTipText     =   "Valor das parcelas"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenc 
         Height          =   285
         Left            =   4320
         TabIndex        =   2
         ToolTipText     =   "Data de vencimento das parcelas"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox CboCredsta 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Nome do crediarista"
         Top             =   360
         Width           =   4815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Alt.frx":0EFE
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCrediario_Alt.frx":0F6E
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "FrmCrediario_Alt.frx":0FE8
         TabIndex        =   9
         ToolTipText     =   "Vencimento das parcelas"
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmCrediario_Alt"
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
    
    
    If CboCredsta.Text = "" Or TxtValorParc.Text = "" Or TxtDtVenc.Text = "" Then
        VPStrBox = MsgBox("Não pode conter campos em branco.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        Conecta
        
        Dim RecCred As New ADODB.Recordset
        Dim RecParc As New ADODB.Recordset
        
        StrSql = "SELECT * FROM tb_Crediario where CodCred=" & VGIntCodCred
        RecCred.Open StrSql, vgCon, 1, 3
        
        RecCred("CodCredsta") = Mid(CboCredsta.Text, Len(CboCredsta.Text) - 10)
        RecCred.Update
        
        RecCred.Close
        
        StrSql = "SELECT * FROM tb_Crediario_Parcela where CodParc=" & VGIntCodParc
        RecParc.Open StrSql, vgCon, 1, 3
        
        RecParc("Vencimento") = FormataDataUS(TxtDtVenc.Text)
        RecParc("Valor") = Mid(TxtValorParc.Text, 4)
        RecParc.Update
        
        RecParc.Close
        
        Desconecta
        
        VPStrBox = MsgBox("Alteração efetuada.", vbInformation, "Pró Ótica 2004 - Informação")

        FrmPrincipal.CmdPesqCred.Value = True
        
        Unload Me
        
        MDIPrincipal.Enabled = True
        MDIPrincipal.WindowState = 2
    
    End If
    
End Sub

Private Sub Form_Resize()
  FrmCrediario_Alt.Left = (MDIPrincipal.Width / 2) - (FrmCrediario_Alt.Width / 2)
  FrmCrediario_Alt.Top = (MDIPrincipal.Height / 3) - (FrmCrediario_Alt.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 3135
    Width = 6450
    Top = 1365
    Left = 3795
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Call MontaCbos
    
    Conecta
    
    Dim RecCred As New ADODB.Recordset
    Dim RecParc As New ADODB.Recordset
    
    StrSql = "SELECT CS.CodCredsta,CS.Nome FROM tb_Crediario as CR, tb_Crediarista as CS where CR.CodCredsta=CS.CodCredsta and CR.CodCred=" & VGIntCodCred
    RecCred.Open StrSql, vgCon, 1, 3

    StrSql = "SELECT Valor,Vencimento FROM tb_Crediario_Parcela where CodCred=" & VGIntCodCred & " and CodParc=" & VGIntCodParc
    RecParc.Open StrSql, vgCon, 1, 3
    
    CboCredsta.Text = RecCred.Fields.Item(1).Value & "                                                                           " & RecCred.Fields.Item(0).Value
    TxtValorParc.Text = FormataMoeda(RecParc.Fields.Item(0).Value)
    TxtDtVenc.Text = FormataData(RecParc.Fields.Item(1).Value)
    
    Desconecta
    
    MDIPrincipal.Enabled = False
End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecCredsta As New ADODB.Recordset
    
    StrSql = "SELECT CodCredsta,Nome FROM tb_Crediarista order by Nome"
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    Do While Not RecCredsta.EOF
        CboCredsta.AddItem (RecCredsta.Fields.Item(1).Value & "                                                                           " & RecCredsta.Fields.Item(0).Value)
        RecCredsta.MoveNext
    Loop
    
    Desconecta
    
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
    End If
End Sub
