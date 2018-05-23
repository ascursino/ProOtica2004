VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmListaArquivos 
   Caption         =   "Lista de arquivos"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
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
   Icon            =   "FrmListaArquivos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   4785
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
      Top             =   3600
      Width           =   4575
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
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   120
         OleObjectBlob   =   "FrmListaArquivos.frx":0CCA
         Top             =   120
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.FileListBox FileList 
         Height          =   1260
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   4335
      End
      Begin VB.DriveListBox DriveRaiz 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   4335
      End
      Begin VB.DirListBox DriveList 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4335
      End
   End
End
Attribute VB_Name = "FrmListaArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String

Private Sub CmdFechar_Click()
    Unload Me
    'MDIPrincipal.Enabled = True
    FrmAssinaturaCarne.Enabled = True
End Sub

Private Sub DriveList_Change()
    FileList.Path = DriveList.Path
End Sub

Private Sub DriveRaiz_Change()
On Error GoTo Erro
   DriveList.Path = DriveRaiz.Drive

Erro:
    If Err.Number = 68 Then
        VPStrBox = MsgBox("Insira um disco no drive.", vbInformation, "Pró Ótica 2004 - Informação")
    End If
End Sub

Private Sub FileList_DblClick()
    FrmAssinaturaCarne.TxtLogo.Text = FileList.Path & "\" & FileList.FileName
End Sub

Private Sub Form_Resize()
  FrmListaArquivos.Left = (MDIPrincipal.Width / 2) - (FrmListaArquivos.Width / 2)
  FrmListaArquivos.Top = (MDIPrincipal.Height / 3) - (FrmListaArquivos.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4950
    Width = 4905
    Top = 1725
    Left = 4965
    
    'MDIPrincipal.Enabled = False
    
    FrmAssinaturaCarne.Enabled = False
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
End Sub
