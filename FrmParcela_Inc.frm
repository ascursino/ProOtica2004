VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Begin VB.Form FrmParcela_Inc 
   Caption         =   "Inclusão de parcelas do crediário"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
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
   Icon            =   "FrmParcela_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   9090
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
      Top             =   3360
      Width           =   8895
      Begin VB.CommandButton CmdOK 
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
         Left            =   6240
         TabIndex        =   1
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   960
         OleObjectBlob   =   "FrmParcela_Inc.frx":0CCA
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
         Left            =   7560
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
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8895
      Begin FPSpread.vaSpread GridParcela 
         Height          =   2895
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8655
         _Version        =   393216
         _ExtentX        =   15266
         _ExtentY        =   5106
         _StockProps     =   64
         ColHeaderDisplay=   0
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   5
         MaxRows         =   1
         Protect         =   0   'False
         RowHeaderDisplay=   2
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmParcela_Inc.frx":0EFE
         UserResize      =   1
      End
   End
End
Attribute VB_Name = "FrmParcela_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    Unload Me
   
    FrmVenda_Inc.Enabled = True
End Sub

Private Sub CmdOK_Click()
    Dim VLIntLinha As Integer
    Dim VLIntLinhaMax As Integer

    VLIntLinha = 1
    
    Do While VLIntLinha <= GridParcela.MaxRows
                
        GridParcela.Row = VLIntLinha
        
        If VLIntLinha = 1 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco01 = GridParcela.Text
            Else
                VGStrBanco01 = 0
            End If
            
            GridParcela.Col = 2
            VGStrChequeDig01 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig01 = VGStrChequeDig01 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData01 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor01 = GridParcela.Text
            
        ElseIf VLIntLinha = 2 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco02 = GridParcela.Text
            Else
                VGStrBanco02 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig02 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig02 = VGStrChequeDig02 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData02 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor02 = GridParcela.Text
        
        ElseIf VLIntLinha = 3 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco03 = GridParcela.Text
            Else
                VGStrBanco03 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig03 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig03 = VGStrChequeDig03 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData03 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor03 = GridParcela.Text
        
        ElseIf VLIntLinha = 4 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco04 = GridParcela.Text
            Else
                VGStrBanco04 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig04 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig04 = VGStrChequeDig04 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData04 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor04 = GridParcela.Text
        
        ElseIf VLIntLinha = 5 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco05 = GridParcela.Text
            Else
                VGStrBanco05 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig05 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig05 = VGStrChequeDig05 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData05 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor05 = GridParcela.Text
        
        ElseIf VLIntLinha = 6 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco06 = GridParcela.Text
            Else
                VGStrBanco06 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig06 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig06 = VGStrChequeDig06 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData06 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor06 = GridParcela.Text
        
        ElseIf VLIntLinha = 7 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco07 = GridParcela.Text
            Else
                VGStrBanco07 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig07 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig07 = VGStrChequeDig07 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData07 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor07 = GridParcela.Text
        
        ElseIf VLIntLinha = 8 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco08 = GridParcela.Text
            Else
                VGStrBanco08 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig08 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig08 = VGStrChequeDig08 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData08 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor08 = GridParcela.Text
        
        ElseIf VLIntLinha = 9 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco09 = GridParcela.Text
            Else
                VGStrBanco09 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig09 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig09 = VGStrChequeDig09 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData09 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor09 = GridParcela.Text
        
        ElseIf VLIntLinha = 10 Then
            GridParcela.Col = 1
            If GridParcela.Text = "" Then
                VGStrBanco10 = GridParcela.Text
            Else
                VGStrBanco10 = 0
            End If
        
            GridParcela.Col = 2
            VGStrChequeDig10 = GridParcela.Text
        
            GridParcela.Col = 3
            If GridParcela.Text <> "" Then
                VGStrChequeDig10 = VGStrChequeDig10 & "-" & GridParcela.Text
            End If
            
            GridParcela.Col = 4
            VGStrData10 = GridParcela.Text
            
            GridParcela.Col = 5
            VGStrValor10 = GridParcela.Text
        
        End If
        
        VLIntLinha = VLIntLinha + 1
    Loop
    
    Unload Me
    
    FrmVenda_Inc.Enabled = True
    
End Sub

Private Sub Form_Resize()
  FrmParcela_Inc.Left = (MDIPrincipal.Width / 2) - (FrmParcela_Inc.Width / 2)
  FrmParcela_Inc.Top = (MDIPrincipal.Height / 3) - (FrmParcela_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 4695
    Width = 9210
    Top = 1500
    Left = 1860
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    FrmVenda_Inc.Enabled = False
    
    Dim VLIntLinha As Integer
    Dim VLIntLinhaMax As Integer
    
    VLIntLinhaMax = FrmVenda_Inc.CboPrazoChqParc.Text
    VLIntLinha = 1
    
    Do While VLIntLinha <= VLIntLinhaMax
                
        GridParcela.Row = VLIntLinha
        GridParcela.Col = 0
        GridParcela.Text = FormataNum(VLIntLinha) & "ª parcela"
        
        VLIntLinha = VLIntLinha + 1
        GridParcela.MaxRows = GridParcela.MaxRows + 1
    Loop
    GridParcela.MaxRows = GridParcela.MaxRows - 1
End Sub

