VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Pró Ótica 2004 - Sistema Integrado de Gestão de Ótica"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10545
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3360
      OleObjectBlob   =   "MDIPrincipal.frx":0CCA
      Top             =   960
   End
   Begin VB.Menu MNUSobre 
      Caption         =   "&Sobre"
   End
   Begin VB.Menu MNUHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    Dim VLStrRegistro As String
    
    If FrmSplash.SHLocker1.SouRegistrado = True Then
        VLStrRegistro = "Registrado"
    Else
        VLStrRegistro = "Trial"
    End If
    
    Screen.MousePointer = vbNormal
    Me.Caption = "Pró Ótica 2004 - Sistema Integrado de Gestão de Ótica - V." & App.Major & "." & App.Minor & "." & App.Revision & " (" & VLStrRegistro & ")"
    
    Skin1.LoadSkin (App.Path & "\ProOtica2004.skn")
    Skin1.ApplySkin (Me.hwnd)

    FrmPrincipal.Show

End Sub

Private Sub MDIForm_Terminate()
    Unload Me
End Sub

Private Sub MNUHelp_Click()
    Dim i&
    i& = ShellExecute(0, "open", App.Path & "\ProOtica2004.chm", "", App.Path, SW_SHOW)
End Sub

Private Sub MNUSobre_Click()
    frmSistema.Show
End Sub
