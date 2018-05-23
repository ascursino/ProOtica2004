VERSION 5.00
Object = "{9C152BB9-D77B-11D7-A6B5-00D009F8C11B}#3.0#0"; "shlock.ocx"
Begin VB.Form FrmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4005
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSplash.frx":0CCA
   ScaleHeight     =   4005
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   5400
      Top             =   120
   End
   Begin SHLock.SHLocker SHLocker1 
      Left            =   4920
      Top             =   1320
      _ExtentX        =   1032
      _ExtentY        =   979
      SenhaProg       =   "ProOtica2004"
   End
   Begin VB.Label LblVersao 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Versão"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Private Sub Form_Load()
'    LblVersao.Caption = "Versão " & App.Major & "." & App.Minor & "." & App.Revision
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If SHLocker1.SouRegistrado = False Then
'       If (SHLocker1.DiasParaSerTestado - SHLocker1.DiasQueUsei) <= 0 Then
'            VGStrLocker = "sim"
'            VGStrBox = MsgBox("O tempo de avaliação deste software expirou.", vbCritical, "Software expirou")
'            FrmLocker.Show
'            Unload Me
'        Else
'            FrmLocker.Show
'            Unload Me
'        End If
'
'    ElseIf SHLocker1.SouRegistrado = True Then
'        MDIPrincipal.Show
'        Unload Me
'
'    End If
'End Sub
'
'Private Sub Timer1_Timer()
'    Screen.MousePointer = vbHourglass
'
'    Unload Me
'
'    If SHLocker1.SouRegistrado = True Then
'       If (SHLocker1.DiasParaSerTestado - SHLocker1.DiasQueUsei) <= 0 Then
'            VGStrLocker = "sim"
'            VGStrBox = MsgBox("O tempo de avaliação deste software expirou.", vbCritical, "Software expirou")
'            FrmLocker.Show
'            Unload Me
'        Else
'            FrmLocker.Show
'            Unload Me
'        End If
'
'    ElseIf SHLocker1.SouRegistrado = False Then
'        MDIPrincipal.Show
'        Unload Me
'
'    End If
'End Sub
'

Option Explicit
Public VPStrTempo As String

Private Sub Form_Load()
    'Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    'Skin1.ApplySkin (Me.hwnd)
    
    LblVersao.Caption = "Versão " & App.Major & "." & App.Minor
    
    VPStrTempo = "sim"
End Sub

Private Sub Timer1_Timer()
    If VPStrTempo = "sim" Then
        Screen.MousePointer = vbHourglass
        
        'Unload Me
        
        If SHLocker1.SouRegistrado = False Then
           'If SHLocker1.DiasParaSerTestado - SHLocker1.DiasQueUsei <= 0 Then
           If SHLocker1.DiasQueUsei > SHLocker1.DiasParaSerTestado Then
                VGStrLocker = "sim"
        
                VGStrBox = MsgBox("O tempo de avaliação deste software expirou.", vbCritical, "Pró Vendas 2004 - Software expirou")
        
                FrmLocker.Show
                Unload Me
            Else
                FrmLocker.Show
                Unload Me
            End If
        
        ElseIf SHLocker1.SouRegistrado = True Then
            MDIPrincipal.Show
            Unload Me
        
        End If
        VPStrTempo = ""
    End If
End Sub

