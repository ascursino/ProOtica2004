Attribute VB_Name = "ModuloConecta"
Public Const SW_SHOW As Long = 5
Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory _
        As String, ByVal nShowCmd As Long) As Long

Global vgCon As New ADODB.Connection 'variável de conexão
Global StrSql As String 'variável da string SQL

Global VGIntCodCli As Long
Global VGIntCodRec As Long
Global VGIntCodMed As Long
Global VGIntCodForn As Long
Global VGIntCodEst As Long
Global VGIntCodProd As Long
Global VGIntCodParc As Long
Global VGIntCodCred As Long
Global VGIntCodCredsta As Long
Global VGIntCodCredstaVenda As Long
Global VGIntTotalCred As Long
Global VGIntCodCx As Long
Global VGIntCodPagar As Long
Global VGIntCodReceber As Long
Global VGIntCodVend As Long
Global VGIntCodOrc As Long
Global VGIntCodVenda As Long
Global VGIntCodVendaRel As Long
Global VGIntCodCredTemp As Long
Global VGIntPropCodCred As Long

Global VGStrLocker As String
Global VGStrBox As String
Global VGStrNomeCredsta As String
Global VGStrTVenda As String
Global VGStrTCred As String
Global VGStrTDeb As String
Global VGStrTMov As String
Global VGStrStatusPagto As String
Global VGStrStatusReceb As String
Global VGStrDescrProd As String
Global VGStrTipoProd As String
Global VGStrCredLista As String
Global VGStrClienteRel As String
Global VGStrProposta As String
Global VGStrAssinatura As String
Global VGStrAssinaturaCarne As String
Global VGStrAssinaturaOrc As String
Global VGStrAssinaturaProp As String
Global VGStrAssinaturaProposta As String
Global VGStrCredsta As String
Global VGStrVendaRapida As String
Global VGStrEntrar As String

Global VGStrEstoqueIncExtra As String
Global VGStrNomeCli As String
Global VGStrForm As String
Global VGStrIncluirProd As String

Global VGStrBanco01 As String
Global VGStrChequeDig01 As String
Global VGStrData01 As String
Global VGStrValor01 As String

Global VGStrBanco02 As String
Global VGStrChequeDig02 As String
Global VGStrData02 As String
Global VGStrValor02 As String

Global VGStrBanco03 As String
Global VGStrChequeDig03 As String
Global VGStrData03 As String
Global VGStrValor03 As String

Global VGStrBanco04 As String
Global VGStrChequeDig04 As String
Global VGStrData04 As String
Global VGStrValor04 As String

Global VGStrBanco05 As String
Global VGStrChequeDig05 As String
Global VGStrData05 As String
Global VGStrValor05 As String

Global VGStrBanco06 As String
Global VGStrChequeDig06 As String
Global VGStrData06 As String
Global VGStrValor06 As String

Global VGStrBanco07 As String
Global VGStrChequeDig07 As String
Global VGStrData07 As String
Global VGStrValor07 As String

Global VGStrBanco08 As String
Global VGStrChequeDig08 As String
Global VGStrData08 As String
Global VGStrValor08 As String

Global VGStrBanco09 As String
Global VGStrChequeDig09 As String
Global VGStrData09 As String
Global VGStrValor09 As String

Global VGStrBanco10 As String
Global VGStrChequeDig10 As String
Global VGStrData10 As String
Global VGStrValor10 As String

Public Function Decipher(ByVal from_text As String) As String

Const MIN_ASC = 32  ' Space.
Const MAX_ASC = 126 ' ~.
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer

    ' Initialize the random number generator.
    offset = 123
    Rnd -1
    Randomize offset

    ' Encipher the string.
    str_len = Len(from_text)
    For i = 1 To str_len
        ch = Asc(Mid$(from_text, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            Decipher = Decipher & Chr$(ch)
        End If
    Next i

End Function

Sub Conecta()       'Conecta com o banco prootica.mdb
    Dim vlFso, vlArquivo
    Dim vgStr As String
    
    '--- Abre arquivo criptografado com a string de conexão
    ''Set vlFso = CreateObject("Scripting.FileSystemObject")
    ''Set vlArquivo = vlFso.OpenTextFile(App.Path & "\prootica2004.dll", ForReading, False)
    
    '--- Descriptografa string de conexão
    ''vgStr = Decipher(vlArquivo.ReadLine)
    
    '--- Configura o tempo de pesquisa
    ''vgCon.ConnectionTimeout = 130
    
    '--- MontaString de conexão
    ''vgStr = "DBQ=" & App.Path & "\prootica2004.mdb;" & _
    ''    "Driver={Microsoft Access Driver (*.mdb)};"
    vgStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\prootica2004.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    
    '--- Abre conexão com SQL
    vgCon.Open vgStr
End Sub

Sub Desconecta()
    vgCon.Close
End Sub
