VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWebScraping 
   Caption         =   "Web Scraping - Cotação das Principais Moedas"
   ClientHeight    =   8640.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900.001
   OleObjectBlob   =   "CotacoesWebScraping.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWebScraping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim plan As Worksheet
Dim dolarAtual As String
Dim euroAtual As String
Dim libraAtual As String
Dim ouroAtual As String
Dim bitcoinAtual As String

Private Sub btnDolarAtualizar_Click()
    
    On Error GoTo trata_Erro
    
    Dim site As String, htmlSite As String, resumoHTML As String
    Dim requisicao As Object
    
    Set requisicao = CreateObject("MSXML2.XMLHTTP.6.0")
    
    site = "https://dolarhoje.com/"
    
    requisicao.Open "get", site, False
    
    requisicao.Send
    
    htmlSite = requisicao.responsetext
    
    resumoHTML = Mid(htmlSite, InStr(htmlSite, "cotMoeda nacional"), 100)
    resumoHTML = Mid(resumoHTML, InStr(resumoHTML, "value="), 11)
    dolarAtual = Right(resumoHTML, 4)
     
    lbDolarScraping.Caption = VBA.Format(dolarAtual, "currency")
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString
    
    Exit Sub
    
trata_Erro:
    MsgBox "Não foi possível recuperar o valor, verifique sua conexão com a internet e tente novamente"
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString
    
End Sub

Private Sub btnDolarInserir_Click()
    
    On Error GoTo trata_Erro
    
    Set plan = Sheets(1)
    
    plan.Select
    
    plan.Range(ActiveCell.Address).Value = VBA.CCur(dolarAtual)

    Exit Sub
    
trata_Erro:
    MsgBox "Erro ao inserir o valor!"
    
End Sub

Private Sub btnEuroAtualizar_Click()

    On Error GoTo trata_Erro
    
    Dim site As String, htmlSite As String, resumoHTML As String
    Dim requisicao As Object
    
    Set requisicao = CreateObject("MSXML2.XMLHTTP.6.0")
    
    site = "https://dolarhoje.com/euro-hoje/"
    
    requisicao.Open "get", site, False
    
    requisicao.Send
    
    htmlSite = requisicao.responsetext
    
    resumoHTML = Mid(htmlSite, InStr(htmlSite, "cotMoeda nacional"), 100)
    resumoHTML = Mid(resumoHTML, InStr(resumoHTML, "value="), 11)
    euroAtual = Right(resumoHTML, 4)
     
    lbEuroScraping.Caption = VBA.Format(euroAtual, "currency")
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString

 Exit Sub
    
trata_Erro:
    MsgBox "Não foi possível recuperar o valor, verifique sua conexão com a internet e tente novamente"
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString
    
End Sub

Private Sub btnEuroInserir_Click()
    
    On Error GoTo trata_Erro
    
    Set plan = Sheets(1)
    
    plan.Select
    
    plan.Range(ActiveCell.Address).Value = VBA.CCur(euroAtual)
    
    Exit Sub
    
trata_Erro:
    MsgBox "Erro ao inserir o valor!"
    
End Sub

Private Sub btnLibraAtualizar_Click()

    On Error GoTo trata_Erro
    
    Dim site As String, htmlSite As String, resumoHTML As String
    Dim requisicao As Object
    
    Set requisicao = CreateObject("MSXML2.XMLHTTP.6.0")
    
    site = "https://dolarhoje.com/libra-hoje/"
    
    requisicao.Open "get", site, False
    
    requisicao.Send
    
    htmlSite = requisicao.responsetext
    
    resumoHTML = Mid(htmlSite, InStr(htmlSite, "cotMoeda nacional"), 100)
    resumoHTML = Mid(resumoHTML, InStr(resumoHTML, "value="), 11)
    libraAtual = Right(resumoHTML, 4)
     
    lbLibraScraping.Caption = VBA.Format(libraAtual, "currency")
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString

 Exit Sub
    
trata_Erro:
    MsgBox "Não foi possível recuperar o valor, verifique sua conexão com a internet e tente novamente"
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString

End Sub

Private Sub btnLibraInserir_Click()

    On Error GoTo trata_Erro
    
    Set plan = Sheets(1)
    
    plan.Select
    
    plan.Range(ActiveCell.Address).Value = VBA.CCur(libraAtual)

Exit Sub
    
trata_Erro:
    MsgBox "Erro ao inserir o valor!"
    
End Sub

Private Sub btnOuroAtualizar_Click()

    On Error GoTo trata_Erro
    
    Dim site As String, htmlSite As String, resumoHTML As String
    Dim requisicao As Object
    
    Set requisicao = CreateObject("MSXML2.XMLHTTP.6.0")
    
    site = "https://dolarhoje.com/ouro-hoje/"
    
    requisicao.Open "get", site, False
    
    requisicao.Send
    
    htmlSite = requisicao.responsetext
    
    resumoHTML = Mid(htmlSite, InStr(htmlSite, "cotMoeda nacional"), 100)
    resumoHTML = Mid(resumoHTML, InStr(resumoHTML, "value="), 11)
    ouroAtual = Right(resumoHTML, 4)
     
    lbOuroScraping.Caption = VBA.Format(ouroAtual, "currency")
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString
 
 Exit Sub
    
trata_Erro:
    MsgBox "Não foi possível recuperar o valor, verifique sua conexão com a internet e tente novamente"
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString

End Sub

Private Sub btnOuroInserir_Click()

    On Error GoTo trata_Erro
    
    Set plan = Sheets(1)
    
    plan.Select
    
    plan.Range(ActiveCell.Address).Value = VBA.CCur(ouroAtual)

Exit Sub
    
trata_Erro:
    MsgBox "Erro ao inserir o valor!"
    
End Sub

Private Sub btnBitcoinAtualizar_Click()

    On Error GoTo trata_Erro
    
    Dim site As String, htmlSite As String, resumoHTML As String
    Dim requisicao As Object
    
    Set requisicao = CreateObject("MSXML2.XMLHTTP.6.0")
    
    site = "https://dolarhoje.com/bitcoin-hoje/"
    
    requisicao.Open "get", site, False
    
    requisicao.Send
    
    htmlSite = requisicao.responsetext
    
    resumoHTML = Mid(htmlSite, InStr(htmlSite, "cotMoeda nacional"), 100)
    resumoHTML = Mid(resumoHTML, InStr(resumoHTML, "value="), 11)
    bitcoinAtual = Right(resumoHTML, 4)
     
    lbBitcoinScraping.Caption = VBA.Format(bitcoinAtual, "currency")
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString
    
     Exit Sub
    
trata_Erro:
    MsgBox "Não foi possível recuperar o valor, verifique sua conexão com a internet e tente novamente"
    
    Set requisicao = Nothing
    site = vbNullString
    htmlSite = vbNullString
    resumoHTML = vbNullString
    
End Sub

Private Sub btnBitcoinInserir_Click()
    
    On Error GoTo trata_Erro
    
    Set plan = Sheets(1)
    
    plan.Select
    
    plan.Range(ActiveCell.Address).Value = VBA.CCur(ouroAtual)

Exit Sub
    
trata_Erro:
    MsgBox "Erro ao inserir o valor!"
    
End Sub

Private Sub UserForm_Initialize()

    Call btnDolarAtualizar_Click
    Call btnEuroAtualizar_Click
    Call btnLibraAtualizar_Click
    Call btnOuroAtualizar_Click
    Call btnBitcoinAtualizar_Click
End Sub
