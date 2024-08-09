VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Forms01 
   Caption         =   "Brame Leilões - Início"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "Forms01.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Forms01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
  Sheets("Dados").Select
End Sub
'------------------------------------------------------------------------------
Private Sub Bot_Cancelar_Click()
  Unload Me
  End
End Sub
'------------------------------------------------------------------------------
Private Sub Bot_Proximo_Click()
  VerificarVazio
  Unload Me
  Forms02.Show
End Sub
Private Sub VerificarVazio()
  If LeiLeandro = True Then
    Dados_Lei = "Leandro Dias Brame"
    Dados_Mat = 130
    Dados_Pro = "Leiloeiro Público"
    Dados_Esc = 1212
    Dados_Pre = "Não Possui"
  ElseIf LeiTeresa = True Then
    Dados_Lei = "Maria Teresa Dias Brame"
    Dados_Mat = 31
    Dados_Pro = "Leiloeira Pública"
    Dados_Esc = 1211
    Dados_Pre = "Luis Cerino de Almeida"
   
  ElseIf LeiLeandro = False And LeiTeresa = False Then
    MsgBox "Selecione um leiloeiro"
    Unload Me
    End
  End If
End Sub
