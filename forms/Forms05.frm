VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Forms05 
   Caption         =   "Brame Leil�es - Relat�rios Anuais"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "Forms05.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Forms05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
  Sheets("Dados").Select
  Application.Run ("AtualizarTela0")
  Application.Run ("ReexibirPlanilhas")
End Sub
'------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
  Application.Run ("OcultarPlanilhas")
  Application.Run ("AtualizarTela1")
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
  Forms2.Show
End Sub
Private Sub VerificarVazio()
  If LeiLeandro = True Then
    Dados_Lei = "Leandro Dias Brame"
    Dados_Mat = 130
    Dados_Pro = "Leiloeiro P�blico"
    Dados_Esc = 1212
    Dados_Pre = "N�o Possui"
  ElseIf LeiTeresa = True Then
    Dados_Lei = "Maria Teresa Dias Brame"
    Dados_Mat = 31
    Dados_Pro = "Leiloeira P�blica"
    Dados_Esc = 1211
    Dados_Pre = "Luis Cerino de Almeida"
   
  ElseIf LeiLeandro = False And LeiTeresa = False Then
    MsgBox "Selecione um leiloeiro"
    Unload Me
    End
  End If
End Sub
