VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Forms03 
   Caption         =   "Brame Leil�es - Relat�rio de Leil�o"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "Forms03.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Forms03"
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
Private Sub SeVazio()
  If TextoEndere�o.Value = "" Or TextoMes.Value = "" Or TextoAno.Value = "" Then
     MsgBox ("O local do arquivo e o m�s s�o obrigat�rios.")
     Unload Me
     End
  End If
  
  If Bot_Excel = False And Bot_Word = False And Bot_PDF = False Then
    MsgBox ("Escolha um tipo de arquivo")
    Unload Me
    End
  End If
End Sub
'------------------------------------------------------------------------------
Private Sub ValoresVariaveis()
  Crit�rio_Mes = TextoMes.Value
  Crit�rio_Ano = TextoAno
  Local_Planilha = TextoEndere�o.Value
  
  If Bot_Excel = True Then
    Extens = "xlsx"
  ElseIf Bot_Word = True Then
    Extens = "docx"
  ElseIf Bot_PDF = True Then
    Extens = "pdf"
  End If
End Sub
'------------------------------------------------------------------------------
Private Sub ModeloDocumento()
  If Bot_Comunicado = True Then
    NomeDo = "Comunicado"
  ElseIf Bot_DREI = True Then
    NomeDo = "DREI"
  ElseIf Bot_Relat�rio = True Then
    NomeDo = "Relat�rio"
   End If
End Sub
'------------------------------------------------------------------------------
Private Sub Bot_Comunicado_Click()
  SeVazio
  ValoresVariaveis
  ModeloDocumento
  Unload Me
  
  If TodosMeses.Value = True Then
  Reprodu��oAnual
  End If
  
  If TodosMeses.Value = False Then
  Comunicado
  End If

End Sub
'------------------------------------------------------------------------------
Private Sub Bot_DREI_Click()
  SeVazio
  ValoresVariaveis
  ModeloDocumento
  Unload Me
  
  If TodosMeses.Value = True Then
  Reprodu��oAnual
  End If
  
  If TodosMeses.Value = False Then
  DREI
  End If
  
End Sub
'------------------------------------------------------------------------------
Private Sub Bot_Relat�rio_Click()
  SeVazio
  ValoresVariaveis
  ModeloDocumento
  Unload Me
  
  If TodosMeses.Value = True Then
  Reprodu��oAnual
  End If
  
  If TodosMeses.Value = False Then
  Relatorio
  End If
  
End Sub
