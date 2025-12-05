VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H80000004&
   Caption         =   "Lydians Ponto"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Registrar"
      Height          =   405
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOBS 
      Caption         =   "Observação"
      Height          =   400
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame fraNome 
      Caption         =   "Selecione seu nome "
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox cmbFuncionarios 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "1"
         Top             =   300
         Width           =   2895
      End
   End
   Begin VB.Frame fraOpcoes 
      Caption         =   "Opções"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
      Begin VB.ComboBox cmbOp 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ponto As clsOperacoes

Private Sub cmbFuncionarios_Click()
   cmbOp.Enabled = True
   cmbOp.BackColor = &H8000000E
End Sub

Private Sub cmdExcel_Click()
   On Error GoTo cmdExcel_E
    
   Dim i As Integer
   Dim j As Integer
   Dim DtHr As Date
   Dim Dt As Date
   Dim strDate
   Dim Hr As String
   Hr = CStr(Time)
   Dt = Date
   DtHr = Now
   
   Dim IndexFuncionario As Integer
   Dim IndexPonto As Integer
   Dim strOpPonto As String
   Dim strFuncionario As String
   Dim strFuncionarioPrimeiroNome As String
   IndexFuncionario = Validar_Retornar_Funcionario
   strFuncionario = Retornar_Funcionario
   IndexPonto = Validar_Retornar_Ponto
   strOpPonto = Retornar_Ponto
   
   Dim NmArq As String
   NmArq = cmbFuncionarios.Text & "-Ponto-"
   
   'Caracteres reservados - \/:*?"<> |
   Dim strCaractere As String
   strCaractere = ""
   
   'Formata data para XX-XX-XXXX
   For i = 1 To Len(CStr(Dt))
      strCaractere = Mid(CStr(Dt), i, 1)
      If IsNumeric(strCaractere) Then
         NmArq = NmArq & strCaractere
         strDate = strDate & strCaractere
      Else
         NmArq = NmArq & "-"
         strDate = strDate & "-"
      End If
   Next i

   strCaractere = ""
   strFuncionarioPrimeiroNome = ""
   
   'Pega o primeiro nome do user para UX
   For i = 1 To Len(cmbFuncionarios.Text)
      strCaractere = Mid(cmbFuncionarios.Text, i, 1)
      If strCaractere <> " " Then
         strFuncionarioPrimeiroNome = strFuncionarioPrimeiroNome & strCaractere
      Else
         Exit For
      End If
   Next i
   
   strCaractere = ""
   strFuncionario = ""
   
   'Formata o nome do funcionário sem espaços, pois é utilizado para criar o arquivo posteriormente
   For i = 1 To Len(cmbFuncionarios.Text)
      strCaractere = Mid(cmbFuncionarios.Text, i, 1)
      If strCaractere <> " " Then
         strFuncionario = strFuncionario & strCaractere
      End If
   Next i

   If IndexFuncionario < 0 Then
      MsgBox "Selecione um funcionário para registrar o ponto.", vbInformation, "Atenção"
      GoTo DestruirObjetos
   Else
      If IndexPonto >= 0 Then
         With Ponto
            If (.RegistrarPontoNaPlanilha(CStr(strDate), CStr(Time), CStr(cmbOp.Text), strFuncionario)) Then
               MsgBox "Ponto registrado! Verifique em caminho", vbInformation, "Sucesso"
      
               If strOpPonto = "Entrada" Then
                  MsgBox "Bem vindo(a), " & strFuncionarioPrimeiroNome & ". Bom serviço!", vbInformation, "Olá"
               
               ElseIf strOpPonto = "Intervalo - Entrada" Then
                  MsgBox "Bom intervalo, " & strFuncionarioPrimeiroNome & ". Até logo!", vbInformation, "Buenas"
               
               ElseIf strOpPonto = "Intervalo - Saída" Then
                  MsgBox "Quase lá! Boa tarde.", vbInformation, "Olá"
               
               Else
                  MsgBox "Bom descanso, " & strFuncionarioPrimeiroNome & ".", vbInformation, "Bye Bye"
               End If
               GoTo DestruirObjetos
            Else
               MsgBox "Ponto não foi registrado! Utilize o cartão ponto para registrar sua batida.", vbExclamation, "Erro"
               MsgBox "CONTATE O SUPORTE: CONTATO" & vbNewLine & vbNewLine & "Descrição do erro: FORM - VALIDAÇÃO obj.CriarPlanilhaExcel(CStr(strDate), CStr(Time), CStr(cmbOp.Text), strFunc).", 16, "Erro"
               GoTo DestruirObjetos
            End If
         End With
      Else
         MsgBox "Selecione uma opção para registrar o ponto.", vbInformation, "Atenção"
         GoTo DestruirObjetos
      End If
   End If
   
'   Dim var As String
'   If Ponto.valorSubstituido Then
'      var = InputBox("Você substituiu seu ponto de " & strOpPonto & "." & " Você está substituindo valorDoPontoAtual por ponto substituido. Quer continuar?")
'   Else
'   End If
   
cmdExcel_E:
   MsgBox "Ponto não foi registrado! Utilize o cartão ponto para registrar sua batida.", vbExclamation, "Erro"
   MsgBox "CONTATE O SUPORTE: CONTATO" & vbNewLine & vbNewLine & "Descrição do erro: FORM - frmInicio.", 16, "Erro"
   
DestruirObjetos:
'Limpar objs
End Sub

Private Sub Form_Load()
   Set Ponto = New clsOperacoes
   cmbFuncionarios.AddItem "Brayan Coimbra"
   cmbFuncionarios.AddItem "Rodrigo Conte"
   cmbFuncionarios.AddItem "Wellington Oliveira"
   cmbFuncionarios.AddItem "Funcionario1"
   cmbFuncionarios.AddItem "Funcionario2"
   
   cmbOp.AddItem "Entrada"
   cmbOp.AddItem "Intervalo - Entrada"
   cmbOp.AddItem "Intervalo - Saída"
   cmbOp.AddItem "Saída"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Ponto = Nothing
End Sub

Private Function Validar_Retornar_Funcionario() As Integer
   Dim i As Integer
   
   For i = 0 To cmbFuncionarios.ListCount
      If cmbFuncionarios.List(i) = cmbFuncionarios.Text Then
         Validar_Retornar_Funcionario = cmbFuncionarios.ListIndex
         Exit Function
      End If
   Next i
   
End Function

Private Function Retornar_Funcionario() As String
   Retornar_Funcionario = cmbFuncionarios.Text
End Function

Private Function Retornar_Ponto() As String
   Retornar_Ponto = cmbOp.Text
End Function

Private Function Validar_Retornar_Ponto() As Integer
   Dim i As Integer
   
   For i = 0 To cmbOp.ListCount
      If cmbOp.List(i) = cmbOp.Text Then
         Validar_Retornar_Ponto = cmbOp.ListIndex
         Exit Function
      End If
   Next i
End Function
