VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public valorSubstituido As Boolean

Public Function Registrar(Caminho As String, nomeArq As String, registroNoArq As String) As Boolean

   On Error GoTo Registrar_E

   Registrar = False
   
   Dim fso As FileSystemObject

   'cria o objeto fso do tipo FileSystemObject
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   'verifica se o diretório existe , senão existir então cria
   If Not (fso.FolderExists(Caminho)) Then
       fso.CreateFolder (Caminho)
   End If
      
   Dim numArquivo As Integer
   numArquivo = FreeFile()
   
   Open Caminho & "\" & nomeArq & ".xlsx" For Append As #numArquivo
   Print #numArquivo, registroNoArq
   Close #numArquivo
    
   Registrar = True

   GoTo DestruirObjetos

Registrar_E:
   Registrar = False
   
DestruirObjetos:
End Function

Public Function RegistrarPontoNaPlanilha(strData As String, strHora As String, strRegistroPonto As String, strFunc As String) As Boolean
   
   On Error GoTo RegistrarPontoNaPlanilha_E
   
   RegistrarPontoNaPlanilha = False
'
'   CriarArquivoExcel strFunc
   Dim objExcel As Object
   Dim objPlanilha As Object
   Dim ArqCriado As Boolean
   Dim fso As Object

   ' Criar objeto fso do tipo FileSystemObject para manipular arquivos
   Set fso = CreateObject("Scripting.FileSystemObject")

   ' Verificar se o arquivo existe
   If Not (fso.FileExists("C:\Users\brayancoimbra\Ponto\" & strFunc & ".xlsx")) Then
       ' Arquivo não existe, criar
       Set objExcel = CreateObject("Excel.Application")
       Set objWorkbook = objExcel.Workbooks.Add
       objWorkbook.SaveAs "C:\Users\brayancoimbra\Ponto\" & strFunc & ".xlsx"
       objWorkbook.Close False
       objExcel.Quit
       ArqCriado = True
   End If

'    Inicializar o Excel
   Set objExcel = CreateObject("Excel.Application")
'   objExcel.Visible = True ' Tornar o Excel visível

'    Abre a planilha do respectivo funcionario
   Set objPlanilha = objExcel.Workbooks.Open("C:\Users\brayancoimbra\Ponto\" & strFunc & ".xlsx").Sheets(1)

'   Se criar um novo arquivo, adiciona os valores abaixo para permitir a inserção de novos registros posteriormente
   If ArqCriado Then
      objPlanilha.Cells(1, 1).Value = "DIA"
      objPlanilha.Cells(1, 2).Value = "Entrada"
      objPlanilha.Cells(1, 3).Value = "Intervalo - Entrada"
      objPlanilha.Cells(1, 4).Value = "Intervalo - Saída"
      objPlanilha.Cells(1, 5).Value = "Saída"
      objPlanilha.Cells(1, 6).Value = "Observação"
   End If

   Dim dataInserida As Boolean
   
   ' Verificar se a data já está presente
   For i = 2 To 500
      If objPlanilha.Cells(i, 1).Value = strData Then
         dataInserida = True
         Exit For
      End If
   Next i

   ' Se a data não estiver presente, insira-a na próxima célula vazia da respectiva coluna (DIA)
   If dataInserida = False Then
      For j = 2 To 500
         If objPlanilha.Cells(j, 1).Value = "" Then
            objPlanilha.Cells(j, 1).Value = strData
            Exit For
         End If
      Next j
   Else
   End If

   ' Verificar se o registro de ponto está presente nas colunas 2 a 5
   If dataInserida = False Then
      For k = 2 To 5
         If objPlanilha.Cells(1, k).Value = strRegistroPonto Then
            ' Procurar por uma célula vazia na coluna correspondente e inserir a hora
            For j = 2 To 500
               If objPlanilha.Cells(j, k).Value = "" Then
                  objPlanilha.Cells(j, k).Value = strHora
                  Exit For
               End If
            Next j
         End If
      Next k
   Else
      ' Verificar se o registro de ponto está presente nas colunas 2 a 5 e se a data já estiver inserida, substitui o registro do ponto atual pelo novo
      For k = 2 To 5
         If objPlanilha.Cells(1, k).Value = strRegistroPonto Then
            ' Procurar por uma célula vazia na coluna correspondente e inserir a hora
            For j = 2 To 500
               If objPlanilha.Cells(j, 1).Value = strData Then
                  objPlanilha.Cells(j, k).Value = strHora
                  valorSubstituido = True
                  Exit For
               End If
            Next j
         End If
      Next k
   End If

   'Salvar e fechar o Excel
   'objExcel.Workbooks("C:\Teste\Teste\PTeste.xlsx").Close SaveChanges:=True
   objExcel.ActiveWorkbook.Save
   objExcel.Quit
   
   RegistrarPontoNaPlanilha = True
   
   GoTo DestruirObjetos

RegistrarPontoNaPlanilha_E:
   RegistrarPontoNaPlanilha = False
   
DestruirObjetos:
   ' Limpar a referência aos objetos
   Set objPlanilha = Nothing
   Set objExcel = Nothing
   Set fso = Nothing
   
End Function

'Private Sub CriarArquivoExcel(strFunc As String)
'
''   On Error GoTo CriarArquivoExcel_E
'
'   Dim objExcel As Object
'   Dim objPlanilha As Object
'   Dim ArqCriado As Boolean
'   Dim fso As Object
'
'   ' Criar objeto fso do tipo FileSystemObject
'   Set fso = CreateObject("Scripting.FileSystemObject")
'
'   ' Verificar se o arquivo existe
'   If Not (fso.FileExists("C:\Users\brayancoimbra\Ponto\" & strFunc & ".xlsx")) Then
'       ' Arquivo não existe, criar
'       Set objExcel = CreateObject("Excel.Application")
'       Set objWorkbook = objExcel.Workbooks.Add
'       objWorkbook.SaveAs "C:\Users\brayancoimbra\Ponto\" & strFunc & ".xlsx"
'       objWorkbook.Close False
'       objExcel.Quit
'       ArqCriado = True
'   End If
'
'   ' Inicializar o Excel
'   Set objExcel = CreateObject("Excel.Application")
'   'objExcel.Visible = True ' Tornar o Excel visível
'
'   ' Abre a planilha do respectivo funcionario
'   Set objPlanilha = objExcel.Workbooks.Open("C:\Users\brayancoimbra\Ponto\" & strFunc & ".xlsx").Sheets(1)
'
'   'Se criar um novo arquivo, adiciona os valores abaixo para permitir a inserção de novos registros posteriormente
'   If ArqCriado Then
'      objPlanilha.Cells(1, 1).Value = "DIA"
'      objPlanilha.Cells(1, 2).Value = "Entrada"
'      objPlanilha.Cells(1, 3).Value = "Intervalo - Entrada"
'      objPlanilha.Cells(1, 4).Value = "Intervalo - Saída"
'      objPlanilha.Cells(1, 5).Value = "Saída"
'   End If
'
'DestruirObjetos:
''Não limpa, pois vai usar no função principal
'End Sub
