## Leitura de Dados para geração de Indicadores

```
### Leitura dos Dados em uma Aba do EXCEL para geração de Indicadores em uma nova Aba
Dim Aba As String
Dim CelMes, CelAno, linha, NumDataVazia As Integer
Dim strMes, strAno, strData, strData_anterior As String
Dim Coluna, strValorContato, ColunaContato As String
Dim NumRow As Double
Dim Response
Dim NumTotalLinhas As Double
Dim PrimeiraLinha As Integer
Dim ContNumeros, ContJaAgendados, ContAgendados, ContSemDireito As Integer
Dim ContSemInteresse, ContEngano, ContNaoAtende, ContRetornar As Integer
Dim ContCliente, ContDoenca As Integer
Dim strValor, strValor_anterior As String
Dim array_Data(1 To 31) As String
Dim array_numeros(1 To 31) As Integer
Dim array_ja_agendados(1 To 31) As Integer
Dim array_agendados(1 To 31) As Integer
Dim array_sem_direto(1 To 31) As Integer
Dim array_sem_interesse(1 To 31) As Integer
Dim array_engano(1 To 31) As Integer
Dim array_nao_atende(1 To 31) As Integer
Dim array_retornar(1 To 31) As Integer
Dim array_cliente(1 To 31) As Integer
Dim array_doenca(1 To 31) As Integer

Dim Total_Numeros, Total_agendados, Total_JaAgendados As Integer
Dim Total_sem_direto, Total_sem_interesse, Total_engano As Integer
Dim Total_nao_atende, Total_retornar As String
Dim Total_cliente, Total_doenca As String

'Alinha os dados no centro das células
Worksheets("INDICADORES").Range("B8:L39").HorizontalAlignment = xlCenter

'Apaga o conteudo das células
Worksheets("INDICADORES").Range("B8:L38").ClearContents
Worksheets("INDICADORES").Range("C9:L39").ClearContents

CelMes = Sheets("INDICADORES").Range("C4") 'Célula Mês
CelAno = Sheets("INDICADORES").Range("E4") 'Célula Ano

'Valida se o campo Mês (Célula C6) está vazia
If IsEmpty(CelMes) Or CelMes = 0 Then
  Response = MsgBox("Favor informar um MÊS" + Chr(13) & Chr(10) + "Valores possíveis: 1 a 12", vbCritical + vbOKOnly, "ATENÇÃO")
  Exit Sub
End If

'Valida se o campo Ano (Célula E6) está vazia
If IsEmpty(CelAno) Or CelAno = 0 Then
  Response = MsgBox("Favor informar um ANO", vbCritical + vbOKOnly, "ATENÇÃO")
  Exit Sub
End If

If CelMes = 1 Then
   strMes = "JAN"
ElseIf CelMes = 2 Then
   strMes = "FEV"
ElseIf CelMes = 3 Then
   strMes = "MAR"
ElseIf CelMes = 4 Then
   strMes = "ABR"
ElseIf CelMes = 5 Then
   strMes = "MAI"
ElseIf CelMes = 6 Then
   strMes = "JUN"
ElseIf CelMes = 7 Then
   strMes = "JUL"
ElseIf CelMes = 8 Then
   strMes = "AGO"
ElseIf CelMes = 9 Then
   strMes = "SET"
ElseIf CelMes = 10 Then
   strMes = "OUT"
ElseIf CelMes = 11 Then
   strMes = "NOV"
ElseIf CelMes = 12 Then
   strMes = "DEZ"
End If

'Concate o Mês e Ano para obter o nome da Aba
strAno = Mid(Trim(Str(CelAno)), 3, 2)
strMesAno = strMes + "-" + strAno

'Obtêm o total de linhas
NumTotalLinhas = Range("A6").End(xlDown).Row 'CAMPO DATA

'Inicializa os contadores
linha = 1
NumDataVazia = 0
strData_anterior = " "
PrimeiraLinha = 6
ContNumeros = 0
ContAgendados = 0
ContJaAgendados = 0
ContSemDireito = 0
ContSemInteresse = 0
ContEngano = 0
ContNaoAtende = 0
ContRetornar = 0
ContCliente = 0
ContDoenca = 0

For NumRow = PrimeiraLinha To NumTotalLinhas
   
   'Lendo o campo DATA - Célula A
   Coluna = "A" + Trim(Str(NumRow))
   strData = Worksheets(strMesAno).Range(Coluna).Text
   ColunaContato = "B" + Trim(Str(NumRow))
   strValorContato = Worksheets(strMesAno).Range(ColunaContato).Text
   If ((Not IsEmpty(Trim(strData))) And (Trim(strData) <> "")) Or (Trim(strData) = "" And (Mid(Trim(strValorContato), 1, 1) = "M")) Then
     NumDataVazia = 0
     
     If Trim(strData) = "06/10/2023" Then
       NumDataVazia = 0
     End If
     
     If (((strData_anterior <> " ") And (strData <> strData_anterior)) _
         Or ((strData_anterior <> " ") And (Mid(Trim(strValorContato), 1, 1) = "M"))) Then
        array_Data(linha) = strData_anterior
        array_numeros(linha) = ContNumeros
        array_ja_agendados(linha) = ContJaAgendados
        array_agendados(linha) = ContAgendados
        array_sem_direto(linha) = ContSemDireito
        array_sem_interesse(linha) = ContSemInteresse
        array_engano(linha) = ContEngano
        array_nao_atende(linha) = ContNaoAtende
        array_retornar(linha) = ContRetornar
        array_cliente(linha) = ContCliente
        array_doenca(linha) = ContDoenca
        ContAgendados = 0
        ContJaAgendados = 0
        ContSemDireito = 0
        ContSemInteresse = 0
        ContNumeros = 0
        ContEngano = 0
        ContNaoAtende = 0
        ContRetornar = 0
        ContCliente = 0
        ContDoenca = 0
        linha = linha + 1
        NumDataVazia = 0
     End If
     
     If (Mid(Trim(strValorContato), 1, 1) <> "M") Then
       strData_anterior = strData
     End If
   End If
   
   If (Not IsEmpty(Trim(strData))) And (Trim(strData) <> "") Then
      ContNumeros = ContNumeros + 1
   End If
   
   'Lendo o campo ORIGEM DO CLIENTE - Célula C
   Coluna = "C" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 2) = "JÁ" Then
     ContJaAgendados = ContJaAgendados + 1
   End If
   
   'Lendo o campo AGENDADO - Célula G
   Coluna = "G" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Trim(strValor) = "X" Then
     ContAgendados = ContAgendados + 1
   End If
   
   'Lendo o campo SEM DIREITO - Célula I
   Coluna = "I" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 1) = "X" Then
     ContSemDireito = ContSemDireito + 1
   End If
   
   'Lendo o campo DOENÇA - Célula J
   Coluna = "J" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 1) = "X" Then
     ContDoenca = ContDoenca + 1
   End If
   
   'Lendo o campo SEM INTERESSE - Célula K
   Coluna = "K" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 1) = "X" Then
     ContSemInteresse = ContSemInteresse + 1
   End If
   
   'Lendo o campo ENGANO - Célula L
   Coluna = "L" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 1) = "X" Then
     ContEngano = ContEngano + 1
   End If
   
   'Lendo o campo NÃO ATENDE - Célula M
   Coluna = "M" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 1) = "X" Then
     ContNaoAtende = ContNaoAtende + 1
   End If
   
   'Lendo o campo CLIENTE - Célula N
   Coluna = "N" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 1) = "X" Then
     ContCliente = ContCliente + 1
   End If
   
   'Lendo o campo RETORNAR - Célula O
   Coluna = "O" + Trim(Str(NumRow))
   strValor = Worksheets(strMesAno).Range(Coluna).Text
   If Mid(Trim(strValor), 1, 1) = "X" Then
     ContRetornar = ContRetornar + 1
   End If
   
   'Testa se a célula está vazia
   If IsEmpty(Trim(strData)) Or ((Trim(strData) = "")) Then
     NumDataVazia = NumDataVazia + 1
   Else
     NumDataVazia = 0
   End If
   
   'Quando encontra 3 celulas vazias consecutivas, escreve a última linha e sai do loop
   If NumDataVazia > 3 Then
     array_Data(linha) = strData_anterior
     array_numeros(linha) = ContNumeros
     array_ja_agendados(linha) = ContJaAgendados
     array_agendados(linha) = ContAgendados
     array_sem_direto(linha) = ContSemDireito
     array_sem_interesse(linha) = ContSemInteresse
     array_engano(linha) = ContEngano
     array_nao_atende(linha) = ContNaoAtende
     array_retornar(linha) = ContRetornar
     array_cliente(linha) = ContCliente
     array_doenca(linha) = ContDoenca
     Exit For
   End If
Next
      
'Inicializa os totalizadores
Total_Numeros = 0
Total_agendados = 0
Total_JaAgendados = 0
Total_sem_direto = 0
Total_sem_interesse = 0
Total_engano = 0
Total_nao_atende = 0
Total_retornar = 0
Total_nao_atende = 0
Total_retornar = 0
Total_cliente = 0
Total_doenca = 0

PrimeiraLinha = 8
For NumRow = 1 To 31
   If array_Data(NumRow) = "" Then
     Exit For
   End If
   
   'Preenche o campo DATA
   Coluna = "B" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna).NumberFormat = "@"
   Worksheets("INDICADORES").Range(Coluna) = array_Data(NumRow)
   Worksheets("INDICADORES").Range(Coluna).NumberFormat = "DD/MM/YYYY;@"
   'Preenche o campo NUMEROS
   Coluna = "C" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_numeros(NumRow)
   'Preenche o campo AGENDADOS
   Coluna = "D" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_agendados(NumRow)
   'Preenche o campo JÁ AGENDADOS
   Coluna = "E" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_ja_agendados(NumRow)
   'Preenche o campo SEM DIREITO
   Coluna = "F" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_sem_direto(NumRow)
   'Preenche o campo SEM DIREITO
   Coluna = "G" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_sem_interesse(NumRow)
   'Preenche o campo ENGANO
   Coluna = "H" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_engano(NumRow)
   'Preenche o campo NÃO ATENDE
   Coluna = "I" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_nao_atende(NumRow)
   'Preenche o campo RETORNAR
   Coluna = "J" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_retornar(NumRow)
   'Preenche o campo CLIENTE
   Coluna = "K" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_cliente(NumRow)
   'Preenche o campo DOENÇA
   Coluna = "L" + Trim(Str(PrimeiraLinha))
   Worksheets("INDICADORES").Range(Coluna) = array_doenca(NumRow)
   
   
   PrimeiraLinha = PrimeiraLinha + 1
   Total_Numeros = Total_Numeros + array_numeros(NumRow)
   Total_agendados = Total_agendados + array_agendados(NumRow)
   Total_JaAgendados = Total_JaAgendados + array_ja_agendados(NumRow)
   Total_sem_direto = Total_sem_direto + array_sem_direto(NumRow)
   Total_sem_interesse = Total_sem_interesse + array_sem_interesse(NumRow)
   Total_engano = Total_engano + array_engano(NumRow)
   Total_nao_atende = Total_nao_atende + array_nao_atende(NumRow)
   Total_retornar = Total_retornar + array_retornar(NumRow)
   Total_cliente = Total_cliente + array_cliente(NumRow)
   Total_doenca = Total_doenca + array_doenca(NumRow)
Next

'TOTAIS
Worksheets("INDICADORES").Range("C39") = Total_Numeros
Worksheets("INDICADORES").Range("D39") = Total_agendados
Worksheets("INDICADORES").Range("E39") = Total_JaAgendados
Worksheets("INDICADORES").Range("F39") = Total_sem_direto
Worksheets("INDICADORES").Range("G39") = Total_sem_interesse
Worksheets("INDICADORES").Range("H39") = Total_engano
Worksheets("INDICADORES").Range("I39") = Total_nao_atende
Worksheets("INDICADORES").Range("J39") = Total_retornar
Worksheets("INDICADORES").Range("K39") = Total_cliente
Worksheets("INDICADORES").Range("L39") = Total_doenca
```
