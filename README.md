# Transformação de Estrutura de Dados da Balança Comercial

Este repositório contém um script VBA desenvolvido para transformar a estrutura de dados da balança comercial, facilitando a utilização de tabelas dinâmicas e análise de dados cross section agrupados por anos de 1997 até 2023.

## Funcionalidades

O script realiza as seguintes operações:

1. Preenche uma coluna com os anos de 1997 a 2023.
2. Copia os dados da planilha de origem ("Resultado") para a planilha de destino ("Planilha1") de forma que cada ano tenha um bloco correspondente de dados.
3. Organiza os dados de modo a permitir análises mais eficientes com tabelas dinâmicas.

## Estrutura do Script

### Variáveis

- `wsResultado`: Planilha de origem dos dados.
- `wsPlanilha1`: Planilha de destino dos dados transformados.
- `startRow`: Linha inicial para começar a preencher os anos.
- `startYear`: Ano inicial (1997).
- `endYear`: Ano final (2023).
- `rowCounter`: Contador de linhas para preenchimento dos anos.
- `yearCounter`: Contador de anos no loop.
- `sourceRange`: Intervalo de origem dos dados a serem copiados.
- `destRange`: Intervalo de destino onde os dados serão colados.

### Passo a Passo

1. **Definição das Planilhas**:
    ```vba
    Set wsResultado = ThisWorkbook.Sheets("Resultado")
    Set wsPlanilha1 = ThisWorkbook.Sheets("Planilha1")
    ```

2. **Configuração dos Parâmetros Iniciais**:
    ```vba
    startRow = 2
    startYear = 1997
    endYear = 2023
    wsPlanilha1.Cells(1, 1).Value = "Ano"
    rowCounter = startRow
    ```

3. **Preenchimento dos Anos**:
    ```vba
    For yearCounter = startYear To endYear
        For i = 2 To 10139
            wsPlanilha1.Cells(rowCounter, 1).Value = yearCounter
            rowCounter = rowCounter + 1
        Next i
    Next yearCounter
    ```

4. **Cópia de Dados**:
    - **Cópia Inicial**:
        ```vba
        Set sourceRange = wsResultado.Range(wsResultado.Cells(2, 1), wsResultado.Cells(10139, 6))
        
        For i = 0 To totalYears - 1
            Set destRange = wsPlanilha1.Range(wsPlanilha1.Cells(startRow + (i * 10138), 2), wsPlanilha1.Cells(startRow + (i * 10138) + 10138, 7))
            sourceRange.Copy destRange
        Next i
        ```

    - **Cópia Adicional**:
        ```vba
        endRow = 10139
        destRow = 2
        
        For colCounter = 59 To 7 Step -2
            Set sourceRange = wsResultado.Range(wsResultado.Cells(startRow, colCounter), wsResultado.Cells(endRow, colCounter + 1))
            Set destRange = wsPlanilha1.Cells(destRow, 8) ' Colar sempre nas colunas H e I (8 e 9)
            sourceRange.Copy destRange
            destRow = destRow + 10138 ' Incrementa a linha de destino para a próxima iteração
        Next colCounter
        ```

## Como Usar

1. Abra o Excel e carregue o arquivo que contém os dados da balança comercial.
2. Pressione `ALT + F11` para abrir o Editor do VBA.
3. Insira um novo módulo e cole o script VBA fornecido.
4. Execute o script (`F5`) para transformar os dados conforme descrito.

## Contribuições

Sinta-se à vontade para fazer um fork deste repositório, criar issues ou enviar pull requests para melhorias no script.

---

Se tiver alguma dúvida ou sugestão, por favor entre em contato. Agradecemos seu interesse e colaboração!
