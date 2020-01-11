Attribute VB_Name = "Module1"
''Variavel de ok para as datas de protocolo
' Public continuamacro As String
' Public quantidadeChaves As Integer
'
''Verifica a internet
''Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" _
''                                                     (ByRef lpdwFlags As Long, _
''                                                      ByVal lpszConnectionName As String, _
''                                                      ByVal dwNameLen As Integer, _
''                                                      ByVal dwReserved As Long) _
''                                                      As Long
'
''Função para download de URL, com condicinamento para versões de Windws 32bits e 64bits
'#If Win64 Then
'
'    Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias _
'        "URLDownloadToFileA" (ByVal pCaller As Long, _
'        ByVal szURL As String, _
'        ByVal szFileName As String, _
'        ByVal dwReserved As Long, _
'        ByVal lpfnCB As Long) As Long
'
'#Else
'
'    Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias _
'        "URLDownloadToFileA" (ByVal pCaller As Long, _
'        ByVal szURL As String, _
'        ByVal szFileName As String, _
'        ByVal dwReserved As Long, _
'        ByVal lpfnCB As Long) As Long
'#End If
'
''Função para donwload de URLS
'Function DownloadFile(Url As String, LocalFilename As String) As Boolean
'    Dim lngRetVal      As Long
'    lngRetVal = URLDownloadToFile(0, Url, LocalFilename, 0, 0)
'    If lngRetVal = 0 Then DownloadFile = True
'End Function
'
'Public Function Get32BitProgramFilesPath() As String
'    If Environ("ProgramW6432") = "" Then
'       '32 bit Windows
'       Get32BitProgramFilesPath = Environ("ProgramFiles")
'    Else
'       '64 bit Windows
'       Get32BitProgramFilesPath = Environ("ProgramFiles(x86)")
'    End If
'End Function
'
''Funçao para verificar se arquivo existe
'Function CaminhoExiste(sCaminho As String) As Boolean
'
'    If Dir(sCaminho) = vbNullString Then
'        CaminhoExiste = False
'    Else
'        CaminhoExiste = True
'    End If
'
'    'A forma abreviada da função pode ser escrita como:
'    'CaminhoExiste = Dir(sCaminho) <> vbNullString
'End Function
'
'Sub geral()
'
''zera variavel
'continuamacro = ""
'quantidadeChaves = 0
'
'
'
'Call contaChaves
'ThisWorkbook.Sheets("PROC_CODE").Range("A13").Value = quantidadeChaves
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Testtando conectividade..."
'Application.ScreenUpdating = False
'
''Testa internet
'Call testeInternet
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Verificando data de protocolo..."
'Application.ScreenUpdating = False
'
'Call integridadeDataProtocolo
'
''Verifica retorno de datas preenchidas
'    If continuamacro = 0 Then
'
'    MsgBox "Notas sem data de protocolo preenchida!"
'
'    Exit Sub
'
'    End If
'
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Buscando número de protocolo e nota... "
'
'Application.ScreenUpdating = False
'
''Obtem informações do número de protocolo e numero da nota
''Informações necessária para downloads de XML e PDF
'Call downloadArquivExtracao2
'
''Call integridadeInfoNota
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Realizando download de XML..."
'Application.ScreenUpdating = False
'
''Realiza download dos XMLs
'Call baixaXML
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Carregando informações do XML..."
'Application.ScreenUpdating = False
'
''Carrega informações do XML
'Call importaXML1
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Calculando quantidade de itens de cada nota..."
'Application.ScreenUpdating = False
'
''Calcula a quantidade de itens da nota
'Call qtdItensXML
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Verificando campos essênciais do XML..."
'Application.ScreenUpdating = False
'
''Verifica integridade do XML
'Call integridadeXML
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Realizando download dos PDFs..."
'Application.ScreenUpdating = False
'
''Baixa o PDFs
'Call baixaPDF
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Verificando integridade dos PDFs..."
'Application.ScreenUpdating = False
'
''Verifica integridade dos PDFs
'Call integridadePDF
'
'If ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Converter para TIF" Then
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Convertendo para TIF..."
'Application.ScreenUpdating = False
'
''Faz conversão para TIF (se necessário)
'Call conversorTIF
'
'End If
'
''Verifica integridade do TIF ou se não é necessario baixa-lo
'Call integridadeTIF
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Verificando número de pedido..."
'Application.ScreenUpdating = False
'
''Busca numero do pedido quando estiver presente no campo xped do XML
'Call numeroPedido
'
''Verifica os numeros de pedidos encontrados, e da ok para macro continar
'Call integridadeNroPedido
'
''Preenche numero de lote
'Call preencheLote
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Verificando notas a serem processadas..."
'Application.ScreenUpdating = False
'
''Verifica se a nota deve ir ou não para o SAFE-DOC
'Call statusNota
'
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Pronto para iniciar upload no SAFE-DOC"
'Application.ScreenUpdating = False
'
'
'Call uploadCapture
'
''Atualiza tela
'Application.ScreenUpdating = True
'ThisWorkbook.Sheets("PROC_CODE").Range("B6").Value = "Upload realizado no SAFE-DOC"
'Application.ScreenUpdating = False
'
'End Sub
'
'Sub contaChaves()
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'quantidadeChaves = quantidadeChaves + 1
'
'cont_linha = cont_linha + 1
'Loop
'
'End Sub
'
'
''Verifica a conexão com a internet
'Sub testeInternet()
'
'Dim sConnType As String * 255
'
'    Dim Ret As Long
'    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
'    If Ret <> 1 Then
'        MsgBox "Verifique a conexão com a internet"
'        Exit Sub
'
'    End If
'
'End Sub
'
'Sub preencheLote()
'
''Preenche numero de lote
'ThisWorkbook.Sheets("PROC_CODE").Range("H7").Value = Format(Now(), "YYMMDDHH") & Mid(Format(Now(), "hh:mm"), 4, 1)
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 9).Value = ThisWorkbook.Sheets("PROC_CODE").Range("H7").Value
'
'
'cont_linha = cont_linha + 1
'Loop
'
'
'End Sub
'
'
'Sub conversorTIF()
'Dim s As String
'Dim irfanview_path As String
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'If ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 6).Value = "OK" And ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Converter para TIF" Then
'
'    'Define caminho do irfanview
'    irfanview_path = Get32BitProgramFilesPath() & "\IrfanView\i_view32.exe"
'
'    'Define o arquivo origem
'    source_pdf_path = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2) & ".pdf"
'
'    'Define arquivo de destino
'    destiny_tif_path = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2) & ".tif"
'
'    'Invoca irfanview e paramentros
'    Call Shell(irfanview_path & " " & source_pdf_path & " /convert=" & destiny_tif_path & "/filepattern=$N /effect=(1, 1, 0) /effect=(6,3,0) /sharpen=33 /tifc=1 /dpi=(300,300)", vbHide)
'    DoEvents
'
'End If
'
'cont_linha = cont_linha + 1
'
'Loop
'
''finaliza irfanview
'Call Shell("C:\Program Files (x86)\IrfanView\i_view32.exe /killmesoftly")
'Call Shell("taskkill /im cmd.exe /f")
'Call Shell("taskkill /im i_view32.exe /f")
'
'End Sub
'
'
'Sub downloadArquivExtracao()
'Dim IE As Object
'
''Parametros
'arquivo_extracao = "arquivo_extracao.xls"
'
'
''Mata processos do IE
'Shell ("taskkill /IM iexplore.exe /f")
'Application.Wait (Now + TimeValue("0:00:01"))
'
''Desfaz objeto
'Set IE = Nothing
'
'    ' Create InternetExplorer Object
'    Set IE = CreateObject("InternetExplorer.Application")
'    IE.Top = 0
'    IE.Left = 0
'    IE.Width = 800
'    IE.Height = 600
'    IE.AddressBar = 0
'    IE.StatusBar = 0
'    IE.Toolbar = 0
'    IE.Visible = False 'We will see the window navigation
'
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
''Abre o portal de fornecedore e aguarda o carregamento da página e carrega por data de inicio e fim
'
'    ' Send the form data To URL As POST binary request
'    IE.navigate "http://fornecedores.portalfiscaldigital.com.br/pfd/excel/excel_proc.php?chave=" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value
'
'    'Desativa warnings
'    Application.DisplayAlerts = False
'
'    'Wait for full load page
'    Do
'    DoEvents
'    Loop Until IE.ReadyState = 4
'
'
'Set allhyperlinks = IE.document.getElementsByTagName("a")
'
'Dim linksarray(5) As String
'
'a = 1
'For Each hyper_link In allhyperlinks
'
'linksarray(a) = (hyper_link)
'a = a + a
'
'Next
'
''Faz o download do Excel
'    Dim sURL    As String
'    Dim LocalFilename   As String
'    Dim FileName As String
'
'    Debug.Assert DownloadFile(linksarray(2), ThisWorkbook.Path & "\" & "arquivo_extracao.xls")
'
''Desliga alerta e atulizaçoes de tela
'Application.DisplayAlerts = False
'Application.ErrorCheckingOptions.NumberAsText = False
'Application.ScreenUpdating = False
'
''Abre excel com dados da chave
'Workbooks.Open ThisWorkbook.Path & "\" & (arquivo_extracao)
'
'
''Ordena por data de A-Z e por chave de A-Z
'    Workbooks("arquivo_extracao.xls").Sheets("RELATORIO").Range("A2").Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Range(Selection, Selection.End(xlDown)).Select
'    ActiveWorkbook.Worksheets("RELATORIO").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("RELATORIO").Sort.SortFields.Add Key:=Range( _
'        "E2:E100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'        xlSortTextAsNumbers
'    ActiveWorkbook.Worksheets("RELATORIO").Sort.SortFields.Add Key:=Range( _
'        "A2:A100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'        xlSortTextAsNumbers
'    ActiveWorkbook.Worksheets("RELATORIO").Sort.SortFields.Add Key:=Range( _
'        "F2:F100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'        xlSortTextAsNumbers
'    With ActiveWorkbook.Worksheets("RELATORIO").Sort
'        .SetRange Range("A2:G100")
'        .Header = xlGuess
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
''Remove as duplicadas, deixando apenas as ultimas notas inseridas
'    Range("A2").Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Range(Selection, Selection.End(xlDown)).Select
'    ActiveSheet.Range("$A$2:$G$2863").RemoveDuplicates Columns:=Array(2, 3, 4, 6, 7), Header:=xlNo
'
''Copia o conteudo para ser processado
'chave_orig = Workbooks("arquivo_extracao.xls").Sheets("RELATORIO").Range("F2").Value
'chave_dest = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value
'
''Verifica se o excel é mesmo da chave de acesso
'If chave_orig = chave_dest And Workbooks("arquivo_extracao.xls").Sheets("RELATORIO").Range("D2") <> "" Then
'
''Copia numero da nota
'Workbooks("arquivo_extracao.xls").Sheets("RELATORIO").Activate
'ActiveSheet.Range("D2").Select
'Selection.Copy
'ThisWorkbook.Sheets("PROC_CODE").Activate
'Cells(cont_linha, 3).Select
'ActiveSheet.Paste
'
''Copia numero de protocolo
'Workbooks("arquivo_extracao.xls").Sheets("RELATORIO").Activate 'ativa  a planilha desejada
'ActiveSheet.Range("A2").Select
'Selection.Copy
'ThisWorkbook.Sheets("PROC_CODE").Activate
'Cells(cont_linha, 4).Select
'ActiveSheet.Paste
'
'Else
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 3).Value = "NOK"
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 4).Value = "NOK"
'
'
'End If
'
''Fecha o aqruivo com dados da chave
'Workbooks("arquivo_extracao.xls").Close SaveChanges:=False
'
''Apaga arquivo com dados da nota
'Kill ThisWorkbook.Path & "\" & (arquivo_extracao)
'
'cont_linha = cont_linha + 1
'ThisWorkbook.Save
'
'Loop
'
'
''Mata processos do IE
'Shell ("taskkill /IM iexplore.exe /f")
'Application.Wait (Now + TimeValue("0:00:01"))
'
''Desfaz objeto
'Set IE = Nothing
'
'End Sub
'
'Sub downloadArquivExtracao2()
'Dim IE As Object
'
''Parametros
'arquivo_extracao = "arquivo_extracao.xls"
'
'
''Mata processos do IE
'Shell ("taskkill /IM iexplore.exe /f")
'DoEvents
'Application.Wait (Now + TimeValue("0:00:01"))
'
''Desfaz objeto
'Set IE = Nothing
'DoEvents
'Application.Wait (Now + TimeValue("0:00:01"))
'
'    ' Create InternetExplorer Object
'    Set IE = CreateObject("InternetExplorer.Application")
'    IE.Top = 0
'    IE.Left = 0
'    IE.Width = 800
'    IE.Height = 600
'    IE.AddressBar = 0
'    IE.StatusBar = 0
'    IE.Toolbar = 0
'    IE.Visible = False 'We will see the window navigation
'
'    ' Send the form data To URL As POST binary request
'    IE.navigate "http://fornecedores.portalfiscaldigital.com.br/pfd/processadas.php?UserID=&Page=1&FirstTime=True"
'
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'Application.ScreenUpdating = False
'
'
''Abre o portal de fornecedore e aguarda o carregamento da página e carrega por data de inicio e fim
'
'
'    'Desativa warnings
'    Application.DisplayAlerts = False
'
'    'Wait for full load page
'    Do
'    DoEvents
'    Loop Until IE.ReadyState = 4
'
''Preenche campo chave
'Set chave = IE.document.all("chave")
'chave.Value = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value
'
'    'Wait for full load page
'    Do
'    DoEvents
'    Loop Until IE.ReadyState = 4
'
''Clica no botão pesquisar
'IE.document.all("filtro").Click
'DoEvents
'
''Wait for full load page
'Application.Wait (Now + TimeValue("0:00:05"))
'Do
'DoEvents
'Loop Until IE.ReadyState = 4
'
''Obtem os dados da nota
'        Dim ws As Worksheet
'        Dim rng As Range
'        Dim tbl As Object
'        Dim rw As Object
'        Dim cl As Object
'        Dim tabno As Long
'        Dim nextrow As Long
'        Dim I As Long
'        Dim tabelamtx(20, 9) As Single
'
'Application.Wait (Now + TimeValue("0:00:01"))
'
''limpa vetor
'ThisWorkbook.Sheets("AreaColagem").Range("A1:K1000").Clear
'ThisWorkbook.Sheets("AreaColagem").Range("A1:K1000").NumberFormat = "@"
'
'
'        cont_tab = 1
'        nextrow = 1
'        For Each tbl In IE.document.getElementsByTagName("TABLE")
'            If cont_tab = 4 Then
'            nextrow = nextrow + 1
'            End If
'            Set rng = ThisWorkbook.Sheets("AreaColagem").Range("A" & nextrow)
'            For Each rw In tbl.Rows
'                For Each cl In rw.Cells
'                If cont_tab = 4 Then
'                    rng.Value = cl.outerText
'                    Set rng = rng.Offset(, 1)
'                    I = I + 1
'                End If
'                Next cl
'                If cont_tab = 4 Then
'                nextrow = nextrow + 1
'                End If
'                Set rng = rng.Offset(1, -I)
'                I = 0
'            Next rw
'         cont_tab = cont_tab + 1
'         Next tbl
'
'Application.Wait (Now + TimeValue("0:00:01"))
'
'If ThisWorkbook.Sheets("AreaColagem").Range("F3").Value = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value Then
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 3).Value = ThisWorkbook.Sheets("AreaColagem").Range("D3").Value
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 4).Value = ThisWorkbook.Sheets("AreaColagem").Range("A3").Value
'
'ElseIf ThisWorkbook.Sheets("AreaColagem").Range("D3").Value = "" Or ThisWorkbook.Sheets("AreaColagem").Range("A3").Value = "" Then
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 3).Value = "Não encontrado"
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 4).Value = "Não encontrado"
'
'Else
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 3).Value = "Não encontrado"
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 4).Value = "Não encontrado"
'
'End If
'
'Application.ScreenUpdating = True
'
'cont_linha = cont_linha + 1
' Loop
'
''limpa vetor
'ThisWorkbook.Sheets("AreaColagem").Range("A1:K1000").Clear
'ThisWorkbook.Sheets("AreaColagem").Range("A1:K1000").NumberFormat = "@"
'
'End Sub
'
'Sub baixaPDF()
'Dim sURL As String
'Dim LocalFilename As String
'
''Desliga alerta e atulizaçoes de tela
'Application.DisplayAlerts = False
'Application.ErrorCheckingOptions.NumberAsText = False
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'If ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 5).Value = "OK" Then
'
''Baixa o PDF
'sURL = "http://fornecedores.portalfiscaldigital.com.br/pfd/danfe/gera_danfe.php?num=" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 4).Value
'arquivo_extracao = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".pdf"
'
'    'LocalFilename = UNC & filename
'    diretorio_extracao = ThisWorkbook.Path & "\"
'    LocalFilename = (diretorio_extracao) & (arquivo_extracao)
'
'    On Error Resume Next
'    Debug.Assert DownloadFile(sURL, LocalFilename)
'
'End If
'
'cont_linha = cont_linha + 1
'
'Loop
'
'End Sub
'
'Sub baixaXML()
'Dim sURL As String
'Dim LocalFilename As String
'
''Desliga alerta e atulizaçoes de tela
'Application.DisplayAlerts = False
'Application.ErrorCheckingOptions.NumberAsText = False
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
''Baixa o PDF
'sURL = "http://fornecedores.portalfiscaldigital.com.br/pfd/danfe/recupera_xml.php?arquivo=" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 4).Value & ".xml"
'arquivo_extracao = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".xml"
'
'    'LocalFilename = UNC & filename
'    diretorio_extracao = ThisWorkbook.Path & "\"
'    LocalFilename = (diretorio_extracao) & (arquivo_extracao)
'
'    Debug.Assert DownloadFile(sURL, LocalFilename)
'
'cont_linha = cont_linha + 1
'
'Loop
'
'End Sub
'
'Sub integridadePDF()
'Dim s As String
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
's = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".pdf"
'
'On Error Resume Next
'If CaminhoExiste(s) = True And FileLen(s) > 500 Then
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 6).Value = "OK"
'
'Else
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 6).Value = "NOK"
'
'End If
'
''Error handle ( para arquivos enexistentes)
'    If Err.Number = "53" Then
'    ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 6).Value = "NOK"
'    Err.Number = 0
'    End If
'
'cont_linha = cont_linha + 1
'Loop
'
'
'End Sub
'
'Sub integridadeDataProtocolo()
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'If ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> "" And ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 10).Value = "" Then
'
'continuamacro = 0
'
'Exit Sub
'
'Else
'
'continuamacro = 1
'
'End If
'
'cont_linha = cont_linha + 1
'Loop
'
'End Sub
'
'
'
'
'Sub integridadeTIF()
'Dim s As String
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
's = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".tif"
'
'On Error Resume Next
'If CaminhoExiste(s) = True And FileLen(s) > 500 Then
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 7).Value = "OK"
'
'ElseIf ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Upload em PDF" Then
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 7).Value = "N/A"
'
'Else
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 7).Value = "NOK"
'
'End If
'
''Error handle ( para arquivos enexistentes)
'    If Err.Number = "53" And ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Converter para TIF" Then
'    ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 7).Value = "NOK"
'    Err.Number = 0
'
'    ElseIf Err.Number = "53" And ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value <> "Converter para TIF" Then
'    ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 7).Value = "N/A"
'    Err.Number = 0
'
'    End If
'
'
'cont_linha = cont_linha + 1
'Loop
'
'
'End Sub
'
'Sub numeroPedido()
'
''Verifica número do pedido
'
'cont_linha = 2
'Do While ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 1).Value <> ""
'
'If ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value = "" Or _
'Application.IsNumber(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value) = False And _
'(Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) <> "570" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) <> "540" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) <> "510" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) <> "800" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) <> "650") _
'Then
'
'fim_string = 10
'
'For ini_string = 1 To 5000
'
'    If IsNumeric(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)) = True And Left(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string), 3) = "510" Then
'    ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)
'    Exit For
'
'    ElseIf IsNumeric(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)) = True And Left(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string), 3) = "540" Then
'    ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)
'    Exit For
'
'    ElseIf IsNumeric(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)) = True And Left(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string), 3) = "570" Then
'    ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)
'    Exit For
'
'    ElseIf IsNumeric(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)) = True And Left(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string), 3) = "800" Then
'    ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)
'    Exit For
'
'    ElseIf IsNumeric(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)) = True And Left(Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string), 3) = "650" Then
'    ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = Mid(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 52).Value, ini_string, fim_string)
'    Exit For
'
'    Else
'
'    ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = "NOK"
'
'    End If
'
'Next ini_string
'
'ElseIf (Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) = "570" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) = "540" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) = "510" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) = "800" Or _
'Left(ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value, 3) = "650") And _
'ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value <> "" Then
'
'
'ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 54).Value
'
'Else
'
'ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 58).Value = "NOK"
'
'End If
'
'cont_linha = cont_linha + 1
'Loop
'
'End Sub
'
'
'Sub integridadeNroPedido()
'Dim pedido As String
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'
'pedido = Application.IfError(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), "")
'
'If pedido <> "" And pedido <> "NOK" Then
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 8).Value = "OK"
'
'Else
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 8).Value = "NOK"
'
'End If
'
'cont_linha = cont_linha + 1
'Loop
'
'
'End Sub
'
'Sub dataProtocolo()
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'If ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> "" And ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 10).Value = "" Then
'
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 10).Value = ThisWorkbook.Sheets("PROC_CODE").Range("F7").Value
'
'End If
'
'
'cont_linha = cont_linha + 1
'Loop
'
'End Sub
'
'
'Sub statusNota()
'
''Defini status da nota para proseeguir o upload no safedoc
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'If ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 5).Value = "NOK" Or _
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 6).Value = "NOK" Or _
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 7).Value = "NOK" Or _
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 8).Value = "NOK" Then _
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 11).Value = "Não processar"
'
'Else
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 11).Value = "Upload Safe-DOC"
'
'End If
'
'cont_linha = cont_linha + 1
'Loop
'
'End Sub
