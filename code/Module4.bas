Attribute VB_Name = "Module4"
'Dim obj As New Class1
'
''Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
''ByVal lpClassName As String, _
''ByVal lpWindowName As String) As Long
''Declare Function SetForegroundWindow Lib "user32" ( _
''ByVal hwnd As Long) As Long
'
'
'
'Sub uploadCapture()
'
'
'hwnd = 0
'hwnd = FindWindow("WindowsForms10.Window.8.app.0.33c0d9d", "SAFE-DOC Capture")
'nrProcSafe = hwnd
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'Application.ScreenUpdating = False
'
'If ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 11).Value = "Upload Safe-DOC" Then
'
'hwnd = 0
'hwnd = FindWindow("WindowsForms10.Window.8.app.0.33c0d9d", "SAFE-DOC Capture")
'nrProcSafe_ = hwnd
'
'If nrProcSafe_ <> nrProcSafe Then
'
'MsgBox "SAFE-DOC Catpture deslogado!"
'Exit Sub
'
'End If
'
'
'obj.Wait 500
'Call Shell("C:\Program Files (x86)\Acesso Digital\SAFE-DOC Capture\SafeDOC.Capture.exe")
'DoEvents
'obj.Wait 500
'hwnd = FindWindow(vbNullString, "SAFE-DOC Capture")
'SetForegroundWindow (hwnd)
'
'
'obj.Wait 500
'SendKeys "%L"
'DoEvents
'
''chama a função de captura e posiciona o cursor TAB no primeiro campo.
'obj.Wait 500
'SendKeys "%C"
'DoEvents
'
''Ativa campo origem
'obj.Wait 500
'SendKeys "{TAB}"
'DoEvents
'obj.Wait 500
''Seleciona Vivo
'SendKeys "VIVO"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Preenche campo Classe
'obj.Wait 500
'SendKeys "Processamento Janela Única - TBRA"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'obj.Wait 500
'SendKeys "{TAB}"
'DoEvents
'
''Campo Status de entrada
'obj.Wait 500
'segOPCP = ThisWorkbook.Sheets("PROC_CODE").Range("D7").Value
'SendKeys (segOPCP)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Esta inserindo estado do destinatário, baseado nas informações do XML
'obj.Wait 500
'SendKeys "VIVO_SP"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Priorização
'obj.Wait 500
'priorizacao = ThisWorkbook.Sheets("PROC_CODE").Range("D10").Value
'SendKeys (priorizacao)
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Trata e inseri data de protocolo
'obj.Wait 500
'data_prot = Format(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 10).Value, "DD/MM/YYYY")
'DoEvents
'SendKeys (data_prot)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''CNPJ
'obj.Wait 500
'SendKeys "^a"
'obj.Wait 500
'SendKeys "{BACKSPACE}"
'obj.Wait 500
'cnpjFornec = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 25, False)
'SendKeys (cnpjFornec)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Nome do fornecedor
'obj.Wait 500
'SendKeys "^a"
'obj.Wait 500
'SendKeys "{BACKSPACE}"
'obj.Wait 500
'nomeFornec = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 26, False)
'SendKeys (nomeFornec)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Filial da emissão
'obj.Wait 500
'filEmi = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 33, False)
'SendKeys (filEmi)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Linha de Negócio
'obj.Wait 500
'SendKeys "TBRA -  BRASIL LTDA"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Tipo do Documento
'obj.Wait 500
'SendKeys "ME - MERCADORIA"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Sub Tipo de Documento
'obj.Wait 500
'SendKeys "MT - MATERIAL"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Segmento
'obj.Wait 500
'
'If Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "57" Then
'SendKeys "CP - Capex"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
'ElseIf Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "65" Then
'SendKeys "OP - Opex"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
'ElseIf Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "54" Or Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "51" Then
'SendKeys "OP - Opex"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
'Else
'SendKeys "CP - Capex"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'End If
'
'
''Numero do documento (Numero nota) e Série
'obj.Wait 500
'SendKeys "^a"
'obj.Wait 500
'SendKeys "{BACKSPACE}"
'obj.Wait 500
'nrNota = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 8, False)
'SendKeys (nrNota)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Data de emissão
'obj.Wait 500
'dtEmi = Format(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 9, False), "DD/MM/YYYY")
'SendKeys (dtEmi)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Data vencimento
'obj.Wait 500
'dtVenc = Format(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 21, False), "DD/MM/YYYY")
'SendKeys (dtVenc)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Valor do Documento
'obj.Wait 500
'vlDoc = Replace(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 20, False), ".", ",")
'SendKeys (vlDoc)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
'
''Número do pedido
'obj.Wait 500
'nrPed = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False)
'SendKeys (nrPed)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Valor ICMS
'obj.Wait 500
'vlICMS = Replace(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 55, False), ".", ",")
'SendKeys (vlICMS)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Valor da Substituição tributária 64
'obj.Wait 500
'vlSICMS = Replace(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 56, False), ".", ",")
'SendKeys (vlSICMS)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Chave de Acesso para Danfes
'obj.Wait 500
'chaveDoc = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2)
'SendKeys (chaveDoc)
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
'
'obj.Wait 500
'SendKeys "{TAB}"
'DoEvents
'
'
''Etiqueta
'obj.Wait 500
'SendKeys "^a"
'obj.Wait 500
'SendKeys "{DELETE}"
'etiqueta = ThisWorkbook.Sheets("PROC_CODE").Range("H7").Value
'obj.Wait 500
'SendKeys (etiqueta)
'DoEvents
'
'
'obj.Wait 500
'SendKeys "%A"
'DoEvents
'
'
'obj.Wait 500
'SendKeys "{TAB}"
'DoEvents
'obj.Wait 500
'SendKeys "{TAB}"
'DoEvents
'obj.Wait 500
'SendKeys "{TAB}"
'DoEvents
'obj.Wait 500
'SendKeys "{TAB}"
'DoEvents
'obj.Wait 500
'SendKeys "~"
'DoEvents
'obj.Wait 500
'SendKeys "{DELETE}"
'DoEvents
'
'obj.Wait 500
'diretorio_extracao = ThisWorkbook.Path & "\"
'SendKeys (diretorio_extracao)
'obj.Wait 500
'SendKeys "~"
'DoEvents
'
''Condiciona se sera feito upload em tif ou pdf
'If ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Converter para TIF" Then
'
'obj.Wait 500
'SendKeys "%N"
'DoEvents
'obj.Wait 500
'nmArquiv = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".tif"
'SendKeys (nmArquiv)
'DoEvents
'
'ElseIf ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Upload em PDF" Then
'
'obj.Wait 500
'SendKeys "%N"
'DoEvents
'obj.Wait 500
'nmArquiv = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".pdf"
'SendKeys (nmArquiv)
'DoEvents
'
'End If
'
'
'obj.Wait 500
'SendKeys "%o" '--Abrir arquivo
'DoEvents
'
'obj.Wait 4500
''SendKeys "%E"
''DoEvents
'
'obj.Wait 1500
''SendKeys "%{F4}"
''DoEvents
''
'obj.Wait 1000
'SendKeys "%L"
'DoEvents
'
'obj.Wait 200
'SendKeys "%{F4}"
'DoEvents
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 11).Value = "Inserido SAFE-DOC"
'
'End If
'
'Application.ScreenUpdating = True
'
'cont_linha = cont_linha + 1
'Loop
'
'End Sub
'
''Add Open Twebst Type Library in Tools/References menu of the VBA editor.
'Sub uploadWEB()
'cont_linha = 15
'
'Dim nmArquiv As String
'Dim priorizacao As String
'Dim filEmi As String
'Dim varSegmento As String
'
'Dim core
'Dim browser
'Set core = CreateObject("OpenTwebst.Core")
'Set browser = core.StartBrowser("https://www.acessodigital.com.br/grupovivo/LoginNFD.aspx?ReturnUrl=/grupovivo/NewObject.aspx")
'
'    Application.Wait (Now + TimeValue("0:00:04"))
'
'    Call browser.FindElement("input text", "id=SignInNFD1_txtUsername").InputText("gestor")
'    Call browser.FindElement("input password", "id=SignInNFD1_txtPassword").InputText("gestor")
'
'
'    Call browser.FindElement("input submit", "id=SignInNFD1_btnLogin").Click
'
'
'    Call browser.FindElement("a", "uiname=Novo").Click
'
'    Call browser.FindElement("select", "id=drpSources").Select("VIVO")
'
'    Call browser.FindElement("img", "src=ig_treeOplus.gif").Click
'    Call browser.FindElement("img", "src=ig_treeLplus.gif").Click
'
'
'    Call browser.FindElement("span", "uiname=Processamento Janela Única - TBRA").Click
'
'    'Condiciona se sera feito upload em tif ou pdf
'If ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Converter para TIF" Then
'
'DoEvents
'nmArquiv = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".tif"
'DoEvents
'
'ElseIf ThisWorkbook.Sheets("PROC_CODE").Range("F10").Value = "Upload em PDF" Then
'
'DoEvents
'nmArquiv = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".pdf"
'
'DoEvents
'
'End If
'
'obj.Wait 500
'    Call browser.FindElement("input file", "id=IDX_FILE").InputText(nmArquiv)
'
'    Call browser.FindElement("select", "id=IDX_465").Select("VIVO_SP")
'
''Priorização
'priorizacao = ThisWorkbook.Sheets("PROC_CODE").Range("D10").Value
'DoEvents
'
'    Call browser.FindElement("td", "uiname=Selecione... Cobiling Janela Expressa (demais) Janela Expressa Operação Jurídico Marketing Padrão RH Tributos Sensível*").Click
'    Call browser.FindElement("select", "id=IDX_471").Select("Janela Expressa (demais)")
'
'
'    Call browser.FindElement("select", "id=IDX_486").Select("NÃO")
'
'
'DoEvents
'data_prot = Format(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 10).Value, "DD/MM/YYYY")
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_292").InputText(data_prot)
'
'
'DoEvents
'cnpjFornec = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 25, False)
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_59").InputText(cnpjFornec)
'
'    Call browser.FindElement("select", "id=IDX_470").Select("Selecione...")
'
'DoEvents
'filEmi = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 33, False)
'DoEvents
'
'    Call browser.FindElement("div", "id=pnlShowForm").Click
'    Call browser.FindElement("select", "id=IDX_416").Select("SP")
'
'
'
'    Call browser.FindElement("select", "id=IDX_223").Select("TBRA -  BRASIL LTDA")
'    Call browser.FindElement("select", "id=IDX_466").Select("ME - MERCADORIA")
'    Call browser.FindElement("select", "id=IDX_467").Select("MT - MATERIAL")
'
'If Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "57" Then
'varSegmento = "CP - CAPEX"
'DoEvents
'
'ElseIf Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "65" Then
'varSegmento = "OP - OPEX"
'DoEvents
'
'ElseIf Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "54" Or Left(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False), 2) = "51" Then
'varSegmento = "OP - OPEX"
'DoEvents
'
'Else
'varSegmento = "CP - CAPEX"
'DoEvents
'
'End If
'
'    Call browser.FindElement("select", "id=IDX_468").Select(varSegmento)
'
'DoEvents
'nrNota = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 8, False)
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_134").InputText(nrNota)
'
'DoEvents
'dtEmi = Format(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 9, False), "DD/MM/YYYY")
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_139").InputText(dtEmi)
'
'DoEvents
'nrPed = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 58, False)
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_227").InputText(nrPed)
'
'DoEvents
'vlICMS = Replace(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 55, False), ".", ",")
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_286").InputText(vlICMS)
'
'DoEvents
'vlSICMS = Replace(Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), 56, False), ".", ",")
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_484").InputText(vlSICMS)
'
'DoEvents
'chaveDoc = ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2)
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_472").InputText(chaveDoc)
'
'
'    Call browser.FindElement("input text", "id=IDX_409").InputText("0")
'    Call browser.FindElement("input text", "id=IDX_414").InputText("0")
'    Call browser.FindElement("input text", "id=IDX_413").InputText("0")
'    Call browser.FindElement("input text", "id=IDX_289").InputText("0")
'    Call browser.FindElement("input text", "id=IDX_284").InputText("0")
'    Call browser.FindElement("input text", "id=IDX_234").InputText("0")
'    Call browser.FindElement("input text", "id=IDX_239").InputText("0")
'    Call browser.FindElement("select", "id=IDX_487").Select("Nota fiscal eletrônica de Material = DANFE")
'    Call browser.FindElement("input text", "id=IDX_236").InputText("")
'
'DoEvents
'etiqueta = ThisWorkbook.Sheets("PROC_CODE").Range("H7").Value
'DoEvents
'
'    Call browser.FindElement("input text", "id=IDX_429").InputText(etiqueta)
'
''    Call browser.FindElement("input submit", "id=btnNew, index=1").Click
'
'End Sub
