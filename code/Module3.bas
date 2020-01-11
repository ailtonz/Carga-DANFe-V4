Attribute VB_Name = "Module3"
'Sub importaXML1()
'    Dim xmlDOM As DOMDocument    ' Variável de manipulação de documento
'    Dim objNodes As IXMLDOMNodeList    ' Lista de nós
'    Dim s As String
'
'
''Array com campos
'    Dim camposXML(1 To 56) As String
'
'    camposXML(1) = "chNFe"
'    camposXML(2) = "cUF"
'    camposXML(3) = "cNF"
'    camposXML(4) = "natOp"
'    camposXML(5) = "cNF"
'    camposXML(6) = "mod"
'    camposXML(7) = "serie"
'    camposXML(8) = "nNF"
'    camposXML(9) = "dEmi"
'    camposXML(10) = "tpNF"
'    camposXML(11) = "cMunFG"
'    camposXML(12) = "tpImp"
'    camposXML(13) = "tpEmis"
'    camposXML(14) = "cDV"
'    camposXML(15) = "tpAmb"
'    camposXML(16) = "tpAmb"
'    camposXML(17) = "finNFE"
'    camposXML(18) = "procEmi"
'    camposXML(19) = "verProc"
'    camposXML(20) = "vNF"
'    camposXML(21) = "dVenc"
'    camposXML(22) = "dhRecbto"
'    camposXML(23) = "nProt"
'    camposXML(24) = "xMotivo"
'    camposXML(25) = "emit/CNPJ"
'    camposXML(26) = "emit/xNome"
'    camposXML(27) = "emit/xFant"
'    camposXML(28) = "emit/enderEmit/xLgr"
'    camposXML(29) = "emit/enderEmit/nro"
'    camposXML(30) = "emit/enderEmit/xBairro"
'    camposXML(31) = "emit/enderEmit/cMun"
'    camposXML(32) = "emit/enderEmit/xMun"
'    camposXML(33) = "emit/enderEmit/UF"
'    camposXML(34) = "emit/enderEmit/CEP"
'    camposXML(35) = "emit/enderEmit/cPais"
'    camposXML(36) = "emit/enderEmit/xPais"
'    camposXML(37) = "emit/IE"
'    camposXML(38) = "emit/CRT"
'    camposXML(39) = "dest/CNPJ"
'    camposXML(40) = "dest/xNome"
'    camposXML(41) = "dest/enderDest/xLgr"
'    camposXML(42) = "dest/enderDest/nro"
'    camposXML(43) = "dest/enderDest/xBairro"
'    camposXML(44) = "dest/enderDest/cMun"
'    camposXML(45) = "dest/enderDest/xMun"
'    camposXML(46) = "dest/enderDest/UF"
'    camposXML(47) = "dest/enderDest/CEP"
'    camposXML(48) = "dest/enderDest/cPais"
'    camposXML(49) = "dest/enderDest/xPais"
'    camposXML(50) = "dest/enderDest/fone"
'    camposXML(51) = "dest/IE"
'    camposXML(52) = "infAdic"
'    camposXML(53) = "vBCST"
'    camposXML(54) = "xPed"
'    camposXML(55) = "total/ICMSTot/vICMS"
'    camposXML(56) = "total/ICMSTot/vST"
'
'    'Desliga atualização de tela
'    Application.ScreenUpdating = True
'
'    'Importa os dados XML:
'
'    'O objeto DOMDocument deve ser usado para manipular dados XML:
'    Set xmlDOM = CreateObject("MSXML2.DOMDocument")
'
'    ' Retira a propriedade assincrona do objeto
'    xmlDOM.async = False
'
'    diretorio_extracao = ThisWorkbook.Path & "\"
'
'    cont_linha = 15
'    cont_linhaD = 2
'    ' Carrega o arquivo especificado para o objeto DOMDocument:
'    Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2) <> ""
'
'        xmlDOM.Load diretorio_extracao & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".xml"
'
'        cont_coluna = 1
'        For I = 1 To 56
'
'            On Error Resume Next
'            Set objNodes = xmlDOM.SelectNodes("//" & camposXML(I))
'            ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linhaD, cont_coluna) = objNodes.Item(0).Text
'
'
'            If Err.Number > 0 Then
'                ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linhaD, cont_coluna) = ""
'                Err.Number = 0
'            End If
'
'            cont_coluna = cont_coluna + 1
'
'        Next I
'
'        s = diretorio_extracao & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".xml"
'        If CaminhoExiste(s) = True Then
'        cont_linhaD = cont_linhaD + 1
'        End If
'        cont_linha = cont_linha + 1
'
'    Loop
'
'End Sub
'
'
'Sub importaXML2()
'
'
''Desativa warnings
'    Application.DisplayAlerts = False
'
'
'    cont_linha = 15
'    Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
'        ActiveWorkbook.XmlMaps("nfeProc_Map").Import Url:= _
'                                                     "C:\Users\angelo.r.perrone\Desktop\MACRO\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".xml"
'
'
'
'        cont_linha = cont_linha + 1
'    Loop
'
'End Sub
'
'Sub qtdItensXML()
'    Dim xmlDOM As DOMDocument    ' Variável de manipulação de documento
'    Dim objNodes As IXMLDOMNodeList    ' Lista de nós
'    Dim I As Integer
'    Dim nmItem As String
'
'    'O objeto DOMDocument deve ser usado para manipular dados XML:
'    Set xmlDOM = CreateObject("MSXML2.DOMDocument")
'
'    'Retira a propriedade assincrona do objeto
'    xmlDOM.async = False
'
'    'Diretorio
'    diretorio_extracao = ThisWorkbook.Path & "\"
'
'    'Lista os item da nota
'    cont_linha = 15
'    Do While ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 1).Value <> ""
'
'        soma_item = 0
'        For I = 0 To 100
'            xmlDOM.Load diretorio_extracao & ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 1) & ".xml"
'            Set objNodes = xmlDOM.SelectNodes("//det/prod/cProd")
'            nmItem = ""
'            On Error Resume Next
'            nmItem = objNodes.Item(I).Text
'
'            If nmItem <> "" Then
'                soma_item = soma_item + 1
'            End If
'
'        Next I
'
'        ThisWorkbook.Sheets("DADOS_XML_LOOP").Cells(cont_linha, 57).Value = soma_item
'
'
'        cont_linha = cont_linha + 1
'
'    Loop
'
'End Sub
'
'Sub integridadeXML()
'Dim s As String
'
'Dim camposVer(1 To 24) As Integer
'camposVer(1) = 8  '"nNF"
'camposVer(2) = 7  '"serie"
'camposVer(3) = 3  '"natOp"
'camposVer(4) = 10  '"tpNF"
'camposVer(5) = 9  '"dEmi"
'
'camposVer(6) = 25  '"CNPJ"     EMIT
'camposVer(7) = 26 '"xNome"     EMIT
'camposVer(8) = 28  '"xLgr"     EMIT
'camposVer(9) = 29  '"nro"     EMIT
'camposVer(10) = 30  '"xBairro"     EMIT
'camposVer(11) = 32  '"xMun"     EMIT
'camposVer(12) = 34  '"CEP"     EMIT
'camposVer(13) = 33  '"UF"     EMIT
'camposVer(14) = 37  '"IE"     EMIT
'
'camposVer(15) = 39  '"CNPJ"      rece
'camposVer(16) = 40 '"xNome"      rece
'camposVer(17) = 41 '"xLgr"      rece
'camposVer(18) = 42  '"nro"      rece
'camposVer(19) = 43  '"xBairro"      rece
'camposVer(20) = 45  '"xMun"      rece
'camposVer(21) = 47  '"CEP"      rece
'camposVer(22) = 46  '"UF"      rece
'camposVer(23) = 51  '"IE"      rece
'
'
'cont_linha = 15
'Do While ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value <> ""
'
's = ThisWorkbook.Path & "\" & ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value & ".xml"
'If CaminhoExiste(s) = True Then
'
'valorCont = 0
'valor = ""
'For I = 1 To 23
'valor = Application.VLookup(ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 2).Value, ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A1:BF10000"), camposVer(I), False)
'
'On Error Resume Next
'If valor = "" Then
'valorCont = valorCont + 1
'End If
'
'Next I
'
'    If valorCont > 0 Then
'
'    ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 5).Value = "NOK"
'
'    Else
'
'    ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 5).Value = "OK"
'
'    End If
'
'Else
'
'ThisWorkbook.Sheets("PROC_CODE").Cells(cont_linha, 5).Value = "NOK"
'
'End If
'
'cont_linha = cont_linha + 1
'Loop
'
'
'End Sub
