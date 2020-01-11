Attribute VB_Name = "Module5"
Sub formatacao()

'Limpa tabelas antes da execução
Call limpaXML
Call limpaPROC_CODE
End Sub


Sub limpaPROC_CODE()

ThisWorkbook.Sheets("PROC_CODE").Range("B15:K20000").Clear
ThisWorkbook.Sheets("PROC_CODE").Range("B15:K20000").NumberFormat = "@"

End Sub

Sub limpaXML()

ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A2:BF20000").Clear
ThisWorkbook.Sheets("DADOS_XML_LOOP").Range("A2:BF20000").NumberFormat = "@"

End Sub

