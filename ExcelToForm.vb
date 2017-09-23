'Preenchimento de Formularios Web Utilizando o InternetExplorer

Sub PreencheCampos()

'Abre a janela do Internet Explorer
Dim IE As Object
Set IE = CreateObject("InternetExplorer.Application")
'Site do formulario que se deseja preencher
IE.Navigate "https://www.onlinepesquisa.com/s/cbc9679"
IE.Visible = True

'Espera o carregamento do site
While IE.Busy
DoEvents
Wend

'Preenche os campos (Inspecionar Elemento)

'1 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707338-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a2")
IE.Document.All("element-35-10553764-7707338-2").Value = ThisWorkbook.Sheets("Planilha1").Range("b2")
IE.Document.All("element-35-10553764-7707338-3").Value = ThisWorkbook.Sheets("Planilha1").Range("c2")
IE.Document.All("element-35-10553764-7707338-4").Value = ThisWorkbook.Sheets("Planilha1").Range("d2")
IE.Document.All("element-35-10553764-7707338-5").Value = ThisWorkbook.Sheets("Planilha1").Range("e2")
IE.Document.All("element-35-10553764-7707338-6").Value = ThisWorkbook.Sheets("Planilha1").Range("f2")
IE.Document.All("element-35-10553764-7707338-7").Value = ThisWorkbook.Sheets("Planilha1").Range("g2")
IE.Document.All("element-35-10553764-7707338-8").Value = ThisWorkbook.Sheets("Planilha1").Range("h2")

'2 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707339-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a3")
IE.Document.All("element-35-10553764-7707339-2").Value = ThisWorkbook.Sheets("Planilha1").Range("b3")
IE.Document.All("element-35-10553764-7707339-3").Value = ThisWorkbook.Sheets("Planilha1").Range("c2")
IE.Document.All("element-35-10553764-7707339-4").Value = ThisWorkbook.Sheets("Planilha1").Range("d2")
IE.Document.All("element-35-10553764-7707339-5").Value = ThisWorkbook.Sheets("Planilha1").Range("e3")
IE.Document.All("element-35-10553764-7707339-6").Value = ThisWorkbook.Sheets("Planilha1").Range("f3")
IE.Document.All("element-35-10553764-7707339-7").Value = ThisWorkbook.Sheets("Planilha1").Range("g3")
IE.Document.All("element-35-10553764-7707339-8").Value = ThisWorkbook.Sheets("Planilha1").Range("h3")

'3 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707340-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707340-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707340-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707340-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707340-5").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707340-6").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707340-7").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707340-8").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

'4 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707341-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707341-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707341-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707341-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707341-5").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707341-6").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707341-7").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707341-8").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

'5 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707342-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707342-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707342-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707342-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707342-5").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707342-6").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707342-7").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707342-8").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

'6 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707343-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707343-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707343-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707343-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707343-5").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707343-6").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707343-7").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553764-7707343-8").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

End Sub
