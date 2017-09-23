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

''Infos Equipe
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

''1
IE.Document.All("element-35-10553738-7707261-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553738-7707261-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

''2
IE.Document.All("element-11-10553739").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

''3
IE.Document.All("element-35-10553741-7707272-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

''4
'1 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707262-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707262-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707262-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707262-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'2 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707263-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707263-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707263-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707263-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'3 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707264-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707264-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707264-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707264-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'4 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707265-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707265-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707265-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707265-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'5 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707266-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707266-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707266-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707266-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'6 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707267-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707267-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707267-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707267-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'7 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707268-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707268-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707268-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707268-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'8 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707269-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707269-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707269-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707269-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'9 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707270-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707270-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707270-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707270-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'10 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707271-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707271-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707271-3").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553740-7707271-4").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

''5
'1 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707273-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553742-7707273-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707274-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553742-7707274-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707275-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553742-7707275-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707276-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553742-7707276-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707277-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553742-7707277-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")

''7
'1 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707278-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553744-7707278-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707279-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553744-7707279-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707280-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553744-7707280-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707281-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553744-7707281-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707282-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")
IE.Document.All("element-35-10553744-7707282-2").Value = ThisWorkbook.Sheets("Planilha1").Range("a4")


End Sub
