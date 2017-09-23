'Preenchimento de Formularios Web Utilizando o InternetExplorer

Private Sub CommandButton1_Click()

'Variavel que representa a planilha
Set sheet1 = Sheets("Planilha1")

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
IE.Document.All("element-35-10553764-7707338-1").Value = sheet1.Range("a2")
IE.Document.All("element-35-10553764-7707338-2").Value = sheet1.Range("b2")
IE.Document.All("element-35-10553764-7707338-3").Value = sheet1.Range("c2")
IE.Document.All("element-35-10553764-7707338-4").Value = sheet1.Range("d2")
IE.Document.All("element-35-10553764-7707338-5").Value = sheet1.Range("e2")
IE.Document.All("element-35-10553764-7707338-6").Value = sheet1.Range("f2")
IE.Document.All("element-35-10553764-7707338-7").Value = sheet1.Range("g2")
IE.Document.All("element-35-10553764-7707338-8").Value = sheet1.Range("h2")
'2 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707339-1").Value = sheet1.Range("a3")
IE.Document.All("element-35-10553764-7707339-2").Value = sheet1.Range("b3")
IE.Document.All("element-35-10553764-7707339-3").Value = sheet1.Range("c2")
IE.Document.All("element-35-10553764-7707339-4").Value = sheet1.Range("d2")
IE.Document.All("element-35-10553764-7707339-5").Value = sheet1.Range("e3")
IE.Document.All("element-35-10553764-7707339-6").Value = sheet1.Range("f3")
IE.Document.All("element-35-10553764-7707339-7").Value = sheet1.Range("g3")
IE.Document.All("element-35-10553764-7707339-8").Value = sheet1.Range("h3")
'3 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707340-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707340-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707340-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707340-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707340-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707340-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707340-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707340-8").Value = sheet1.Range("a4")
'4 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707341-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707341-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707341-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707341-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707341-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707341-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707341-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707341-8").Value = sheet1.Range("a4")
'5 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707342-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707342-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707342-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707342-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707342-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707342-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707342-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707342-8").Value = sheet1.Range("a4")
'6 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10553764-7707343-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707343-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707343-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707343-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707343-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707343-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707343-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553764-7707343-8").Value = sheet1.Range("a4")

''1
IE.Document.All("element-35-10553738-7707261-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553738-7707261-2").Value = sheet1.Range("a4")

''2
IE.Document.All("element-11-10553739").Value = sheet1.Range("a4")

''3
IE.Document.All("element-35-10553741-7707272-1").Value = sheet1.Range("a4")

''4
'1 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707262-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707262-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707262-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707262-4").Value = sheet1.Range("a4")
'2 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707263-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707263-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707263-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707263-4").Value = sheet1.Range("a4")
'3 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707264-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707264-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707264-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707264-4").Value = sheet1.Range("a4")
'4 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707265-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707265-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707265-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707265-4").Value = sheet1.Range("a4")
'5 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707266-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707266-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707266-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707266-4").Value = sheet1.Range("a4")
'6 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707267-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707267-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707267-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707267-4").Value = sheet1.Range("a4")
'7 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707268-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707268-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707268-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707268-4").Value = sheet1.Range("a4")
'8 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707269-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707269-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707269-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707269-4").Value = sheet1.Range("a4")
'9 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707270-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707270-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707270-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707270-4").Value = sheet1.Range("a4")
'10 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707271-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707271-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707271-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553740-7707271-4").Value = sheet1.Range("a4")

''5
'1 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707273-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553742-7707273-2").Value = sheet1.Range("a4")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707274-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553742-7707274-2").Value = sheet1.Range("a4")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707275-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553742-7707275-2").Value = sheet1.Range("a4")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707276-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553742-7707276-2").Value = sheet1.Range("a4")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707277-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553742-7707277-2").Value = sheet1.Range("a4")

''7
'1 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707278-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553744-7707278-2").Value = sheet1.Range("a4")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707279-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553744-7707279-2").Value = sheet1.Range("a4")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707280-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553744-7707280-2").Value = sheet1.Range("a4")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707281-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553744-7707281-2").Value = sheet1.Range("a4")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553744-7707282-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553744-7707282-2").Value = sheet1.Range("a4")

''8
'1 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707283-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707283-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707283-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707283-4").Value = sheet1.Range("a4")
'2 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707284-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707284-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707284-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707284-4").Value = sheet1.Range("a4")
'3 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707285-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707285-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707285-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707285-4").Value = sheet1.Range("a4")
'4 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707286-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707286-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707286-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707286-4").Value = sheet1.Range("a4")
'4 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707287-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707287-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707287-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553745-7707287-4").Value = sheet1.Range("a4")


''11
'1 Ano 1 QtdRegs 2
IE.Document.All("element-34-10553748-7707288-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707288-2").Value = sheet1.Range("a4")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707289-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707289-2").Value = sheet1.Range("a4")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707290-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707290-2").Value = sheet1.Range("a4")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707291-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707291-2").Value = sheet1.Range("a4")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707292-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707292-2").Value = sheet1.Range("a4")
'6 Ano 1 QtdRegs 2
IE.Document.All("element-34-10553748-7707293-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707293-2").Value = sheet1.Range("a4")
'7 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707294-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707294-2").Value = sheet1.Range("a4")
'8 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707295-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707295-2").Value = sheet1.Range("a4")
'9 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707296-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707296-2").Value = sheet1.Range("a4")
'10 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707297-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707297-2").Value = sheet1.Range("a4")
'11 Ano 1 QtdRegs 2
IE.Document.All("element-34-10553748-7707298-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707298-2").Value = sheet1.Range("a4")
'12 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707299-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707299-2").Value = sheet1.Range("a4")
'13 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707300-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707300-2").Value = sheet1.Range("a4")
'14 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707301-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707301-2").Value = sheet1.Range("a4")
'15 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707302-1").Value = sheet1.Range("a4")
IE.Document.All("element-34-10553748-7707302-2").Value = sheet1.Range("a4")

''12
'1 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707303-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553749-7707303-2").Value = sheet1.Range("a4")
'2 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707304-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553749-7707304-2").Value = sheet1.Range("a4")
'3 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707305-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553749-7707305-2").Value = sheet1.Range("a4")
'4 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707306-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553749-7707306-2").Value = sheet1.Range("a4")
'5 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707307-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553749-7707307-2").Value = sheet1.Range("a4")

''16
'1 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707308-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553753-7707308-2").Value = sheet1.Range("a4")
'2 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707309-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553753-7707309-2").Value = sheet1.Range("a4")
'3 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707310-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553753-7707310-2").Value = sheet1.Range("a4")
'4 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707311-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553753-7707311-2").Value = sheet1.Range("a4")
'5 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707312-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553753-7707312-2").Value = sheet1.Range("a4")

''18
'1 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553755-7707313-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553755-7707313-2").Value = sheet1.Range("a4")
'2 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553755-7707314-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553755-7707314-2").Value = sheet1.Range("a4")
'3 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553755-7707315-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553755-7707315-2").Value = sheet1.Range("a4")
'4 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553755-7707316-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553755-7707316-2").Value = sheet1.Range("a4")
'5 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553755-7707317-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553755-7707317-2").Value = sheet1.Range("a4")

''19
'1 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553756-7707318-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10553756-7707318-2").Value = sheet1.Range("a4")


End Sub
