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
IE.Document.All("element-35-10553738-7707261-1").Value = sheet1.Range("a7")
IE.Document.All("element-35-10553738-7707261-2").Value = sheet1.Range("b7")

''2
IE.Document.All("element-11-10553739").Value = sheet1.Range("a10")

''3
IE.Document.All("element-35-10553741-7707272-1").Value = sheet1.Range("a13")

''4
'1 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707262-1").Value = sheet1.Range("a16")
IE.Document.All("element-35-10553740-7707262-2").Value = sheet1.Range("b16")
IE.Document.All("element-35-10553740-7707262-3").Value = sheet1.Range("c16")
IE.Document.All("element-35-10553740-7707262-4").Value = sheet1.Range("d16")
'2 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707263-1").Value = sheet1.Range("a17")
IE.Document.All("element-35-10553740-7707263-2").Value = sheet1.Range("b17")
IE.Document.All("element-35-10553740-7707263-3").Value = sheet1.Range("c17")
IE.Document.All("element-35-10553740-7707263-4").Value = sheet1.Range("d17")
'3 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707264-1").Value = sheet1.Range("a18")
IE.Document.All("element-35-10553740-7707264-2").Value = sheet1.Range("b18")
IE.Document.All("element-35-10553740-7707264-3").Value = sheet1.Range("c18")
IE.Document.All("element-35-10553740-7707264-4").Value = sheet1.Range("d18")
'4 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707265-1").Value = sheet1.Range("a19")
IE.Document.All("element-35-10553740-7707265-2").Value = sheet1.Range("b19")
IE.Document.All("element-35-10553740-7707265-3").Value = sheet1.Range("c19")
IE.Document.All("element-35-10553740-7707265-4").Value = sheet1.Range("d19")
'5 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707266-1").Value = sheet1.Range("a20")
IE.Document.All("element-35-10553740-7707266-2").Value = sheet1.Range("b20")
IE.Document.All("element-35-10553740-7707266-3").Value = sheet1.Range("c20")
IE.Document.All("element-35-10553740-7707266-4").Value = sheet1.Range("d20")
'6 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707267-1").Value = sheet1.Range("a21")
IE.Document.All("element-35-10553740-7707267-2").Value = sheet1.Range("b21")
IE.Document.All("element-35-10553740-7707267-3").Value = sheet1.Range("c21")
IE.Document.All("element-35-10553740-7707267-4").Value = sheet1.Range("d21")
'7 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707268-1").Value = sheet1.Range("a22")
IE.Document.All("element-35-10553740-7707268-2").Value = sheet1.Range("b22")
IE.Document.All("element-35-10553740-7707268-3").Value = sheet1.Range("c22")
IE.Document.All("element-35-10553740-7707268-4").Value = sheet1.Range("d22")
'8 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707269-1").Value = sheet1.Range("a23")
IE.Document.All("element-35-10553740-7707269-2").Value = sheet1.Range("b23")
IE.Document.All("element-35-10553740-7707269-3").Value = sheet1.Range("c23")
IE.Document.All("element-35-10553740-7707269-4").Value = sheet1.Range("d23")
'9 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707270-1").Value = sheet1.Range("a24")
IE.Document.All("element-35-10553740-7707270-2").Value = sheet1.Range("b24")
IE.Document.All("element-35-10553740-7707270-3").Value = sheet1.Range("c24")
IE.Document.All("element-35-10553740-7707270-4").Value = sheet1.Range("d24")
'10 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10553740-7707271-1").Value = sheet1.Range("a25")
IE.Document.All("element-35-10553740-7707271-2").Value = sheet1.Range("b25")
IE.Document.All("element-35-10553740-7707271-3").Value = sheet1.Range("c25")
IE.Document.All("element-35-10553740-7707271-4").Value = sheet1.Range("d25")

''5
'1 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707273-1").Value = sheet1.Range("a28")
IE.Document.All("element-35-10553742-7707273-2").Value = sheet1.Range("b28")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707274-1").Value = sheet1.Range("a29")
IE.Document.All("element-35-10553742-7707274-2").Value = sheet1.Range("b29")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707275-1").Value = sheet1.Range("a30")
IE.Document.All("element-35-10553742-7707275-2").Value = sheet1.Range("b30")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707276-1").Value = sheet1.Range("a31")
IE.Document.All("element-35-10553742-7707276-2").Value = sheet1.Range("b31")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-35-10553742-7707277-1").Value = sheet1.Range("a32")
IE.Document.All("element-35-10553742-7707277-2").Value = sheet1.Range("b32")

''7
'1 Autor 1 QtdCits 2
IE.Document.All("element-35-10553744-7707278-1").Value = sheet1.Range("a35")
IE.Document.All("element-35-10553744-7707278-2").Value = sheet1.Range("b35")
'2 Autor 1 QtdCits 2
IE.Document.All("element-35-10553744-7707279-1").Value = sheet1.Range("a36")
IE.Document.All("element-35-10553744-7707279-2").Value = sheet1.Range("b36")
'3 Autor 1 QtdCits 2
IE.Document.All("element-35-10553744-7707280-1").Value = sheet1.Range("a37")
IE.Document.All("element-35-10553744-7707280-2").Value = sheet1.Range("b37")
'4 Autor 1 QtdCits 2
IE.Document.All("element-35-10553744-7707281-1").Value = sheet1.Range("a38")
IE.Document.All("element-35-10553744-7707281-2").Value = sheet1.Range("b38")
'5 Autor 1 QtdCits 2
IE.Document.All("element-35-10553744-7707282-1").Value = sheet1.Range("a39")
IE.Document.All("element-35-10553744-7707282-2").Value = sheet1.Range("b39")

''8
'1 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707283-1").Value = sheet1.Range("a42")
IE.Document.All("element-35-10553745-7707283-2").Value = sheet1.Range("b42")
IE.Document.All("element-35-10553745-7707283-3").Value = sheet1.Range("c42")
IE.Document.All("element-35-10553745-7707283-4").Value = sheet1.Range("d42")
'2 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707284-1").Value = sheet1.Range("a43")
IE.Document.All("element-35-10553745-7707284-2").Value = sheet1.Range("b43")
IE.Document.All("element-35-10553745-7707284-3").Value = sheet1.Range("c43")
IE.Document.All("element-35-10553745-7707284-4").Value = sheet1.Range("d43")
'3 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707285-1").Value = sheet1.Range("a44")
IE.Document.All("element-35-10553745-7707285-2").Value = sheet1.Range("b44")
IE.Document.All("element-35-10553745-7707285-3").Value = sheet1.Range("c44")
IE.Document.All("element-35-10553745-7707285-4").Value = sheet1.Range("d44")
'4 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707286-1").Value = sheet1.Range("a45")
IE.Document.All("element-35-10553745-7707286-2").Value = sheet1.Range("b45")
IE.Document.All("element-35-10553745-7707286-3").Value = sheet1.Range("c45")
IE.Document.All("element-35-10553745-7707286-4").Value = sheet1.Range("d45")
'4 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10553745-7707287-1").Value = sheet1.Range("a46")
IE.Document.All("element-35-10553745-7707287-2").Value = sheet1.Range("b46")
IE.Document.All("element-35-10553745-7707287-3").Value = sheet1.Range("c46")
IE.Document.All("element-35-10553745-7707287-4").Value = sheet1.Range("d46")


''11
'1 Ano 1 QtdRegs 2
IE.Document.All("element-34-10553748-7707288-1").Value = sheet1.Range("a49")
IE.Document.All("element-34-10553748-7707288-2").Value = sheet1.Range("b49")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707289-1").Value = sheet1.Range("a50")
IE.Document.All("element-34-10553748-7707289-2").Value = sheet1.Range("b50")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707290-1").Value = sheet1.Range("a51")
IE.Document.All("element-34-10553748-7707290-2").Value = sheet1.Range("b51")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707291-1").Value = sheet1.Range("a52")
IE.Document.All("element-34-10553748-7707291-2").Value = sheet1.Range("b52")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707292-1").Value = sheet1.Range("a53")
IE.Document.All("element-34-10553748-7707292-2").Value = sheet1.Range("b53")
'6 Ano 1 QtdRegs 2
IE.Document.All("element-34-10553748-7707293-1").Value = sheet1.Range("a54")
IE.Document.All("element-34-10553748-7707293-2").Value = sheet1.Range("b54")
'7 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707294-1").Value = sheet1.Range("a55")
IE.Document.All("element-34-10553748-7707294-2").Value = sheet1.Range("b55")
'8 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707295-1").Value = sheet1.Range("a56")
IE.Document.All("element-34-10553748-7707295-2").Value = sheet1.Range("b56")
'9 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707296-1").Value = sheet1.Range("a57")
IE.Document.All("element-34-10553748-7707296-2").Value = sheet1.Range("b57")
'10 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707297-1").Value = sheet1.Range("a58")
IE.Document.All("element-34-10553748-7707297-2").Value = sheet1.Range("b58")
'11 Ano 1 QtdRegs 2
IE.Document.All("element-34-10553748-7707298-1").Value = sheet1.Range("a59")
IE.Document.All("element-34-10553748-7707298-2").Value = sheet1.Range("b59")
'12 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707299-1").Value = sheet1.Range("a60")
IE.Document.All("element-34-10553748-7707299-2").Value = sheet1.Range("b60")
'13 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707300-1").Value = sheet1.Range("a61")
IE.Document.All("element-34-10553748-7707300-2").Value = sheet1.Range("b61")
'14 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707301-1").Value = sheet1.Range("a62")
IE.Document.All("element-34-10553748-7707301-2").Value = sheet1.Range("b62")
'15 Autor 1 QtdPubs 2
IE.Document.All("element-34-10553748-7707302-1").Value = sheet1.Range("a63")
IE.Document.All("element-34-10553748-7707302-2").Value = sheet1.Range("b63")

''12
'1 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707303-1").Value = sheet1.Range("a66")
IE.Document.All("element-35-10553749-7707303-2").Value = sheet1.Range("b66")
'2 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707304-1").Value = sheet1.Range("a67")
IE.Document.All("element-35-10553749-7707304-2").Value = sheet1.Range("b67")
'3 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707305-1").Value = sheet1.Range("a68")
IE.Document.All("element-35-10553749-7707305-2").Value = sheet1.Range("b68")
'4 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707306-1").Value = sheet1.Range("a69")
IE.Document.All("element-35-10553749-7707306-2").Value = sheet1.Range("b69")
'5 Pais 1 QtdPubs 2
IE.Document.All("element-35-10553749-7707307-1").Value = sheet1.Range("a70")
IE.Document.All("element-35-10553749-7707307-2").Value = sheet1.Range("b70")

''16
'1 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707308-1").Value = sheet1.Range("a73")
IE.Document.All("element-35-10553753-7707308-2").Value = sheet1.Range("b73")
'2 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707309-1").Value = sheet1.Range("a74")
IE.Document.All("element-35-10553753-7707309-2").Value = sheet1.Range("b74")
'3 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707310-1").Value = sheet1.Range("a75")
IE.Document.All("element-35-10553753-7707310-2").Value = sheet1.Range("b75")
'4 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707311-1").Value = sheet1.Range("a76")
IE.Document.All("element-35-10553753-7707311-2").Value = sheet1.Range("b76")
'5 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10553753-7707312-1").Value = sheet1.Range("a77")
IE.Document.All("element-35-10553753-7707312-2").Value = sheet1.Range("b77")

''18
'1 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10553755-7707313-1").Value = sheet1.Range("a80")
IE.Document.All("element-35-10553755-7707313-2").Value = sheet1.Range("b80")
'2 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10553755-7707314-1").Value = sheet1.Range("a81")
IE.Document.All("element-35-10553755-7707314-2").Value = sheet1.Range("b81")
'3 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10553755-7707315-1").Value = sheet1.Range("a82")
IE.Document.All("element-35-10553755-7707315-2").Value = sheet1.Range("b82")
'4 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10553755-7707316-1").Value = sheet1.Range("a83")
IE.Document.All("element-35-10553755-7707316-2").Value = sheet1.Range("b83")
'5 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10553755-7707317-1").Value = sheet1.Range("a84")
IE.Document.All("element-35-10553755-7707317-2").Value = sheet1.Range("b84")

''19
'1 Congresso 1 Regs 2
IE.Document.All("element-35-10553756-7707318-1").Value = sheet1.Range("a87")
IE.Document.All("element-35-10553756-7707318-2").Value = sheet1.Range("b87")
'2 Congresso 1 Regs 2
IE.Document.All("element-35-10553756-7707319-1").Value = sheet1.Range("a88")
IE.Document.All("element-35-10553756-7707319-2").Value = sheet1.Range("b88")
'3 Congresso 1 Regs 2
IE.Document.All("element-35-10553756-7707320-1").Value = sheet1.Range("a89")
IE.Document.All("element-35-10553756-7707320-2").Value = sheet1.Range("b89")
'4 Congresso 1 Regs 2
IE.Document.All("element-35-10553756-7707321-1").Value = sheet1.Range("a90")
IE.Document.All("element-35-10553756-7707321-2").Value = sheet1.Range("b90")
'5 Congresso 1 Regs 2
IE.Document.All("element-35-10553756-7707322-1").Value = sheet1.Range("a91")
IE.Document.All("element-35-10553756-7707322-2").Value = sheet1.Range("b91")

''20
'1 Universidade 1 Regs 2
IE.Document.All("element-35-10553757-7707323-1").Value = sheet1.Range("a94")
IE.Document.All("element-35-10553757-7707323-2").Value = sheet1.Range("b94")
'2 Universidade 1 Regs 2
IE.Document.All("element-35-10553757-7707324-1").Value = sheet1.Range("a95")
IE.Document.All("element-35-10553757-7707324-2").Value = sheet1.Range("b95")
'3 Universidade 1 Regs 2
IE.Document.All("element-35-10553757-7707325-1").Value = sheet1.Range("a96")
IE.Document.All("element-35-10553757-7707325-2").Value = sheet1.Range("b96")
'4 Universidade 1 Regs 2
IE.Document.All("element-35-10553757-7707326-1").Value = sheet1.Range("a97")
IE.Document.All("element-35-10553757-7707326-2").Value = sheet1.Range("b97")
'5 Universidade 1 Regs 2
IE.Document.All("element-35-10553757-7707327-1").Value = sheet1.Range("a98")
IE.Document.All("element-35-10553757-7707327-2").Value = sheet1.Range("b98")

''22
'1 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707328-1").Value = sheet1.Range("a101")
IE.Document.All("element-35-10553759-7707328-2").Value = sheet1.Range("b101")
'2 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707329-1").Value = sheet1.Range("a102")
IE.Document.All("element-35-10553759-7707329-2").Value = sheet1.Range("b102")
'3 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707330-1").Value = sheet1.Range("a103")
IE.Document.All("element-35-10553759-7707330-2").Value = sheet1.Range("b103")
'4 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707331-1").Value = sheet1.Range("a104")
IE.Document.All("element-35-10553759-7707331-2").Value = sheet1.Range("b104")
'5 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707332-1").Value = sheet1.Range("a105")
IE.Document.All("element-35-10553759-7707332-2").Value = sheet1.Range("b105")
'6 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707333-1").Value = sheet1.Range("a106")
IE.Document.All("element-35-10553759-7707333-2").Value = sheet1.Range("b106")
'7 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707334-1").Value = sheet1.Range("a107")
IE.Document.All("element-35-10553759-7707334-2").Value = sheet1.Range("b107")
'8 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707335-1").Value = sheet1.Range("a108")
IE.Document.All("element-35-10553759-7707335-2").Value = sheet1.Range("b108")
'9 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707336-1").Value = sheet1.Range("a109")
IE.Document.All("element-35-10553759-7707336-2").Value = sheet1.Range("b109")
'10 PalavraChave 1 Regs 2
IE.Document.All("element-35-10553759-7707337-1").Value = sheet1.Range("a110")
IE.Document.All("element-35-10553759-7707337-2").Value = sheet1.Range("b110")

End Sub
