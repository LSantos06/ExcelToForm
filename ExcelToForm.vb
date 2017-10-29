'Preenchimento de Formularios Web Utilizando o InternetExplorer

Private Sub PreencherFormulario_Click()

'Variavel que representa a planilha
Set sheet1 = Sheets("Planilha1")

'Abre a janela do Internet Explorer
Dim IE As Object
Set IE = CreateObject("InternetExplorer.Application")
'Site do formulario que se deseja preencher
IE.Navigate "https://www.onlinepesquisa.com/s/2dfbb71"
IE.Visible = True

'Espera o carregamento do site
While IE.Busy
DoEvents
Wend

'Preenche os campos (Inspecionar Elemento)

''Infos Equipe
'1 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10835283-7896841-1").Value = sheet1.Range("a2")
IE.Document.All("element-35-10835283-7896841-2").Value = sheet1.Range("b2")
IE.Document.All("element-35-10835283-7896841-3").Value = sheet1.Range("c2")
IE.Document.All("element-35-10835283-7896841-4").Value = sheet1.Range("d2")
IE.Document.All("element-35-10835283-7896841-5").Value = sheet1.Range("e2")
IE.Document.All("element-35-10835283-7896841-6").Value = sheet1.Range("f2")
IE.Document.All("element-35-10835283-7896841-7").Value = sheet1.Range("g2")
IE.Document.All("element-35-10835283-7896841-8").Value = sheet1.Range("h2")
'2 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10835283-7896842-1").Value = sheet1.Range("a3")
IE.Document.All("element-35-10835283-7896842-2").Value = sheet1.Range("b3")
IE.Document.All("element-35-10835283-7896842-3").Value = sheet1.Range("c2")
IE.Document.All("element-35-10835283-7896842-4").Value = sheet1.Range("d2")
IE.Document.All("element-35-10835283-7896842-5").Value = sheet1.Range("e3")
IE.Document.All("element-35-10835283-7896842-6").Value = sheet1.Range("f3")
IE.Document.All("element-35-10835283-7896842-7").Value = sheet1.Range("g3")
IE.Document.All("element-35-10835283-7896842-8").Value = sheet1.Range("h3")
'3 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10835283-7896843-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896843-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896843-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896843-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896843-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896843-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896843-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896843-8").Value = sheet1.Range("a4")
'4 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10835283-7896844-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896844-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896844-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896844-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896844-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896844-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896844-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896844-8").Value = sheet1.Range("a4")
'5 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10835283-7896845-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896845-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896845-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896845-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896845-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896845-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896845-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896845-8").Value = sheet1.Range("a4")
'6 Nome 1 Matricula 2 Curso 3 Id da equipe 4 CPF 5 RG 6 email 7 tel 8
IE.Document.All("element-35-10835283-7896846-1").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896846-2").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896846-3").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896846-4").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896846-5").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896846-6").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896846-7").Value = sheet1.Range("a4")
IE.Document.All("element-35-10835283-7896846-8").Value = sheet1.Range("a4")

''1
IE.Document.All("element-35-10835256-7896764-1").Value = sheet1.Range("a7")
IE.Document.All("element-35-10835256-7896764-2").Value = sheet1.Range("b7")

''2
IE.Document.All("element-11-10835257").Value = sheet1.Range("a10")

''3
IE.Document.All("element-35-10835259-7896775-1").Value = sheet1.Range("a13")

''4
'1 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7896765-1").Value = sheet1.Range("a16")
IE.Document.All("element-35-10835258-7896765-2").Value = sheet1.Range("b16")
IE.Document.All("element-35-10835258-7896765-3").Value = sheet1.Range("c16")
IE.Document.All("element-35-10835258-7896765-4").Value = sheet1.Range("d16")
'2 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7896766-1").Value = sheet1.Range("a17")
IE.Document.All("element-35-10835258-7896766-2").Value = sheet1.Range("b17")
IE.Document.All("element-35-10835258-7896766-3").Value = sheet1.Range("c17")
IE.Document.All("element-35-10835258-7896766-4").Value = sheet1.Range("d17")
'3 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7896767-1").Value = sheet1.Range("a18")
IE.Document.All("element-35-10835258-7896767-2").Value = sheet1.Range("b18")
IE.Document.All("element-35-10835258-7896767-3").Value = sheet1.Range("c18")
IE.Document.All("element-35-10835258-7896767-4").Value = sheet1.Range("d18")
'4 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7896768-1").Value = sheet1.Range("a19")
IE.Document.All("element-35-10835258-7896768-2").Value = sheet1.Range("b19")
IE.Document.All("element-35-10835258-7896768-3").Value = sheet1.Range("c19")
IE.Document.All("element-35-10835258-7896768-4").Value = sheet1.Range("d19")
'5 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7896769-1").Value = sheet1.Range("a20")
IE.Document.All("element-35-10835258-7896769-2").Value = sheet1.Range("b20")
IE.Document.All("element-35-10835258-7896769-3").Value = sheet1.Range("c20")
IE.Document.All("element-35-10835258-7896769-4").Value = sheet1.Range("d20")
'6 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7896770-1").Value = sheet1.Range("a21")
IE.Document.All("element-35-10835258-7707270-2").Value = sheet1.Range("b21")
IE.Document.All("element-35-10835258-7707270-3").Value = sheet1.Range("c21")
IE.Document.All("element-35-10835258-7707270-4").Value = sheet1.Range("d21")
'7 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7707271-1").Value = sheet1.Range("a22")
IE.Document.All("element-35-10835258-7707271-2").Value = sheet1.Range("b22")
IE.Document.All("element-35-10835258-7707271-3").Value = sheet1.Range("c22")
IE.Document.All("element-35-10835258-7707271-4").Value = sheet1.Range("d22")
'8 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7707272-1").Value = sheet1.Range("a23")
IE.Document.All("element-35-10835258-7707272-2").Value = sheet1.Range("b23")
IE.Document.All("element-35-10835258-7707272-3").Value = sheet1.Range("c23")
IE.Document.All("element-35-10835258-7707272-4").Value = sheet1.Range("d23")
'9 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7707273-1").Value = sheet1.Range("a24")
IE.Document.All("element-35-10835258-7707273-2").Value = sheet1.Range("b24")
IE.Document.All("element-35-10835258-7707273-3").Value = sheet1.Range("c24")
IE.Document.All("element-35-10835258-7707273-4").Value = sheet1.Range("d24")
'10 Eixo1Nome 1 Eixo1Fator 2 Eixo2Nome 3 Eixo2Fator 4
IE.Document.All("element-35-10835258-7707274-1").Value = sheet1.Range("a25")
IE.Document.All("element-35-10835258-7707274-2").Value = sheet1.Range("b25")
IE.Document.All("element-35-10835258-7707274-3").Value = sheet1.Range("c25")
IE.Document.All("element-35-10835258-7707274-4").Value = sheet1.Range("d25")

''5
'1 Autor 1 QtdPubs 2
IE.Document.All("element-35-10835260-7896776-1").Value = sheet1.Range("a28")
IE.Document.All("element-35-10835260-7896776-2").Value = sheet1.Range("b28")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-35-10835260-7896777-1").Value = sheet1.Range("a29")
IE.Document.All("element-35-10835260-7896777-2").Value = sheet1.Range("b29")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-35-10835260-7896778-1").Value = sheet1.Range("a30")
IE.Document.All("element-35-10835260-7896778-2").Value = sheet1.Range("b30")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-35-10835260-7896779-1").Value = sheet1.Range("a31")
IE.Document.All("element-35-10835260-7896779-2").Value = sheet1.Range("b31")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-35-10835260-7896780-1").Value = sheet1.Range("a32")
IE.Document.All("element-35-10835260-7896780-2").Value = sheet1.Range("b32")

''7
'1 Autor 1 QtdCits 2
IE.Document.All("element-35-10835262-7896781-1").Value = sheet1.Range("a35")
IE.Document.All("element-35-10835262-7896781-2").Value = sheet1.Range("b35")
'2 Autor 1 QtdCits 2
IE.Document.All("element-35-10835262-7896782-1").Value = sheet1.Range("a36")
IE.Document.All("element-35-10835262-7896782-2").Value = sheet1.Range("b36")
'3 Autor 1 QtdCits 2
IE.Document.All("element-35-10835262-7896783-1").Value = sheet1.Range("a37")
IE.Document.All("element-35-10835262-7896783-2").Value = sheet1.Range("b37")
'4 Autor 1 QtdCits 2
IE.Document.All("element-35-10835262-7896784-1").Value = sheet1.Range("a38")
IE.Document.All("element-35-10835262-7896784-2").Value = sheet1.Range("b38")
'5 Autor 1 QtdCits 2
IE.Document.All("element-35-10835262-7896785-1").Value = sheet1.Range("a39")
IE.Document.All("element-35-10835262-7896785-2").Value = sheet1.Range("b39")

''8
'1 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10835263-7896786-1").Value = sheet1.Range("a42")
IE.Document.All("element-35-10835263-7896786-2").Value = sheet1.Range("b42")
IE.Document.All("element-35-10835263-7896786-3").Value = sheet1.Range("c42")
IE.Document.All("element-35-10835263-7896786-4").Value = sheet1.Range("d42")
'2 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10835263-7896787-1").Value = sheet1.Range("a43")
IE.Document.All("element-35-10835263-7896787-2").Value = sheet1.Range("b43")
IE.Document.All("element-35-10835263-7896787-3").Value = sheet1.Range("c43")
IE.Document.All("element-35-10835263-7896787-4").Value = sheet1.Range("d43")
'3 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10835263-7896788-1").Value = sheet1.Range("a44")
IE.Document.All("element-35-10835263-7896788-2").Value = sheet1.Range("b44")
IE.Document.All("element-35-10835263-7896788-3").Value = sheet1.Range("c44")
IE.Document.All("element-35-10835263-7896788-4").Value = sheet1.Range("d44")
'4 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10835263-7896789-1").Value = sheet1.Range("a45")
IE.Document.All("element-35-10835263-7896789-2").Value = sheet1.Range("b45")
IE.Document.All("element-35-10835263-7896789-3").Value = sheet1.Range("c45")
IE.Document.All("element-35-10835263-7896789-4").Value = sheet1.Range("d45")
'5 NomeTrab 1 Ano 2 Autor 3 QtdeCitacoes 4
IE.Document.All("element-35-10835263-7896790-1").Value = sheet1.Range("a46")
IE.Document.All("element-35-10835263-7896790-2").Value = sheet1.Range("b46")
IE.Document.All("element-35-10835263-7896790-3").Value = sheet1.Range("c46")
IE.Document.All("element-35-10835263-7896790-4").Value = sheet1.Range("d46")


''11
'1 Ano 1 QtdRegs 2
IE.Document.All("element-34-10835266-7896791-1").Value = sheet1.Range("a49")
IE.Document.All("element-34-10835266-7896791-2").Value = sheet1.Range("b49")
'2 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896792-1").Value = sheet1.Range("a50")
IE.Document.All("element-34-10835266-7896792-2").Value = sheet1.Range("b50")
'3 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896793-1").Value = sheet1.Range("a51")
IE.Document.All("element-34-10835266-7896793-2").Value = sheet1.Range("b51")
'4 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896794-1").Value = sheet1.Range("a52")
IE.Document.All("element-34-10835266-7896794-2").Value = sheet1.Range("b52")
'5 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896795-1").Value = sheet1.Range("a53")
IE.Document.All("element-34-10835266-7896795-2").Value = sheet1.Range("b53")
'6 Ano 1 QtdRegs 2
IE.Document.All("element-34-10835266-7896796-1").Value = sheet1.Range("a54")
IE.Document.All("element-34-10835266-7896796-2").Value = sheet1.Range("b54")
'7 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896797-1").Value = sheet1.Range("a55")
IE.Document.All("element-34-10835266-7896797-2").Value = sheet1.Range("b55")
'8 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896798-1").Value = sheet1.Range("a56")
IE.Document.All("element-34-10835266-7896798-2").Value = sheet1.Range("b56")
'9 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896799-1").Value = sheet1.Range("a57")
IE.Document.All("element-34-10835266-7896799-2").Value = sheet1.Range("b57")
'10 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896800-1").Value = sheet1.Range("a58")
IE.Document.All("element-34-10835266-7896800-2").Value = sheet1.Range("b58")
'11 Ano 1 QtdRegs 2
IE.Document.All("element-34-10835266-7896801-1").Value = sheet1.Range("a59")
IE.Document.All("element-34-10835266-7896801-2").Value = sheet1.Range("b59")
'12 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896802-1").Value = sheet1.Range("a60")
IE.Document.All("element-34-10835266-7896802-2").Value = sheet1.Range("b60")
'13 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896803-1").Value = sheet1.Range("a61")
IE.Document.All("element-34-10835266-7896803-2").Value = sheet1.Range("b61")
'14 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896804-1").Value = sheet1.Range("a62")
IE.Document.All("element-34-10835266-7896804-2").Value = sheet1.Range("b62")
'15 Autor 1 QtdPubs 2
IE.Document.All("element-34-10835266-7896805-1").Value = sheet1.Range("a63")
IE.Document.All("element-34-10835266-7896805-2").Value = sheet1.Range("b63")

''12
'1 Pais 1 QtdPubs 2
IE.Document.All("element-35-10835267-7896806-1").Value = sheet1.Range("a66")
IE.Document.All("element-35-10835267-7896806-2").Value = sheet1.Range("b66")
'2 Pais 1 QtdPubs 2
IE.Document.All("element-35-10835267-7896807-1").Value = sheet1.Range("a67")
IE.Document.All("element-35-10835267-7896807-2").Value = sheet1.Range("b67")
'3 Pais 1 QtdPubs 2
IE.Document.All("element-35-10835267-7896808-1").Value = sheet1.Range("a68")
IE.Document.All("element-35-10835267-7896808-2").Value = sheet1.Range("b68")
'4 Pais 1 QtdPubs 2
IE.Document.All("element-35-10835267-7896809-1").Value = sheet1.Range("a69")
IE.Document.All("element-35-10835267-7896809-2").Value = sheet1.Range("b69")
'5 Pais 1 QtdPubs 2
IE.Document.All("element-35-10835267-7896810-1").Value = sheet1.Range("a70")
IE.Document.All("element-35-10835267-7896810-2").Value = sheet1.Range("b70")

''16
'1 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10835271-7896811-1").Value = sheet1.Range("a73")
IE.Document.All("element-35-10835271-7896811-2").Value = sheet1.Range("b73")
'2 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10835271-7896812-1").Value = sheet1.Range("a74")
IE.Document.All("element-35-10835271-7896812-2").Value = sheet1.Range("b74")
'3 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10835271-7896813-1").Value = sheet1.Range("a75")
IE.Document.All("element-35-10835271-7896813-2").Value = sheet1.Range("b75")
'4 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10835271-7896814-1").Value = sheet1.Range("a76")
IE.Document.All("element-35-10835271-7896814-2").Value = sheet1.Range("b76")
'5 NomePeriodico 1 QtdPubs 2
IE.Document.All("element-35-10835271-7896815-1").Value = sheet1.Range("a77")
IE.Document.All("element-35-10835271-7896815-2").Value = sheet1.Range("b77")

''18
'1 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10835273-7896816-1").Value = sheet1.Range("a80")
IE.Document.All("element-35-10835273-7896816-2").Value = sheet1.Range("b80")
'2 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10835273-7896817-1").Value = sheet1.Range("a81")
IE.Document.All("element-35-10835273-7896817-2").Value = sheet1.Range("b81")
'3 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10835273-7896818-1").Value = sheet1.Range("a82")
IE.Document.All("element-35-10835273-7896818-2").Value = sheet1.Range("b82")
'4 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10835273-7896819-1").Value = sheet1.Range("a83")
IE.Document.All("element-35-10835273-7896819-2").Value = sheet1.Range("b83")
'5 NomeAgencia 1 QtdProjsFinan 2
IE.Document.All("element-35-10835273-7896820-1").Value = sheet1.Range("a84")
IE.Document.All("element-35-10835273-7896820-2").Value = sheet1.Range("b84")

''19
'1 Congresso 1 Regs 2
IE.Document.All("element-35-10835275-7896821-1").Value = sheet1.Range("a87")
IE.Document.All("element-35-10835275-7896821-2").Value = sheet1.Range("b87")
'2 Congresso 1 Regs 2
IE.Document.All("element-35-10835275-7896822-1").Value = sheet1.Range("a88")
IE.Document.All("element-35-10835275-7896822-2").Value = sheet1.Range("b88")
'3 Congresso 1 Regs 2
IE.Document.All("element-35-10835275-7896823-1").Value = sheet1.Range("a89")
IE.Document.All("element-35-10835275-7896823-2").Value = sheet1.Range("b89")
'4 Congresso 1 Regs 2
IE.Document.All("element-35-10835275-7896824-1").Value = sheet1.Range("a90")
IE.Document.All("element-35-10835275-7896824-2").Value = sheet1.Range("b90")
'5 Congresso 1 Regs 2
IE.Document.All("element-35-10835275-7896825-1").Value = sheet1.Range("a91")
IE.Document.All("element-35-10835275-7896825-2").Value = sheet1.Range("b91")

''20
'1 Universidade 1 Regs 2
IE.Document.All("element-35-10835276-7896826-1").Value = sheet1.Range("a94")
IE.Document.All("element-35-10835276-7896826-2").Value = sheet1.Range("b94")
'2 Universidade 1 Regs 2
IE.Document.All("element-35-10835276-7896827-1").Value = sheet1.Range("a95")
IE.Document.All("element-35-10835276-7896827-2").Value = sheet1.Range("b95")
'3 Universidade 1 Regs 2
IE.Document.All("element-35-10835276-7896828-1").Value = sheet1.Range("a96")
IE.Document.All("element-35-10835276-7896828-2").Value = sheet1.Range("b96")
'4 Universidade 1 Regs 2
IE.Document.All("element-35-10835276-7896829-1").Value = sheet1.Range("a97")
IE.Document.All("element-35-10835276-7896829-2").Value = sheet1.Range("b97")
'5 Universidade 1 Regs 2
IE.Document.All("element-35-10835276-7896830-1").Value = sheet1.Range("a98")
IE.Document.All("element-35-10835276-7896830-2").Value = sheet1.Range("b98")

''22
'1 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896831-1").Value = sheet1.Range("a101")
IE.Document.All("element-35-10835278-7896831-2").Value = sheet1.Range("b101")
'2 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896832-1").Value = sheet1.Range("a102")
IE.Document.All("element-35-10835278-7896832-2").Value = sheet1.Range("b102")
'3 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896833-1").Value = sheet1.Range("a103")
IE.Document.All("element-35-10835278-7896833-2").Value = sheet1.Range("b103")
'4 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896834-1").Value = sheet1.Range("a104")
IE.Document.All("element-35-10835278-7896834-2").Value = sheet1.Range("b104")
'5 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896835-1").Value = sheet1.Range("a105")
IE.Document.All("element-35-10835278-7896835-2").Value = sheet1.Range("b105")
'6 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896836-1").Value = sheet1.Range("a106")
IE.Document.All("element-35-10835278-7896836-2").Value = sheet1.Range("b106")
'7 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896837-1").Value = sheet1.Range("a107")
IE.Document.All("element-35-10835278-7896837-2").Value = sheet1.Range("b107")
'8 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896838-1").Value = sheet1.Range("a108")
IE.Document.All("element-35-10835278-7896838-2").Value = sheet1.Range("b108")
'9 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896839-1").Value = sheet1.Range("a109")
IE.Document.All("element-35-10835278-7896839-2").Value = sheet1.Range("b109")
'10 PalavraChave 1 Regs 2
IE.Document.All("element-35-10835278-7896840-1").Value = sheet1.Range("a110")
IE.Document.All("element-35-10835278-7896840-2").Value = sheet1.Range("b110")

''Analise
IE.Document.All("element-13-10835261").Value = sheet1.Range("k9")
IE.Document.All("element-13-10835264").Value = sheet1.Range("k10")
IE.Document.All("element-13-10835265").Value = sheet1.Range("k11")
IE.Document.All("element-13-10835268").Value = sheet1.Range("k12")
IE.Document.All("element-13-10835269").Value = sheet1.Range("k13")
IE.Document.All("element-13-10835270").Value = sheet1.Range("k14")
IE.Document.All("element-13-10835272").Value = sheet1.Range("k15")
IE.Document.All("element-13-10835277").Value = sheet1.Range("k16")
IE.Document.All("element-13-10835279").Value = sheet1.Range("k17")
IE.Document.All("element-13-10835280").Value = sheet1.Range("k18")
IE.Document.All("element-13-10835281").Value = sheet1.Range("k19")

End Sub
