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
IE.Document.All("element-35-10553764-7707338-1").Value = ThisWorkbook.Sheets("Planilha1").Range("a1")


End Sub
