<%
'For�a o programador a declarar todas as vari�veis, evitando erro de digita��o no uso das vari�veis
Option Explicit

'N�o deixa informa��es no Cache
Response.Expires = 0

'Declara��o das vari�veis
Dim objConn, strQuery, sql_query, RsQuery, campo, sql, autonum
Dim nome, telefone, email, comentario, ObjRs

'Atrubui��o dos valores as respectivas vari�veis
nome = Request.Form("nome")
telefone = Request.Form("telefone")
email = Request.Form("email")
comentario = Request.Form("comentario")
if comentario = "" then
	comentario = " "
end if
autonum = Request.Form("autonum")

'Cria o objeto RecordSet e atribui a vari�vel 
Set objConn =  Server.CreateObject("ADODB.Connection")
'Abre a conex�o com o banco de dados utilizando o Drive {Microsoft Access...
'(para utilizar outro, ex: Paradox � s� substituir o Drive pelo do Paradox)
'(*.mdb) indica que o arquivo utiliza extens�o mdb
objConn.Open "DBQ=" & Server.MapPath("contato.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","username","password"

strQuery = "UPDATE contato SET nome = '"&nome&"', telefone='"&telefone&"', email='"&email&"', comentario='"&comentario&"' WHERE autonum ="&autonum

'Caso ocorra um erro esta fun��o de erro ser� chamada
On error Resume Next
'Executaa inser��o no Banco de Dados 
Set ObjRs = objConn.Execute(strQuery)
'Fecha o Objeto de Conex�o
objConn.close
'"APAGA" qualquer instancia que possa ter no objeto objRs e objConn
Set objRs = Nothing
Set objConn = Nothing 
if err = 0 Then
	'Redireciona o usu�rio caso n�o tenha ocorrido erro na transa��o
	response.redirect "sucesso.asp"
end if
%>