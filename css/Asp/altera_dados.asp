<%
'Forчa o programador a declarar todas as variсveis, evitando erro de digitaчуo no uso das variщveis
Option Explicit

'Nуo deixa informaчѕes no Cache
Response.Expires = 0

'Declaraчуo das variсveis
Dim objConn, strQuery, sql_query, RsQuery, campo, sql, autonum
Dim nome, telefone, email, comentario, ObjRs

'Atrubuiчуo dos valores as respectivas variсveis
nome = Request.Form("nome")
telefone = Request.Form("telefone")
email = Request.Form("email")
comentario = Request.Form("comentario")
if comentario = "" then
	comentario = " "
end if
autonum = Request.Form("autonum")

'Cria o objeto RecordSet e atribui a variсvel 
Set objConn =  Server.CreateObject("ADODB.Connection")
'Abre a conexуo com o banco de dados utilizando o Drive {Microsoft Access...
'(para utilizar outro, ex: Paradox щ sѓ substituir o Drive pelo do Paradox)
'(*.mdb) indica que o arquivo utiliza extensуo mdb
objConn.Open "DBQ=" & Server.MapPath("contato.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","username","password"

strQuery = "UPDATE contato SET nome = '"&nome&"', telefone='"&telefone&"', email='"&email&"', comentario='"&comentario&"' WHERE autonum ="&autonum

'Caso ocorra um erro esta funчуo de erro serс chamada
On error Resume Next
'Executaa inserчуo no Banco de Dados 
Set ObjRs = objConn.Execute(strQuery)
'Fecha o Objeto de Conexуo
objConn.close
'"APAGA" qualquer instancia que possa ter no objeto objRs e objConn
Set objRs = Nothing
Set objConn = Nothing 
if err = 0 Then
	'Redireciona o usuсrio caso nуo tenha ocorrido erro na transaчуo
	response.redirect "sucesso.asp"
end if
%>