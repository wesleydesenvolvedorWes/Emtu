<%
'Forчa a declaraчуo de todas as variсveis
Option Explicit
'Nуo deixa informaчѕes no Cache
Response.Expires = 0
'Declaraчуo das variсveis
Dim objConn, objRs, strQuery, strConnection, nome, telefone, email, comentario
'Atrubuiчуo dos valores as respectivas variсveis
nome = Request.Form("nome")
telefone = Request.Form("telefone")
email = Request.Form("email")
comentario = Request.Form("comentario")
'Cria o objeto RecordSet e atribui a variсvel 
Set objConn =  Server.CreateObject("ADODB.Connection")
'Abre a conexуo com o banco de dados utilizando o Drive {Microsoft Access...
'(para utilizar outro, ex: Paradox щ sѓ substituir o Drive pelo do Paradox)
'(*.mdb) indica que o arquivo utiliza extensуo mdb
objConn.Open "DBQ=" & Server.MapPath("contato.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","username","password"
'Insere os dados na tabela aberta
strQuery = "INSERT INTO contato (nome,telefone,email,comentario) VALUES ('"&nome&"','"&telefone&"','"&email&"','"&comentario&"')"
'Caso ocorra um erro esta funчуo de erro serс chamada
On error Resume Next
'Executa a inserчуo no Banco de Dados 
Set ObjRs = objConn.Execute(strQuery)
'Fecha o Objeto de Conexуo
objConn.close
'"APAGA" qualquer instancia que possa ter no objeto objRs e objConn
Set objRs = Nothing
Set objConn = Nothing
'Caso a funчуo On Error Resume Next nуo tenha sido chamada o objeto err serс = a 0
if err = 0 Then
	'Redireciona o usuсrio caso nуo tenha ocorrido erro na transaчуo
	response.redirect "sucesso.asp"
end if
%>