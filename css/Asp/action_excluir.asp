<%
Option Explicit
Response.Expires = 0
Dim objConn, stringSQL, strConnection, array_id, i, sql_id, id
id = Request.QueryString("checkbox")
'Caso ocorra algum erro os precessos não são interrompidos 
'e é passado para a próxima linha de comando
On error Resume Next
' Conectando com o banco de dados contato.mdb
Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "DBQ=" & Server.MapPath("contato.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","username","password"
'Deletando registro da tabela contato onde esta a id
	%>
<html>
<head>
<LINK REL=stylesheet HREF="liks_etc.css" TYPE="text/css">
<title>Tela de Exclusão - ::ASPBRASIL::</title>
</head>
<body bgcolor="#FFFFFF">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr> 
		<td>
<% 
		if err = 0 and id <> "" then
			array_id = split(id,",")
			For i=0 to ubound(array_id)
				sql_id = sql_id & "contato.autonum = " & Trim(array_id(i)) & " OR "
														 'campo texto, entao" & Trim(array_id(i)) & " OR "
														 'caso numerico '" & Trim(array_id(i)) & "' OR "
			Next
			sql_id = left(sql_id,(len(sql_id)-4))
			stringSQL = "DELETE * FROM contato WHERE "&sql_id&""
			objConn.Execute(stringSQL)
			objConn.close
			Set objConn = Nothing
		  
%>		  	
				  <table width="100%" border="0" cellspacing="2" cellpadding="2">
					<tr align="center"> 
					 <td bgcolor="#f5f5f5" width="30%"> 
						<div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="form_inclusao.asp" class="menu">Incluir</a></font></b></font></div>
					  </td>
					  <td bgcolor="#f5f5f5" width="30%"> 
						<div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="form_exclui.asp" class="menu">Excluir</a></font></b></font></div>
					  </td>
					  <td bgcolor="#f5f5f5" width="35%"> 
						<div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CCCCCC"><a href="escolhe_pra_auterar.asp" class="menu">Alterar</a></font></b></font></div>
					  </td>
					</tr>
				  </table>
			<table border="0" width="100%" height="8" cellpadding="2" align="center">
			  <tr bgcolor="#0099FF"> 
				<td colspan="7" height="1" align="center"> <font size="2" color="FFFFFF"><b><font face="Verdana, Arial, Helvetica, sans-serif">Seus 
				  dados foram excluidos com sucesso!</font></b></font> </td>
			  </tr>
			</table>
		<%else%>
			  <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Ocorreu algum erro!<br>Nenhum dado foi excluido!</b><br><a href="javascript:history.back(-1)">Volta</a></font></div>
		<%End if%>
		</td>
	  </tr>
	</table>
	</body>
	</html>






