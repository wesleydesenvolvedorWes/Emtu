<html>
<LINK REL=stylesheet HREF="liks_etc.css" TYPE="text/css">
<head>
<title>ASPBRASIL</title>
<script language="javascript">
function valida_campo()
{
<!--
var nome = document.form.nome.value
if (nome==""){
	alert("Entre com seu nome!");
	document.form.nome.focus()
	return false
	}
var telefone = document.form.telefone.value
if (telefone==""){
	alert("Entre com seu telefone!");
	document.form.telefone.focus()
	return false
	}
var email=document.form.email.value;
if (email==""){
	alert("Entre com seu email!")
	document.form.email.focus()
return false
	}
}
function confere(){
if (document.form.email.value.indexOf('@', 0) == -1 || document.form.email.value.indexOf('.', 0) == -1){ alert("E-mail invalido!");
	document.form.email.focus()
	}
}
//-->
</script>
</head>
<body>
<form method="post" action="insert_into.asp" name="form" onsubmit="return valida_campo()">
  <div align="center">
    <center>
      <table width="44%" border="0" cellspacing="2" cellpadding="2">
        <tr align="center"> 
          <td bgcolor="#f5f5f5" width="35%"> 
            <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><font color="#CCCCCC">Incluir</font></font></b></font></div>
          </td>
          <td bgcolor="#f5f5f5" width="30%"> 
            <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="form_exclui.asp" class="menu">Excluir</a></font></b></font></div>
          </td>
          <td bgcolor="#f5f5f5" width="35%"> 
            <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CCCCCC"><a href="escolhe_pra_auterar.asp" class="menu">Alterar</a></font></b></font></div>
          </td>
        </tr>
      </table>
      <table border="0" width="300" bgcolor="F5F5F5">
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2" height="34"> 
            <div align="center"><font size="2"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000066">Cadastro 
              ASPBRASIL.</font></b></font></div>
          </td>
        </tr>
        <tr> 
          <td width="76"><font size="2" face="Verdana" color="0000cc">Nome:</font></td>
          <td width="210"> 
            <input type="text" name="nome" size="20" style="font-family: Verdana; font-size: 8 pt;COLOR: 0000CC;">
          </td>
        </tr>
        <tr> 
          <td width="76"><font size="2" face="Verdana" color="0000cc">E-mail:</font></td>
          <td width="210"> 
            <input type="text" name="email" size="20" style="font-family: Verdana; font-size: 8 pt; COLOR: 0000CC;" onBlur="confere()">
          </td>
        </tr>
        <tr> 
          <td width="76"><font size="2" face="Verdana" color="0000cc">Telefone:</font></td>
          <td width="210"> 
            <input type="text" name="telefone" size="20" style="font-family: Verdana; font-size: 8 pt;COLOR: 0000CC;">
          </td>
        </tr>
        <tr> 
          <td width="76"><font size="2" face="Verdana" color="0000cc">Comentário:</font></td>
          <td width="210"> 
            <textarea rows="4" name="comentario" cols="20" style="font-family: Verdana; font-size: 8 pt;COLOR: 0000CC;"></textarea>
          </td>
        </tr>
        <tr> 
          <td width="286" colspan="2"> 
            <p align="center"> 
              <input type="submit" value="Enviar" name="enviar">
          </td>
        </tr>
      </table>
    </center>
  </div>
</form>
</body>
</html>
