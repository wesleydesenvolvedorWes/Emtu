<%
Option Explicit
Response.Expires = 0
Dim objConn, objRs, strQuery, strConnection

'Conectando com o banco de dados contato.mdb
Set objConn =  Server.CreateObject("ADODB.Connection")
objConn.Open "DBQ=" & Server.MapPath("contato.mdb") & ";Driver={Microsoft Access Driver (*.mdb)}","username","password"

'Seleciona da tabela contato
strQuery = "SELECT * FROM contato"
Set ObjRs = objConn.Execute(strQuery)
%>
<html>
<LINK REL=stylesheet HREF="liks_etc.css" TYPE="text/css">
<head>
<title>Tela de Consulta - ASPBRASIL</title>
</head>
<body bgcolor="#FFFFFF">
<div align="center" style="width: 756; height: 119">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1" align="center">
    <tr>
      <td width="448" valign="top" height="136"> 
        <table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
          <tr bgcolor="ffffff"> 
            <td colspan="3">
              <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr>
                  <td bgcolor="#f5f5f5">
                    <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="form_inclusao.asp" class="menu">Incluir</a></font></b></font></div>
                  </td>
                  <td bgcolor="#f5f5f5">
                    <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CCCCCC"><a href="form_exclui.asp" class="menu"><font color="#000099">Excluir</font></a></font></b></font></div>
                  </td>
                  <td bgcolor="#f5f5f5">
                    <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CCCCCC">Alterar</font></b></font></div>
                  </td>
                </tr>
              </table>
              
            </td>
          </tr>
          <tr> 
            <td colspan="3"> 
              <table width="736" border="0" cellspacing="0" cellpadding="0" height="18">
                <tr> 
                  <td align="center" height="1" width="734"> <font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="3"><b><br>
                    Selecione o(s) &iacute;ten(s) a ser(em) excluido(s)</b></font></td>
                </tr>
              </table>
              <form method="GET" action="form_altera.asp">
                <table width="736" border="0" cellspacing="0" cellpadding="0" height="1">
                  <tr> 
                    <td height="1" align="center" width="734"> 
                      <table border="0" width="100%" height="63" cellpadding="2" align="center">
                        <tr bgcolor="#0099FF"> 
                          <td width="22" height="2" align="center"> <font size="1" color="FFFFFF"><b><font face="Verdana">N&ordm;:</font></b> 
                            </font></td>
                          <td width="151" height="2" align="center"> <font size="1" color="FFFFFF"><b><font face="Verdana">Nome:</font></b> 
                            </font></td>
                          <td colspan="2" height="2" align="center"> <font size="1" color="FFFFFF"><b></b> 
                            </font> <font size="1" color="FFFFFF"><b><font face="Verdana">Telefone:</font></b> 
                            </font></td>
                          <td width="162" height="2" align="center"> <font size="1" color="FFFFFF"><b><font face="Verdana">E-mail:</font></b> 
                            </font></td>
                          <td width="261" height="2" align="center"><font size="1" color="FFFFFF"><b><font face="Verdana">Comentário:</font></b> 
                            </font></td>
                          <td width="66" height="2" align="center"> <font size="1" color="FFFFFF"> 
                            <input type="submit" name="Submit" value="Alterar">
                            </font></td>
                        </tr>
                        <%While Not objRS.EOF %> 
                        <tr bgcolor="#FF9900"> 
                          <td width="22" height="2" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              <%Response.write objRS("autonum")%> </font> </b></font></div>
                          </td>
                          <td width="151" height="2" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              <%Response.write objRS("nome")%> </font> </b></font></div>
                          </td>
                          <td colspan="2" height="2" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              </font> </b></font> <font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif"><%Response.write objRS("telefone")%> 
                              </font></b></font></div>
                          </td>
                          <td width="162" height="2" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif"><a href="mailto:<%Response.write objRS("email")%>" class="menu"><%Response.write objRS("email")%></a> 
                              </font></b></font></div>
                          </td>
                          <td width="261" height="2" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%Response.write objRS("comentario")%></font></b></font></div>
                          </td>
                          <td width="66" height="2" align="center"> 
                            <div align="center"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              <input type="radio" name="radio" value="<%=objRS(0)%>">
                              </font> </b></font></div>
                          </td>
                        </tr>
                        <%
  'Move para o próximo registro
  objRS.MoveNext
  Wend
  'Fechando as conexões
  objRs.close
  objConn.close
  Set objRs = Nothing
  Set objConn = Nothing
  %> 
                      </table>
                    </td>
                  </tr>
                </table>
              </form>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</div>
</body>
</html>










