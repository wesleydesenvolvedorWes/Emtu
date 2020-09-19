<%
Option Explicit
Response.Expires = 0
Dim objConn, objRs, strQuery
Dim strConnection

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
<title>Visão Geral - ::ASPBRASIL::</title>
</head>
<body bgcolor="#FFFFFF">
<div align="center" style="width: 756; height: 119">
  <table width="710" border="0" cellspacing="0" cellpadding="0" height="1">
    <tr>
      <td width="448" valign="top" height="136"> 
        <table border="0" cellpadding="0" cellspacing="0" width="740">
          <tr bgcolor="ffffff"> 
            <td colspan="3">
              <table width="100%" border="0" cellspacing="2" cellpadding="2">
                <tr>
                  <td bgcolor="#f5f5f5">
                    <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="form_inclusao.asp" class="menu">Incluir</a></font></b></font></div>
                  </td>
                  <td bgcolor="#f5f5f5">
                    <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CCCCCC">Excluir</font></b></font></div>
                  </td>
                  <td bgcolor="#f5f5f5">
                    <div align="center"><font color="#000099"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CCCCCC"><a href="escolhe_pra_auterar.asp" class="menu">Alterar</a></font></b></font></div>
                  </td>
                </tr>
              </table>
              
            </td>
          </tr>
          <tr> 
            <td colspan="3"> 
              <table width="736" border="0" cellspacing="0" cellpadding="0" height="18">
                <tr> 
                  <td align="center" height="1" width="734"> <font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="3"><b>Vis&atilde;o 
                    Geral </b></font></td>
                </tr>
              </table>
              <form method="GET" action="action_excluir.asp">
                <table width="736" border="0" cellspacing="0" cellpadding="0" height="1">
                  <tr> 
                    <td height="1" align="center" width="734"> 
                      <table border="0" width="740" height="63" cellpadding="2">
                        <tr bgcolor="#0099FF"> 
                          <td width="22" height="1" align="center"> <font size="1" color="FFFFFF"><b><font face="Verdana">N&ordm;:</font></b> 
                            </font></td>
                          <td width="151" height="1" align="center"> <font size="1" color="FFFFFF"><b><font face="Verdana">Nome:</font></b> 
                            </font></td>
                          <td colspan="2" height="1" align="center"> <font size="1" color="FFFFFF"><b></b> 
                            </font> <font size="1" color="FFFFFF"><b><font face="Verdana">Telefone:</font></b> 
                            </font></td>
                          <td width="162" height="1" align="center"> <font size="1" color="FFFFFF"><b><font face="Verdana">E-mail:</font></b> 
                            </font></td>
                          <td width="261" height="1" align="center"><font size="1" color="FFFFFF"><b><font face="Verdana">Comentário:</font></b> 
                            </font></td>
                          <td width="66" height="1" align="center"> <font size="1" color="FFFFFF"> 
                            <input type="submit" name="Submit" value="Excluir">
                            </font></td>
                        </tr>
                        <%While Not objRS.EOF %> 
                        <tr bgcolor="#FF9900"> 
                          <td width="22" height="1" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              <%Response.write objRS("autonum")%> </font> </b></font></div>
                          </td>
                          <td width="151" height="1" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              <%Response.write objRS("nome")%> </font> </b></font></div>
                          </td>
                          <td colspan="2" height="1" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              </font> </b></font> <font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif"><%Response.write objRS("telefone")%> 
                              </font></b></font></div>
                          </td>
                          <td width="162" height="1" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif"><a href="mailto:<%Response.write objRS("email")%>" class="menu"><%Response.write objRS("email")%></a> 
                              </font></b></font></div>
                          </td>
                          <td width="261" height="1" align="center"> 
                            <div align="left"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%Response.write objRS("comentario")%></font></b></font></div>
                          </td>
                          <td width="66" height="1" align="center"> 
                            <div align="center"><font size="1"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000"> 
                              <input type="checkbox" name="checkbox" value="<%=objRS(0)%>">
                              </font> </b></font></div>
                          </td>
                        </tr>
                        <%
  'Movendo para o proximo registro
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










