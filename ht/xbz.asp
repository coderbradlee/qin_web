<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/body.css" rel="stylesheet" type="text/css">
</head>
<body>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th class="tableHeaderText" colspan=5 height=25>���͹���</th>	
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">��������</div></td>
    <td width="21%" class="forumRowHighlight"><div align="center">��������</div></td>
    <td width="31%" class="forumRowHighlight"><div align="center">ȷ������</div></td>
  </tr>
    <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from xbz order by anclassidorder ",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
	%>
     <form name="form1" method="post" action="sbzsave.asp?action=edit&id=<%=int(rs("anclassid"))%>">
    <tr>
      <td class="forumRowHighlight"><div align="center"><input name="anclass" type="text" id="anclass" size="12" value="<%=trim(rs("anclass"))%>"></div></td>
	 <td class="forumRowHighlight"><div align="center"><input name="anclassidorder" type="text" id="anclassidorder" size="4" value="<%=int(rs("anclassidorder"))%>"></div></td>
     <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit" value="��  ��">&nbsp; <a href="sbzsave.asp?id=<%=int(rs("anclassid"))%>&action=del" onClick="return confirm('�˲�����ɾ���˴����°�����С�������Ʒ����ȷ������ɾ��������')"><font color=red>ɾ��</font></a> </div></td>
   </tr>
   </form>
   <%rs.MoveNext
          loop
          paixu=rs.RecordCount
          end if%>
</table>
<br>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th class="tableHeaderText" colspan=5 height=25>�������</th>	
  </tr>
  <tr> 
    <td class="forumRowHighlight" colspan=3 ><div align="center">ע�⣺�������Ʋ��ܺ��зǷ��ַ�</div></td>
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">��������</div></td>
   <td width="21%" class="forumRowHighlight"><div align="center">��������</div></td>
   <td width="31%" class="forumRowHighlight"><div align="center">ȷ������</div></td>
 </tr>
 <form name="form1" method="post" action="sbzsave.asp?action=add">
  <tr>
    <td class="forumRowHighlight"><div align="center"><input name="anclass2" type="text" id="anclass2" size="12"></div></td>
    <td class="forumRowHighlight"><div align="center"><input name="anclassidorder2" type="text" id="anclassidorder2" size="4" value="<%=paixu+1%>"></div></td>
    <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit3" value="�� ��"></div></td>
  </tr>
  </form>
</table>
</body>
</html>
