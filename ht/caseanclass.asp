<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/body.css" rel="stylesheet" type="text/css">
</head>
<body>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th class="tableHeaderText" colspan=6 height=25>�������</th>	
  </tr>
  <tr>
    <td width="29%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">��������</div></td>
    <td width="34%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">Ӣ������</td>
    <td width="16%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">��������</div></td>
    <td width="21%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">ȷ������</div></td>
  </tr>
    <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from jd_caseclass order by flag asc ",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
	%>
     <form name="form1p" method="post" action="savecaseanclass.asp?action=edit&id=<%=int(rs("id"))%>">
    <tr>
      <td class="forumRowHighlight"><div align="center"><input name="anclass" type="text" id="anclass" size="12" value="<%=trim(rs("classname"))%>"></div></td>
	  <td align="center" class="forumRowHighlight"><input name="e_anclass" type="text" id="e_anclass" size="12" value="<%=trim(rs("e_classname"))%>"></td>
	  <td class="forumRowHighlight"><div align="center"><input name="anclassidorder" type="text" id="anclassidorder" size="4" value="<%=int(rs("flag"))%>"></div></td>
     <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit" value="��  ��">&nbsp; 
     
     <% if rs("id")>37 then %>
     <a href="savecaseanclass.asp?id=<%=int(rs("id"))%>&action=del" onClick="return confirm('�˲�����ɾ���˴����°�����С�������Ʒ����ȷ������ɾ��������')"><font color=red>ɾ��</font></a> 
     <% end if %>
     
     
     </div></td>
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
	<th class="tableHeaderText" colspan=6 height=25>�������</th>	
  </tr>
  <tr> 
    <td class="forumRowHighlight" colspan=4 ><div align="center">ע�⣺�������Ʋ��ܺ��зǷ��ַ�</div></td>
  </tr>
  <tr>
    <td width="29%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">��������</div></td>
    <td width="34%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">Ӣ������</td>
    <td width="16%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">��������</div></td>
   <td width="21%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">ȷ������</div></td>
 </tr>
 <form name="form1s" method="post" action="savecaseanclass.asp?action=add">
  <tr>
    <td class="forumRowHighlight"><div align="center"><input name="anclass2" type="text" id="anclass2" size="12"></div></td>
    <td align="center" class="forumRowHighlight"><input name="e_anclass2" type="text" id="e_anclass2" size="12"></td>
    <td class="forumRowHighlight"><div align="center"><input name="anclassidorder2" type="text" id="anclassidorder2" size="4" value="<%=paixu+1%>"></div></td>
    <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit3" value="�� ��"></div></td>
  </tr>
  </form>
</table>
</body>
</html>
