<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/body.css" rel="stylesheet" type="text/css">
</head>
<body>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th class="tableHeaderText" colspan=5 height=25><%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>�̼����͹���</th>	
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">��������</div></td>
    <td width="21%" class="forumRowHighlight"><div align="center">��������</div></td>
    <td width="31%" class="forumRowHighlight"><div align="center">ȷ������</div></td>
  </tr>
    <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from sjsort where sjid="&request.QueryString("sjid")&" order by anclassidorder asc ",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>��û�з���</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
	%>
     <form name="form1" method="post" action="savesjtype.asp?action=edit&id=<%=int(rs("anclassid"))%>">
    <tr>
      <td class="forumRowHighlight"><div align="center"><input name="anclass" type="text" id="anclass" size="16" value="<%=trim(rs("anclass"))%>" style="text-align:center; background:#FFCC99">
        <input name="sjid" type="hidden" id="sjid" value="<%=request.querystring("sjid")%>">
      </div></td>
	 <td class="forumRowHighlight"><div align="center"><input name="anclassidorder" type="text" id="anclassidorder" size="4" value="<%=int(rs("anclassidorder"))%>" style="text-align:center"></div></td>
     <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit" value="��  ��">&nbsp; <a href="savesjtype.asp?id=<%=int(rs("anclassid"))%>&action=del&sjid=<%=request.QueryString("sjid")%>" onClick="return confirm('�˲�����ɾ���˴����°�����С�������Ʒ����ȷ������ɾ��������')"><font color=red>ɾ��</font></a> </div></td>
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
	<th class="tableHeaderText" colspan=5 height=25>���<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>�̼�����</th>	
  </tr>
  <tr> 
    <td class="forumRowHighlight" colspan=3 ><div align="center">ע�⣺�������Ʋ��ܺ��зǷ��ַ�</div></td>
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">��������</div></td>
   <td width="21%" class="forumRowHighlight"><div align="center">��������</div></td>
   <td width="31%" class="forumRowHighlight"><div align="center">ȷ������</div></td>
 </tr>
 <form name="form1" method="post" action="savesjtype.asp?action=add">
  <tr>
    <td class="forumRowHighlight"><div align="center"><input name="anclass2" type="text" id="anclass2" size="12">
      <input name="sjid" type="hidden" id="sjid" value="<%=request.querystring("sjid")%>">
</div></td>
    <td class="forumRowHighlight"><div align="center"><input name="anclassidorder2" type="text" id="anclassidorder2" size="4" value="<%=paixu+1%>"></div></td>
    <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit3" value="�� ��"></div></td>
  </tr>
  </form>
</table>
</body>
</html>
