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
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>商家类型管理</th>	
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">类型名称</div></td>
    <td width="21%" class="forumRowHighlight"><div align="center">类型排序</div></td>
    <td width="31%" class="forumRowHighlight"><div align="center">确定操作</div></td>
  </tr>
    <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from sjsort where sjid="&request.QueryString("sjid")&" order by anclassidorder asc ",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>还没有分类</font></center>"
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
     <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit" value="修  改">&nbsp; <a href="savesjtype.asp?id=<%=int(rs("anclassid"))%>&action=del&sjid=<%=request.QueryString("sjid")%>" onClick="return confirm('此操作会删除此大类下包含的小分类和商品！您确定进行删除操作吗？')"><font color=red>删除</font></a> </div></td>
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
	<th class="tableHeaderText" colspan=5 height=25>添加<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>商家类型</th>	
  </tr>
  <tr> 
    <td class="forumRowHighlight" colspan=3 ><div align="center">注意：各项名称不能含有非法字符</div></td>
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">类型名称</div></td>
   <td width="21%" class="forumRowHighlight"><div align="center">类型排序</div></td>
   <td width="31%" class="forumRowHighlight"><div align="center">确定操作</div></td>
 </tr>
 <form name="form1" method="post" action="savesjtype.asp?action=add">
  <tr>
    <td class="forumRowHighlight"><div align="center"><input name="anclass2" type="text" id="anclass2" size="12">
      <input name="sjid" type="hidden" id="sjid" value="<%=request.querystring("sjid")%>">
</div></td>
    <td class="forumRowHighlight"><div align="center"><input name="anclassidorder2" type="text" id="anclassidorder2" size="4" value="<%=paixu+1%>"></div></td>
    <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit3" value="添 加"></div></td>
  </tr>
  </form>
</table>
</body>
</html>
