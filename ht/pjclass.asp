<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/body.css" rel="stylesheet" type="text/css">
</head>
<body>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th class="tableHeaderText" colspan=6 height=25>大类别管理</th>	
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">分类名称</div></td>
    <td width="33%" align="center" class="forumRowHighlight">图片地址</td>
    <td width="21%" align="center" class="forumRowHighlight"><div align="center">分类排序</div></td>
    <td width="31%" class="forumRowHighlight"><div align="center">确定操作</div></td>
  </tr>
    <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from peijian_class order by anclassidorder ",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>还没有分类</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
	%>
     <form name="form1" method="post" action="savepjanclass.asp?action=edit&id=<%=int(rs("anclassid"))%>">
    <tr>
      <td class="forumRowHighlight"><div align="center"><input name="anclass" type="text" id="anclass" size="12" value="<%=trim(rs("anclass"))%>"></div></td>
	  <td align="center" class="forumRowHighlight"><input type="text" name="tupian" value="<%=rs("tupian")%>"></td>
	  <td class="forumRowHighlight"><div align="center"><input name="anclassidorder" type="text" id="anclassidorder" size="4" value="<%=int(rs("anclassidorder"))%>"></div></td>
     <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit" value="修  改">&nbsp; <a href="savepjanclass.asp?id=<%=int(rs("anclassid"))%>&action=del" onClick="return confirm('此操作会删除此大类下包含的小分类和商品！您确定进行删除操作吗？')"><font color=red>删除</font></a> </div></td>
   </tr>
    <tr>
      <td colspan="3" align="right" class="forumRowHighlight"><textarea name="jianjie" cols="80" rows="6" id="jianjie"><%=rs("jianjie")%></textarea></td>
      <td class="forumRowHighlight">&nbsp;</td>
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
	<th class="tableHeaderText" colspan=5 height=25>添加大类别</th>	
  </tr>
  <tr> 
    <td class="forumRowHighlight" colspan=3 ><div align="center">注意：各项名称不能含有非法字符</div></td>
  </tr>
  <tr>
    <td width="33%" class="forumRowHighlight"><div align="center">分类名称</div></td>
   <td width="21%" class="forumRowHighlight"><div align="center">分类排序</div></td>
   <td width="31%" class="forumRowHighlight"><div align="center">确定操作</div></td>
 </tr>
 <form name="form1" method="post" action="savepjanclass.asp?action=add">
  <tr>
    <td class="forumRowHighlight"><div align="center"><input name="anclass2" type="text" id="anclass2" size="12"></div></td>
    <td class="forumRowHighlight"><div align="center"><input name="anclassidorder2" type="text" id="anclassidorder2" size="4" value="<%=paixu+1%>"></div></td>
    <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit3" value="添 加"></div></td>
  </tr>
  <tr>
    <td colspan="2" align="right" class="forumRowHighlight"><textarea name="jianjie" cols="80" rows="6" id="jianjie"></textarea></td>
    <td class="forumRowHighlight">&nbsp;</td>
    </tr>
  </form>
</table>
</body>
</html>
