<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>

<body>



      <% 
if trim(request.querystring("action"))="del" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Jiekuan where id="&id
	rs.open sql,conn,3,2
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('删除成功!');location='?page="&page&"&keywords="&keywords&"'</script>"
	response.end
end if
 %>


      <% 
if trim(request.querystring("action"))="rzyes" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Jiekuan where id="&id
	rs.open sql,conn,1,3
	rs("check")=true
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('审核成功!');location='?page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"'</script>"
	response.end
end if
 %>


      <% 
if trim(request.querystring("action"))="rzno" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Jiekuan where id="&id
	rs.open sql,conn,3,2
	rs("check")=false
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('已取消审核!');location='?page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"'</script>"
	response.end
end if
 %>

      <% 
if trim(request.querystring("action"))="tjno" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Jiekuan where id="&id
	rs.open sql,conn,3,2
	rs("jiekuan_tuijian")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('已取消推荐!');location='?page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"'</script>"
	response.end
end if
 %>

      <% 
if trim(request.querystring("action"))="tjyes" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Jiekuan where id="&id
	rs.open sql,conn,3,2
	rs("jiekuan_tuijian")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('推荐成功!');location='?page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"'</script>"
	response.end
end if
 %>





<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:7px">
  <form name="form" method="post" action="?Action=sousuo">
    <tr> 
      <td align="left">　信息搜索：
        <input name="keywords" type="text" class="input" id="keywords" style="width:150px;height:21px" onFocus='this.select()' onBlur="if (this.value ==''){this.value=this.defaultValue}" onClick="if(this.value=='输入信息关键词')this.value=''" value="输入信息关键词">
	  <input name="Submit" type="submit" class="bt" id="Submit" value="搜索">
      </td>
      <td align="right">&nbsp;</td>
    </tr>
  </form>
</table>










<table width="100%" border="0" cellspacing="0" cellpadding="8">
  <tr>
    <td valign="top"></td>
  </tr>
</table>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif" class="zw">
  <tr align="center">
    <td width="143" align="center" bgcolor="#DFEFFF"><b>编号</b></td>
    <td align="left" bgcolor="#DFEFFF"><b>信息名称</b></td>
    <td width="68" align="center" bgcolor="#DFEFFF">发布人</td>
    <td width="260" align="center" bgcolor="#DFEFFF">发布时间</td>
    <td width="50" align="center" bgcolor="#DFEFFF"><b>认证</b></td>
    <td width="50" align="center" bgcolor="#DFEFFF"><b>推荐</b></td>
    <td width="50" align="center" bgcolor="#DFEFFF"><b>删除</b></td>
    <td width="50" align="center" bgcolor="#DFEFFF"><b>编辑</b></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="8"></td>
  </tr>
</table>
<%
	user_ename=request.QueryString("user_ename")
	keywords=trim(request("keywords"))
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Jiekuan where 1=1 "
	
	if user_ename<>"" then
	sql=sql+" and user_name='"&user_ename&"' "
	end if
	
	if keywords<>"" then
	sql=sql+" and user_name like '%"&keywords&"%' or jiekuan_title like '%"&keywords&"%'"
	end if
	
	sql=sql+" order by id desc"
	
	rs.open sql,conn,1,1
	rs.pagesize=13
	
	
	if not rs.bof then

	if request.QueryString("page")<>"" then
	page=cint(trim(request.querystring("page")))
	else
	page=1
	end if
	if page<1 then
		page=1
	elseif page>rs.pagecount then
		page=rs.pagecount
	end if
	rs.absolutepage=page
		  %>
<%
 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" class="zw">
  <tr>
    <td width="141" align="center" class="line"><%=rs("jiekuan_bianhao")%></td>
    <td class="line">&nbsp;<a href="../Jiedai_Jiekuan_Show.asp?jid-jiedai=<%=rs("id")%>&uid=<%=left(rs("user_uid"),4)%>-<%=rs("user_uid")%>" class="zwe" target="_blank"><%= rs("Jiekuan_Title") %></a></td>
    <td width="69" align="center" class="line"><a href="?user_ename=<%=rs("user_name")%>">
    <%if rs("user_name")="jiedai_niming" then response.write"匿名" else response.write rs("user_name")%></a></td>
    <td width="259" align="center" class="line"><%=rs("addtime")%></td>
    <td width="50" align="center" class="line">
	
	
	<%if rs("check")=true then%>
	<a href="?action=rzno&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>"><img src="images/Unlock.gif" alt="已认证" width="14" height="13" border="0"></a>
	<%else%>
	<a href="?action=rzyes&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>"><img src="images/lock.gif" alt="未认证" width="11" height="12" border="0"></a>
	<%end if%>	</td>
    <td width="50" align="center" class="line">
	
	
	<%if rs("jiekuan_tuijian")=1 then%>
	<a href="?action=tjyes&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>"><img src="images/Ok.gif" alt="已推荐" width="16" height="16" border="0" /></a>
	<%else%>
	<a href="?action=tjno&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>"><img src="images/err.gif" alt="未推荐" width="12" height="11" border="0" /></a>
	<%end if%>	</td>
    <td width="50" align="center" class="line"><a href="?action=del&amp;jid=<%= rs("id") %>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>" onClick="return window.confirm('确定删除吗?');" style="font-size:12px; color:#666666"><img src="images/del.gif" alt="删除信息" width="16" height="16" border="0" /></a> </td>
    <td width="50" align="center" class="line"><a href="Edit_Jiekuan.asp?jid=<%= rs("id") %>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>" style="font-size:12px; color:#666666"><img src="images/Edit.gif" alt="修改信息" width="12" height="12" border="0" /></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="5"></td>
  </tr>
</table>
<% 
rs.movenext
next
else

response.write("<center>")
response.write("<font color=red>暂无</font>信息！")
response.write("<br><br></center>")
end if


%>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif" class="zw">
  <tr>
    <td align="center">第<%= page %>页&nbsp;
        <% if page<>1 and page<>"" then %>
        <a href="?action=list&amp;page=1&amp;classid=<%= classid %>&user_ename=<%=user_ename%>&keywords=<%=keywords%>" class="zw">首页</a>
        <% else %>
      首页
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&amp;page=<%= page-1 %>&amp;classid=<%= classid %>&user_ename=<%=user_ename%>&keywords=<%=keywords%>" class="zw">上一页</a>
      <% else %>
      上一页
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&amp;page=<%= page+1 %>&amp;classid=<%= classid %>&user_ename=<%=user_ename%>&keywords=<%=keywords%>" class="zw">下一页</a>
      <% else %>
      下一页
      <% end if %>
      &nbsp;
      <% if page<rs.recordcount then %>
      <a href="?action=list&amp;page=<%= rs.recordcount %>&amp;classid=<%= classid %>&user_ename=<%=user_ename%>&keywords=<%=keywords%>" class="zw">末页</a>
      <% else %>
      末页
      <% end if %>
      &nbsp;总数<%= rs.recordcount %>条</td>
  </tr>
</table>
</body>
</html>
