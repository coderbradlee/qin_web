<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<STYLE type=text/css>
BODY {
	BACKGROUND-COLOR: #F4FBFF;
	margin-left: 6px;
	margin-top: 6px;
	margin-right: 6px;
	margin-bottom: 6px;
	  }
</STYLE>
</head>

<body>
<% 
if trim(request.querystring("action"))="list" then
%>
<br>
<%
dim rs,sql
set rs=server.createobject("adodb.recordset")
	sql="select * from tuangou order by id desc"
	rs.open sql,conn,1,1
	rs.pagesize=5
	
	
if not rs.eof then
	
	
	if request.querystring("page")<>"" then
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
if rs.bof then
	response.write("<br><br><br><br><br><br>")
	response.write("<font color=red>暂无</font>留言！")
	response.write("<br><br><br><br><br><br>")
end if
for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input" style="margin-bottom:10px">
  <tr bgcolor="#DFEFFF">
    <td class="line enfont2"><b>发布人</b>：<%= rs("username") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;发布时间：<font style="font-size:11px"><%= rs("addtime") %></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;操作IP：<a href="http://tools.hxstat.com/ip/?ip=<%= rs("gip") %>" target="_blank"><%= rs("gip") %></a></td>
    <td align="right" bgcolor="#DFEFFF" class="line"><a href="?action=del&id=<%= rs("id") %>" onClick="return window.confirm('确定删除吗?');"><img src="images/cms-ico6.gif" width="8" height="7" border="0" align="absmiddle" title="点击删除"></a>&nbsp; </td>
  </tr>
  <tr>
    <td colspan="2" class="enfont2 content">
        <div><b>产品名称：</b><%=rs("pname")%></div>
        <div><b>手机号码：</b><%=rs("mobile")%></div>
      <div><b>固定电话：</b><%=rs("tel")%></div>
<div><b>联系地址：</b><%=rs("address")%></div>    
<div><b>购买时间：</b><%=rs("buytime")%></div>      
  
<div><b>说明：</b><%=rs("content")%></div>    </td>
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
end if
%>
<table width="100%" height="30" border="0" cellpadding="0" cellspacing="1" bordercolor="#cccccc">
  <tr>
    <td align="center" background="images/bg_title.gif">第<%= page %>页&nbsp;
        <% if page<>1 then %>
        <a href="?action=list&page=1">首页</a>
        <% else %>
      首页
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>">上一页</a>
      <% else %>
      上一页
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>">下一页</a>
      <% else %>
      下一页
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%=rs.pagecount%>">末页</a>
      <% else %>
      末页
      <% end if %>
      &nbsp;总数<%= rs.recordcount %>条</td>
    <td width="150" align="center" background="images/bg_title.gif">go
      <select name="select" onchange='javascript:window.open(this.options[this.selectedindex].value,"_self")'>
          <%for m = 1 to rs.pagecount%>
          <option value="?action=list&page=<%=m%>"><%=m%></option>
          <% next %>
        </select>
      页</td>
  </tr>
</table>
<% end if %>
<% 
if trim(request.querystring("action"))="del" then
		id=trim(request.querystring("id"))
		id=replacebadchar(id)
		set rs=server.createobject("adodb.recordset")
		sql="select * from tuangou where id="&id
		rs.open sql,conn,1,3
		rs.delete
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script>alert('删除成功!');location='?action=list'</script>"
	end if
 %>
<% if trim(request.querystring("action"))="check" then
if trim(request.form("submit"))="审 核" then
	id=trim(request.querystring("id"))
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from liuyan where id="&id
	rs.open sql,conn,1,3
	select case rs("check")
		case "未通过"
			rs("check")="通过"
		case "通过"
			rs("check")="未通过"
	end select
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script>alert('审核成功!');location='?action=list'</script>"
end if
 %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">确认<font color="#ff0000">审核留言</font>吗？<br>
<form name="form1" method="post" action="?action=check&id=<%= trim(request.querystring("id")) %>">
  <table width=200 border=0 cellpadding=0 
cellspacing=0 bordercolor=#9cacd0 class=table_out>
    <tr align="center">
      <td height=15><input type="submit" name="submit" value="审 核"></td>
      <td><input type="reset" name="submit2" value="取 消" onClick="javascript:history.go(-1)"></td>
    </tr>
  </table>
</form></td>
  </tr>
</table>
<% end if %>
<% 
if trim(request.querystring("action"))="reply" then
	if trim(request.form("submit"))="回 复" then
		dim reply
		content=trim(request.form("content"))
		id=trim(request.querystring("id"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from liuyan where id="&id
		rs.open sql,conn,1,3
		rs("reply")=content
		rs("rtime")=now()
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script>alert('回复成功!');location='?action=list'</script>"
		response.end
	end if
id=trim(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
sql="select * from liuyan where id="&id
rs.open sql,conn,1,1
 %>
<form name="form1" method="post" action="?action=reply&id=<%= trim(request.querystring("id")) %>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="input" style="margin-bottom:10px">
        <tr bgcolor="#DFEFFF"> 
          <td class="line enfont2"><b>留言人</b>：<%= rs("name") %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;留言时间：<font style="font-size:11px"><%= rs("time") %></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;操作IP：<a href="http://tools.hxstat.com/ip/?ip=<%=RS("GIP")%>" target="_blank"><%=RS("GIP")%></a></td>
        </tr>
        <tr> 
          <td class="enfont2 content">
		  <div><b>电话：</b><%= rs("tel") %></div><div><b>Email：</b><%= rs("email") %></div>
		  <div><b>留言内容：</b><%= rs("body") %></div>
		  </td>
        </tr>
      </table>
  
  
  
  
  
  
  
  
  
  
  
  
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
          <tr> 
            <td><b><font color="#FF3300">管理员回复↓</font></b> 
              <textarea name="content" class="input" id="content" style="height:100px"><%= rs("reply") %></textarea></td>
          </tr>
          <tr> 
            <td height="28"><input name="Submit" type="submit" class="bt" value="回 复"></td>
          </tr>
      </table>

  
  
  
  
  
  
  
  
  
  
  
  
</form>
<% end if %>
</body>
</html>
									  