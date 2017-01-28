<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body>

<% 
if trim(request.querystring("action"))="list" then
classid=trim(request.querystring("classid"))
set rs=server.createobject("adodb.recordset")
sql="select * from jiedai_danye order by id desc"
rs.open sql,conn,1,1
rs.pagesize=10
page=cint(trim(request.querystring("page")))
if page<1 then
    page=1
elseif page>rs.pagecount then
page=rs.pagecount
end if
rs.absolutepage=page
 %>
<table width="600" height="25" border="1" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
  <tr> 
    <td align="center"> <strong><font color="#215dc6">类型管理</font></strong> </td>
  </tr>
</table><br>
<%
if rs.bof then response.write("<center><br><br><br><br><br><br><font color=red>暂无</font>信息！<br><br><br><br><br><br></center>")
 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
<table width="600" height="25" border="1" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr align="center">
    <td width="50" align="center"><%= rs("id") %></td>
    <td align="left"><span class="style1">&nbsp;<%= rs("classid") %></span></td>
    <td width="50" align="center"><a href="?action=edit&id=<%= rs("id") %>">[修改]</a></td>
  </tr>
</table>
<br>
<% 
rs.movenext
next
%>
<table width="600" height="25" border="1" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
  <tr> 
    <td align="center">第<%= page %>页&nbsp; 
      <% if page<>1 then %>
      <a href="?action=list&page=1&classid=<%= classid %>">首页</a> 
      <% else %>
      首页 
      <% end if %>
      &nbsp; 
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>&classid=<%= classid %>">上一页</a> 
      <% else %>
      上一页 
      <% end if %>
      &nbsp; 
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>&classid=<%= classid %>">下一页</a> 
      <% else %>
      下一页 
      <% end if %>
      &nbsp; 
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%=rs.pagecount%>&classid=<%= classid %>">末页</a> 
      <% else %>
      末页 
      <% end if %>
      &nbsp;总数<%= rs.recordcount %>条</td>
    <td width="100" align="center">go 
      <select name="select" onchange='javascript:window.open(this.options[this.selectedindex].value,"_self")'>
        <%for m = 1 to rs.pagecount%>
        <option value="?action=list&page=<%=m%>&classid=<%= classid %>"><%=m%></option>
        <% next %>
      </select>
    页</td>
  </tr>
</table>
<% end if %>


<% if trim(request.querystring("action"))="add" then
	if trim(request.form("submit"))="添加" then
		classid=trim(request.form("classid"))
		for i = 1 to request.form("content1").count
		  scontent = scontent & request.form("content1")(i)
		next
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_danye"
		rs.open sql,conn,1,3
		rs.addnew
		rs("classid")=classid
		rs("body")=scontent
		rs.update
		rs.requery
		response.write("<script>alert('添加成功');location='?action=list';</script>")
	end if
%>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_add(){
	var notnull;
	notnull=true;
	if (document.form1.classid.value==""){
		alert("标题不能为空！");
		document.form1.classid.focus();
		notnull=false;
		}
		return notnull;
	}
</script>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<br>
<br>
<form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
  <table width="600" height="92" border="1" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr> 
      <td width="50" height="30" align="center">标题:</td>
      <td>&nbsp;
      <input name="classid" type="text" id="classid" size="40"></td></tr>
    <tr>
      <td height="30" align="center">
	  内容:</td>
      <td>
	  <input type="hidden" name="content1" value="">
	   <iframe id="ewebeditor1" src="../ewebeditor/ewebeditor.asp?id=content1&style=standard" frameborder="0" scrolling="no" width="550" height="350"></iframe>	
	  </td>
    </tr>
    <tr>
      <td height="30" colspan="2" align="center" background="images/bg_title.gif"><input type="submit" name="submit" value="添加">
&nbsp;
<input type="reset" name="submit4" value="重置"></td>
    </tr>
  </table>
</form>
<% end if %>




<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td align="center" valign="top">
<% if trim(request.querystring("action"))="edit" then
	if trim(request.form("add"))="add" then
		id=trim(request.querystring("id"))
		classid=trim(request.form("classid"))
		for i = 1 to request.form("content1").count
		  scontent = scontent & request.form("content1")(i)
		next
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_danye where id="&id
		rs.open sql,conn,1,3
'		rs("classid")=classid
		rs("body")=scontent
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		
		
		
	end if
	
		id=trim(request.querystring("id"))
		sql="select * from jiedai_danye where id="&id
		set rs=conn.execute(sql)

%>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.classid.value==""){
		alert("标题不能为空！");
		document.form1.classid.focus();
		notnull=false;
		}
		return notnull;
	}
</script>
<form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit()">
  <table width="100%" height="417" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr>
      <td height="20" bgcolor="#D3E5FA" style="padding-left:15"><b><%= rs("classid") %></b>&nbsp;
        <input name="add" type="hidden" id="add" value="add"></td>
      </tr>
    <tr>
      <td height="323" align="center" valign="top">
      
      	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
      	  <input type="hidden" name="content1" value="<%=rs("body")%>">
      <iframe id="ewebeditor1" src="../jiedaied/jiedaied.asp?id=content1&style=jiedaiedit" frameborder="0" scrolling="no" width="100%" height="340"></iframe></td>
      </tr>
    <tr>
      <td height="30" align="left" valign="top" background="images/bg_title.gif" style="padding-left:50">        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table>
        <%if request.form("add")="add" then
		 response.write"<img src=images/cms-ico7.gif width=12 height=11><font color=#ff0000><b>"&rs("classid")&"-</b>信息已修改成功</font>"
		 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table><%end if%>
        <input type="image" name="imageField" id="imageField" src="images/submit-bt.gif">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5"></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5"></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5"></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
<% end if %>



    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    </td>
  </tr>
</table>

<% if trim(request.querystring("action"))="del" then %>
<% 
if trim(request.form("submit"))="确 认" then
	id=trim(request.querystring("id"))
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_danye where id="&id
	rs.open sql,conn,2,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write("<script>alert('删除成功');location='?action=list';</script>")
end if
 %>
<br><br><br><br><br><br><br><br>
确认删除<font color="#ff0000">新闻类型</font>吗？<br>
<form name="form1" method="post" action="?action=del&id=<%= trim(request.querystring("id")) %>">
  <table width=200 border=1 cellpadding=0 
cellspacing=0 bordercolor=#9cacd0 class=table_out>
    <tr align="center"> 
      <td height=15><input type="submit" name="submit" value="确 认"></td>
      <td><input type="reset" name="submit2" value="取 消" onClick="javascript:history.go(-1)"></td>
    </tr>
  </table>
</form>
<% end if %>
</body>
</html>                                                                             