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
	sql="select * from jiedai_Wenda order by id desc"
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
	response.write("û����Ϣ<font color=red></font>��")
	response.write("<br><br><br><br><br><br>")
end if
for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input" style="margin-bottom:10px">
  <tr bgcolor="#DFEFFF">
    <td class="line enfont2"><b>����</b>��<%= rs("uname") %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���ʱ�䣺<font style="font-size:11px"><%= rs("addtime") %></font>&nbsp;&nbsp;&nbsp;</td>
    <td align="right" bgcolor="#DFEFFF" class="line"><a href="?action=reply&id=<%= rs("id") %>">�ظ�</a>  <a href="?action=del&id=<%= rs("id") %>" onClick="return window.confirm('ȷ��ɾ����?');"><img src="images/cms-ico6.gif" width="8" height="7" border="0" align="absmiddle" title="���ɾ��"></a>&nbsp; </td>
  </tr>
  <tr>
    <td colspan="2" class="enfont2 content"><div><b>���⣺</b><%=rs("utitle")%>&nbsp;</div>
        <div><b>�绰��</b><%=rs("utel")%></div>
      <div><b>Email��</b><%=rs("uemailo")%></div>
      <div><b>���ݣ�</b><%=rs("ucontent")%></div>
      <table width="100%" border="0" cellspacing="0" cellpadding="3" style="border: 1px solid #D8CA9A;margin-top:5px">
          <tr>
            <td colspan="4" bgcolor="#FFFCF0" class="content enfont2"><b>�ظ���</b><font color="#FF0000"><%=rs("reply")%></font><br><%=rs("rtime")%></td>
          </tr>
      </table></td>
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
    <td align="center" background="images/bg_title.gif">��<%= page %>ҳ&nbsp;
        <% if page<>1 then %>
        <a href="?action=list&page=1">��ҳ</a>
        <% else %>
      ��ҳ
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%=rs.pagecount%>">ĩҳ</a>
      <% else %>
      ĩҳ
      <% end if %>
      &nbsp;����<%= rs.recordcount %>��</td>
    <td width="150" align="center" background="images/bg_title.gif">go
      <select name="select" onchange='javascript:window.open(this.options[this.selectedindex].value,"_self")'>
          <%for m = 1 to rs.pagecount%>
          <option value="?action=list&page=<%=m%>"><%=m%></option>
          <% next %>
        </select>
      ҳ</td>
  </tr>
</table>
<% end if %>
<% 
if trim(request.querystring("action"))="del" then
		id=trim(request.querystring("id"))
		id=replacebadchar(id)
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_Wenda where id="&id
		rs.open sql,conn,1,3
		rs.delete
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script>alert('ɾ���ɹ�!');location='?action=list'</script>"
	end if
 %>
<% if trim(request.querystring("action"))="check" then
if trim(request.form("submit"))="�� ��" then
	id=trim(request.querystring("id"))
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Wenda where id="&id
	rs.open sql,conn,1,3
	select case rs("check")
		case "δͨ��"
			rs("check")="ͨ��"
		case "ͨ��"
			rs("check")="δͨ��"
	end select
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script>alert('��˳ɹ�!');location='?action=list'</script>"
end if
 %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center">ȷ��<font color="#ff0000">�������</font>��<br>
<form name="form1" method="post" action="?action=check&id=<%= trim(request.querystring("id")) %>">
  <table width=200 border=0 cellpadding=0 
cellspacing=0 bordercolor=#9cacd0 class=table_out>
    <tr align="center">
      <td height=15><input type="submit" name="submit" value="�� ��"></td>
      <td><input type="reset" name="submit2" value="ȡ ��" onClick="javascript:history.go(-1)"></td>
    </tr>
  </table>
</form></td>
  </tr>
</table>
<% end if %>
<% 
if trim(request.querystring("action"))="reply" then
	if trim(request.form("submit"))="�� ��" then
		dim reply
		content=trim(request.form("content"))
		id=trim(request.querystring("id"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_Wenda where id="&id
		rs.open sql,conn,1,3
		rs("reply")=content
		rs("rtime")=now()
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script>alert('�ظ��ɹ�!');location='?action=list'</script>"
		response.end
	end if
id=trim(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
sql="select * from jiedai_Wenda where id="&id
rs.open sql,conn,1,1
 %>
<form name="form1" method="post" action="?action=reply&id=<%= trim(request.querystring("id")) %>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="input" style="margin-bottom:10px">
        <tr bgcolor="#DFEFFF"> 
          <td class="line enfont2"><strong>����</strong>��<%= rs("uname") %> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ʱ�䣺<font style="font-size:11px"><%= rs("addtime") %></font>&nbsp;&nbsp;&nbsp;</td>
        </tr>
        <tr> 
          <td class="enfont2 content">
		  <div><b>���⣺</b><%= rs("utitle") %></div>
		  <div><b>�绰��</b><%= rs("utel") %></div><div><b>Email��</b><%= rs("uemailo") %></div>
		  <div><b>���ݣ�</b><%= rs("ucontent") %></div>
		  </td>
        </tr>
      </table>
  
  
  
  
  
  
  
  
  
  
  
  
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
          <tr> 
            <td><b><font color="#FF3300">����Ա�ظ���</font></b> 
              <textarea name="content" class="input" id="content" style="height:100px"><%= rs("reply") %></textarea></td>
          </tr>
          <tr> 
            <td height="28"><input name="Submit" type="submit" class="bt" value="�� ��"></td>
          </tr>
      </table>

  
  
  
  
  
  
  
  
  
  
  
  
</form>
<% end if %>
</body>
</html>
									  