<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE44line1 {font-size: 14px; line-height: 24px; height:30;font-weight: normal; color: #07519a; border: 1px solid #aaccee;}
-->
</style>
</head>

<body>

<% 
if trim(request.querystring("action"))="list" then
if trim(request.form("submit"))="�� ��" then
	dim title,url,address,body
	title=trim(request.form("title"))
	url=trim(request.form("url"))
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Link"
	rs.open sql,conn,2,3
	rs.addnew
	rs("title")=title
	if request.form("flag")="" then
	rs("flag")=100
	else
	rs("flag")=request.form("flag")
	end if
	rs("url")=url
	rs.update
	rs.requery
	response.write "<script>alert('��ӳɹ�!');location.href='?action=list';</script>"
end if





OrderBy = request("OrderBy") '��ǰ��¼��������� at_no ��at_id�Ǽ�¼��id ���������id������Ƿֿ��� 
at_ID = request("at_ID") '��ǰ��¼id 
action = request("action") '�ƶ����� 
paixu=request("paixu")




if paixu="up" then '���� 
		'�����ж��ǲ����Ѿ��ƶ�����ǰ 
		sql="select top 1 id from jiedai_Link where flag<"&OrderBy&" order by flag desc" 
		set rs=server.createobject("adodb.recordset") 
		rs.open sql,conn,1,3 
		if rs.eof then 'ǰ��û�м�¼ �� 
		rs.close 
		set rs=nothing 
		response.write "<script>alert('���󣬸�����Ϣ�Ѿ�λ����λ��');window.history.back();</script>" 
		response.end 
		end if 
		before_id=rs("id") 'ǰһ����¼��id 
		rs.close 
		set rs=nothing 

		'�޸�ǰһ����¼��id 
		sql2="select flag from jiedai_Link where id="&before_id '��ѯǰһ����¼ 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		before_Orderby=rs2("flag") '��¼�ü�¼��������� 
		rs2("flag")=OrderBy '���ĸü�¼��������� 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_Link where ID="&at_ID '��ѯ��ǰ��¼��Ҫ�ƶ��ļ�¼�� 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("flag")=before_Orderby '�޸ĵ�ǰ��¼���������Ϊǰһ��¼����ţ����� ������¼��Ž��� ʵ������ 
		rs3.update 
		rs3.close 
		set rs3=nothing 

elseif paixu="Down" then '���������Ƶ���һ�� 
		sql="select top 1 id from jiedai_Link where flag>"&OrderBy&" order by flag asc" 
		set rs=server.createobject("adodb.recordset") 
		rs.open sql,conn,1,3 
		if rs.eof then 
		rs.close 
		set rs=nothing 
		response.write "<script>alert('���󣬸�����Ϣ�Ѿ�λ�����һλ��');window.history.back();</script>" 
		response.end 
		end if 
		end_id=rs("ID") 
		'response.Write(end_id) 
		'response.End() 
		rs.close 
		set rs=nothing 
		
		sql2="select flag from jiedai_Link where ID="&end_id 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		end_Orderby=rs2("flag") 
		rs2("flag")=OrderBy 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_Link where ID="&at_ID 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("flag")=end_Orderby 
		rs3.update 
		rs3.close 
		set rs3=nothing 
end if 




 %>
<script language="javascript" type="text/javascript">
function check_add(){
var notnull;
notnull=true;
if (document.form1.title.value==""){
notnull=false;
alert("��վ���Ʋ���Ϊ�գ�");
document.form1.title.focus();
}
else
if (document.form1.url.value==""){
notnull=false;
alert("��ַ����Ϊ�գ�");
document.form1.url.focus();
}
return notnull;
}
</script>
<br>
<table width="100%" height="115" border="0" cellpadding="0" cellspacing="0" class="input" style="padding:10px">
  <tr>
    <td><table width="100%" height="174" border="0" cellpadding="0" cellspacing="0">
      <form name="form1" id="form2" method="post" action="?action=list" onSubmit="return check_add()">
        <tr>
          <td width="305" align="right">��վ���ƣ�<font color="#ff0000">[����]</font></td>
          <td style="padding-left:6px;"><input name="title" type="text" class="STYLE44line1" id="title" size="30" style="cursor:hand">
              <font color="#ff0000">*</font></td>
        </tr>
        <tr>
          <td align="right">��ַ��ַ��<font color="#ff0000">[����]</font></td>
          <td style="padding-left:6px;"><input name="url" type="text" class="STYLE44line1" id="url" size="30" style="cursor:hand">
              <font color="#ff0000">*</font> </td>
        </tr>
        <tr bordercolor="#215dc6">
          <td align="right">����</td>
          <td style="padding-left:8px"><input name="flag" type="text" class="STYLE44line1" id="flag" size="8"></td>
        </tr>

        <tr>
          <td colspan="2" align="center" background="images/bg_title.gif" style="padding-right:40px"><input type="submit" name="submit" value="�� ��" style="width:80; height:30; cursor:hand">
            &nbsp;
            <input type="reset" name="submit3" value="�� ��" style="width:80; height:30; cursor:hand"></td>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%" height="115" border="0" cellpadding="0" cellspacing="0" class="input" style="padding:10px">
  <tr>
    <td><br>
      <%
	  



set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Link order by flag asc"
	rs.open sql,conn,1,1
	rs.pagesize=6
	if not rs.eof then
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
	  
	  
	  
if rs.bof then
response.write("<center><br><br><br><br><br><br>")
response.write("<font color=red>����</font>���ӣ�")
response.write("<br><br><br><br><br><br></center>")
end if
for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td width="338" align="left" style="padding-left:6px"><%=i%>-<a href="<%=rs("url")%>" target=_blank><%=rs("title")%></a></td>
          <td>&nbsp;<%= rs("url")  %></td>
          <td width="98"><table width="60" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="30"><a href="?action=list&paixu=up&OrderBy=<%=rs("flag")%>&at_id=<%=rs("id")%>"><img src="images/up.gif" width="15" height="16" border="0"></a> ��</td>
              <td width="30"><a href="?action=list&paixu=Down&OrderBy=<%=rs("flag")%>&at_id=<%=rs("id")%>"><img src="images/down.gif" width="15" height="16" border="0"></a></td>
            </tr>
          </table></td>
          <td width="50" align="center"><a href="?action=del&id=<%= rs("id") %>" onClick="return window.confirm('ȷ��ɾ����?');">[ɾ��]</a></td>
          <td width="50" align="center"><a href="?action=edit&id=<%= rs("id")%>&page=<%=page%>">[�޸�]</a></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" class="xuanxian">
        <tr>
          <td height="4"></td>
        </tr>
      </table>
      <% 
rs.movenext
next
end if
%>
      <br>
      <table width="100%" height="30" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
        <tr>
          <td align="center">��<%= page %>ҳ&nbsp;
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
            <% if page<rs.recordcount then %>
            <a href="?action=list&page=<%= rs.recordcount %>">ĩҳ</a>
            <% else %>
            ĩҳ
            <% end if %>
            &nbsp;����<%= rs.recordcount %>��</td>
          <td width="200" align="center">go
            <select name="select2" onchange='javascript:window.open(this.options[this.selectedindex].value,"_self")'>
                <%for m = 1 to rs.pagecount%>
                <option value="?action=list&page=<%=m%>"><%=m%></option>
                <% next %>
              </select>
            ҳ</td>
        </tr>
      </table>
    <% end if %></td>
  </tr>
</table>
<br>
<% 
if trim(request.querystring("action"))="del" then
		id=trim(request.querystring("id"))
		id=replacebadchar(id)
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_Link where id="&id
		rs.open sql,conn,1,3
		rs.delete
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		response.write "<script>alert('ɾ���ɹ�!');location='?action=list'</script>"
	end if
 %>
    
    
    
    
    
    
    
    
    
    <% 
if trim(request.querystring("action"))="edit" then
	if trim(request.form("submit"))="�� ��" then
		id=trim(request.querystring("id"))
		id=replacebadchar(id)
		title=trim(request.form("title"))
		url=trim(request.form("url"))
		page=trim(request.form("page"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_Link where id="&id
		rs.open sql,conn,1,3
		rs("title")=title
		rs("url")=url
		rs("flag")=request.form("flag")
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		response.write "<script>alert('�޸ĳɹ�!');location='?action=list&page="&page&"'</script>"
	end if
	id=trim(request.querystring("id"))
	sql="select * from jiedai_Link where id="&id
	set rs=conn.execute(sql)
 %>
<script language="javascript" type="text/javascript">
function check_edit(){
var notnull;
notnull=true;
if (document.form1.title.value==""){
notnull=false;
alert("��վ���Ʋ���Ϊ�գ�");
document.form1.url.focus();
}
else
if (document.form1.url.value==""){
notnull=false;
alert("��ַ����Ϊ�գ�");
document.form1.url.focus();
}
return notnull;
}
</script>
<br>
<table width="100%" height="171" border="0" cellpadding="0" cellspacing="0" bordercolor="#215dc6">
  <form name="form1" id="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit()">
    <tr> 
      <td width="350" align="right">��վ���ƣ�<font color="#ff0000">[����]</font></td>
      <td style="padding-left:8px"> <input name="title" type="text" class="STYLE44line1" id="title"  value="<%= rs("title") %>" size="30"> 
        <font color="#ff0000">*</font></td>
    </tr>
    <tr> 
      <td align="right">��ַ��ַ��<font color="#ff0000">[����]</font></td>
      <td style="padding-left:8px"> <input name="url" type="text" class="STYLE44line1" id="url"  value="<%= rs("url") %>" size="30"> 
      <font color="#ff0000">*</font>      </td>
    </tr>
    <tr>
      <td align="right">����</td>
      <td style="padding-left:8px"><input name="flag" type="text" class="STYLE44line1" id="flag" value="<%= rs("flag") %>" size="8"></td>
    </tr>
    <tr> 
      <td colspan="2" align="center" background="images/bg_title.gif"> 
        <input name="page" type="hidden" id="page" value="<%=request.querystring("page")%>">
        <input type="submit" name="submit" value="�� ��" style="width:80; height:30; cursor:hand"> 
      &nbsp; <input type="button" name="submit2" value="�� ��" onClick="javascript:history.go(-1)" style="width:80; height:30; cursor:hand"></td>
    </tr>
  </form>
</table>
<% end if %>  

</body>
</html>                                                                                