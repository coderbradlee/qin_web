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












OrderBy = request("OrderBy") '��ǰ��¼��������� at_no ��at_id�Ǽ�¼��id ���������id������Ƿֿ��� 
at_ID = request("at_ID") '��ǰ��¼id 
action = request("action") '�ƶ����� 
paixu=request("paixu")




if paixu="up" then '���� 
		'�����ж��ǲ����Ѿ��ƶ�����ǰ 
		sql="select top 1 id from jiedai_fuwu where flag<"&OrderBy&" order by flag desc" 
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
		sql2="select flag from jiedai_fuwu where id="&before_id '��ѯǰһ����¼ 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		before_Orderby=rs2("flag") '��¼�ü�¼��������� 
		rs2("flag")=OrderBy '���ĸü�¼��������� 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_fuwu where ID="&at_ID '��ѯ��ǰ��¼��Ҫ�ƶ��ļ�¼�� 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("flag")=before_Orderby '�޸ĵ�ǰ��¼���������Ϊǰһ��¼����ţ����� ������¼��Ž��� ʵ������ 
		rs3.update 
		rs3.close 
		set rs3=nothing 

elseif paixu="Down" then '���������Ƶ���һ�� 
		sql="select top 1 id from jiedai_fuwu where flag>"&OrderBy&" order by flag asc" 
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
		
		sql2="select flag from jiedai_fuwu where ID="&end_id 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		end_Orderby=rs2("flag") 
		rs2("flag")=OrderBy 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_fuwu where ID="&at_ID 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("flag")=end_Orderby 
		rs3.update 
		rs3.close 
		set rs3=nothing 
end if 


















%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td><% 
classid=trim(request.querystring("classid"))
set rs=server.createobject("adodb.recordset")
sql="select * from jiedai_fuwu order by flag asc"
rs.open sql,conn,1,1
rs.pagesize=10

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
 %>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" bgcolor="#D3E5FA">
  <tr> 
    <td style="padding-left:15"><strong><font color="#215dc6">��Ϣ����</font></strong> </td>
  </tr>
</table><br>
<%
if rs.bof then response.write("<center><br><br><br><br><br><br><font color=red>����</font>��Ϣ��<br><br><br><br><br><br></center>")
 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr align="center">
    <td width="50" align="center" class="line"><%= rs("flag") %></td>
    <td width="950" align="left" class="line"><span class="style1">&nbsp;<%= rs("classid") %></span>   <%if rs("e_classid")<>"" then 
		   response.Write"<img src=""images/en.jpg"" />" 
		   else
		   response.Write"<img src=""images/noen.jpg"" />" 
		   end if
		   %></td>
    <td width="132" align="center" class="line"><table width="60" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="30"><a href="?action=list&paixu=up&OrderBy=<%=rs("flag")%>&at_id=<%=rs("id")%>"><img src="images/up.gif" width="15" height="16" border="0"></a> ��</td>
        <td width="30"><a href="?action=list&paixu=Down&OrderBy=<%=rs("flag")%>&at_id=<%=rs("id")%>"><img src="images/down.gif" width="15" height="16" border="0"></a></td>
      </tr>
    </table></td>
    <td width="50" align="center" class="line"><a href="?action=edit&id=<%= rs("id") %>">[�޸�]</a></td>
    <td width="50" align="center" class="line"><a href="?action=del&id=<%= rs("id") %>" onClick="return window.confirm('ȷ��ɾ����?');">[ɾ��]</a></td>
  </tr>
</table>
<br>
<% 
rs.movenext
next
else
response.write"������Ϣ"
end if
%>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr> 
    <td width="1122" align="center">��<%= page %>ҳ&nbsp; 
      <% if page<>1 then %>
      <a href="?action=list&page=1&classid=<%= classid %>">��ҳ</a> 
      <% else %>
      ��ҳ 
      <% end if %>
      &nbsp; 
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>&classid=<%= classid %>">��һҳ</a> 
      <% else %>
      ��һҳ 
      <% end if %>
      &nbsp; 
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>&classid=<%= classid %>">��һҳ</a> 
      <% else %>
      ��һҳ 
      <% end if %>
      &nbsp; 
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%=rs.pagecount%>&classid=<%= classid %>">ĩҳ</a> 
      <% else %>
      ĩҳ 
      <% end if %>
      &nbsp;����<%= rs.recordcount %>��</td>
    <td width="110" align="center">ת���� 
      <select name="select" onchange='javascript:window.open(this.options[this.selectedindex].value,"_self")'>
        <%for m = 1 to rs.pagecount%>
        <option value="?action=list&page=<%=m%>&classid=<%= classid %>"><%=m%></option>
        <% next %>
      </select>
    ҳ</td>
  </tr>
</table>
</td>
  </tr>
</table>
<% end if %>
<% if trim(request.querystring("action"))="add" then
	if trim(request.form("add"))="add" then
	
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_fuwu"
		rs.open sql,conn,1,3
		rs.addnew
		
		if request.form("flag")<>"" then
		rs("flag")=request.form("flag")
		else
		rs("flag")=100
		end if
		rs("body")=request.form("content")
		rs("classid")=trim(request.form("classid"))
		rs("e_body")=request.form("e_content")
		rs("e_classid")=trim(request.form("e_classid"))
		rs.update
		rs.requery
		rs.close
		set rs=nothing
	end if
%>
<script language="javascript" type="text/javascript">
// ��֤�û���������
function check_add(){
if(document.form1.classid.value==""){alert("����������������");document.form1.classid.focus();return false;}
return true
}
</script>
<form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
  <table width="100%" height="92" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr> 
      <td width="130" height="30" align="right">���ı���:</td>
      <td width="863"><input name="classid" type="text" id="classid" size="40"> 
      ������:
      <input name="flag" type="text" id="flag" size="8"></td>
    </tr>
	
	
	
	<tr> 
      <td width="130" height="30" align="right">��title��Ӣ�ı���:</td>
      <td><input name="e_classid" type="text" id="e_classid" size="40"> 
      ��</td>
    </tr>
	
	
	
	
	
	
    <tr>
      <td height="30" align="center">
	  ��������:</td>
      <td>
      	<textarea name="content" cols="" rows="" style="display:none"></textarea>
	   <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>	
	  </td>
    </tr>
	
	
	
	
	
	
	
	
	
	<tr>
      <td height="30" align="center">
	  Ӣ������:</td>
      <td>
      	<textarea name="e_content" cols="" rows="" style="display:none"></textarea>
	   <iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>	
	  </td>
    </tr>
	
	
	
	
	
	
	
	
    <tr>
      <td height="30" colspan="2" background="images/bg_title.gif" style="padding-left:50"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="9"></td>
        </tr>
      </table>
        <%if request.form("add")="add" then
		 response.write"<img src=images/cms-ico7.gif width=12 height=11><font color=#ff0000><b></b>��Ϣ����ӳɹ�</font>"
		 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table><%end if%><input type="image" name="imageField2" src="images/submit-bt.gif">
<input name="add" type="hidden" id="add" value="add"></td>
    </tr>
  </table>
</form>
<% end if %>








<%

if trim(request.querystring("action"))="edit" then

%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td align="center" valign="top" >
<% 
	if trim(request.form("add"))="add" then
		id=trim(request.querystring("id"))
		classid=trim(request.form("classid"))
		for i = 1 to request.form("content").count
		  scontent = scontent & request.form("content")(i)
		next
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_fuwu where id="&id
		rs.open sql,conn,1,3
		rs("classid")=classid
		rs("body")=scontent
		rs("flag")=request.form("flag")
		rs("e_body")=request.form("e_content")
		rs("e_classid")=trim(request.form("e_classid"))
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		
		
		
	end if
	
		id=trim(request.querystring("id"))
		sql="select * from jiedai_fuwu where id="&id
		set rs=conn.execute(sql)

%>
<script language="javascript" type="text/javascript">
// ��֤�û���������
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.classid.value==""){
		alert("���ⲻ��Ϊ�գ�");
		document.form1.classid.focus();
		notnull=false;
		}
		return notnull;
	}
</script>
<form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit()">
  <table width="100%" height="417" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
 
   
        <input name="add" type="hidden" id="add" value="add">
    
	   <%if request.form("add")="add" then
		 response.write"<Tr><td></td><td height=30><img src=images/cms-ico7.gif width=12 height=11><font color=#ff0000><b>"&rs("classid")&"-</b>��Ϣ���޸ĳɹ�</font></td></tr>"
		 end if
		 %>
	  
	    <tr> 
      <td width="130" height="30" align="right">���ı���:</td>
      <td width="863"><input name="classid" type="text" id="classid" size="40" value="<%= rs("classid") %>"> 
      ������:
      <input name="flag" type="text" id="flag" size="8" value="<%=rs("flag")%>"></td>
    </tr>
	
	
	
	<tr> 
      <td width="130" height="30" align="right">��title��Ӣ�ı���:</td>
      <td><input name="e_classid" type="text" id="e_classid" size="40" value="<%= rs("e_classid") %>"> 
      ��</td>
    </tr>
	  
	  
	  
	  
	      <tr>
      <td height="30" align="center">
	  ��������:</td>
      <td>
      	<textarea name="content" cols="" rows="" style="display:none"><%=rs("body")%></textarea>
	   <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>	
	  </td>
    </tr>
	

	
	
	<tr>
      <td height="30" align="center">
	  Ӣ������:</td>
      <td>
      	<textarea name="e_content" cols="" rows="" style="display:none"><%=rs("e_body")%></textarea>
	   <iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>	
	  </td>
    </tr>
	  

    <tr>
      <td height="30" align="left" valign="top" background="images/bg_title.gif" style="padding-left:50">        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table>
       
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table>
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



    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    </td>
  </tr>
</table>
<% end if %>





<% if trim(request.querystring("action"))="del" then %>
<% 
	id=trim(request.querystring("id"))
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_fuwu where id="&id
	rs.open sql,conn,2,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write("<script>alert('ɾ���ɹ�');location='?action=list';</script>")
end if
 %>
</body>
</html>                                                                             