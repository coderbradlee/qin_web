<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body>









      <% 
if trim(request.querystring("zhiding"))="zdyes" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Down where id="&id
	rs.open sql,conn,1,3
	rs("dhide")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('������!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	response.end
end if
 %>


      <% 
if trim(request.querystring("zhiding"))="zdno" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Down where id="&id
	rs.open sql,conn,3,2
	rs("dhide")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('����ʾ!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	response.end
end if
 %>
 












<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td><%
if trim(request.querystring("action"))="add" then
if trim(request.form("submit"))="���" then
	title=trim(request.form("title"))
	image=trim(request.form("image"))	
	for i = 1 to request.form("content1").count
	  scontent = scontent & request.form("content1")(i)
	next
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Down"
	rs.open sql,conn,1,3
	rs.addnew
	rs("title")=title
	rs("daddress")=image
	rs("anclassid")=int(request("anclassid")) '����
	rs("dtime")=now()
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('��ӳɹ�!');location='?action=list'</script>"
end if
%>
      <script language="javascript" type="text/javascript">
// ��֤�û���������
function check_add(){
	var notnull;
	notnull=true;
	if (document.myform.title.value==""){
		alert("��Ϣ���Ʋ���Ϊ�գ�");
		document.myform.title.focus();
		notnull=false;
		}
	else
	if (document.myform.image.value==""){
		alert("���ص�ַ����Ϊ�գ�");
		document.myform.image.focus();
		notnull=false;
		}
	return notnull;
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">�����Ϣ</font></strong></td>
        </tr>
      </table>
      <br>
      <form name="form1" method="post" action="?action=add" onSubmit="return check_add();">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr height="30">
            <td align="center">����:</td>
            <td><%
	  set rs=server.CreateObject("adodb.recordset")
     rs.open "select * from fwclass order by anclassidorder",conn,1,1
	if rs.eof and rs.bof then
	response.write "���������Ŀ��"
	response.end
	else
  %>
                <select name="anclassid" size="1" id="anclassid">
                  <option selected value="<%=rs("anclassid")%>"><%=trim(rs("anclass"))%></option>
                  <%
        dim selclass
         selclass=rs("anclassid")
        rs.movenext
        do while not rs.eof
	%>
                  <option value="<%=rs("anclassid")%>"><%=trim(rs("anclass"))%></option>
                  <%
        rs.movenext
        loop
		end if
        rs.close
	%>
              </select></td>
          </tr>
          <tr>
            <td width="100" align="center">��Ϣ���ƣ�</td>
            <td><input name="title" type="text" id="title" size="40" style="height:30">
                <font color="#ff0000">*[20��]</font></td>
          </tr>
          <tr>
            <td align="center">��Ϣ��ַ��</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%"><input name="image" type="text" id="imagebig" size="40" style="height:30"></td>
                <td width="67%"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
          
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="���" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit2" value="����" style="width:80; height:30; cursor:hand"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <%
if trim(request.querystring("action"))="edit" then
if trim(request.form("submit"))="�޸�" then
	id=trim(request.querystring("id"))
	title=trim(request.form("title"))
	image=trim(request.form("image"))	
	for i = 1 to request.form("content1").count
	  scontent = scontent & request.form("content1")(i)
	next
	dim rs,sql
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Down where id="&id
	rs.open sql,conn,1,3
	rs("title")=title
	rs("daddress")=image
	rs("anclassid")=int(request("anclassid")) '����
	rs("dtime")=now()
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�޸ĳɹ�!');location='?action=list'</script>"
end if
id=trim(request.querystring("id"))
id=replacebadchar(id)
set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Down where id="&id
	rs.open sql,conn,1,1
%>
      <script language="javascript" type="text/javascript">
// ��֤�û���������
function check_edit(){
	var notnull;
	notnull=true;
	if (document.myform.title.value==""){
		alert("���ⲻ��Ϊ�գ�");
		document.myform.title.focus();
		notnull=false;
		}
	else
	if (document.myform.image.value==""){
		alert("�ļ�����Ϊ�գ�");
		document.myform.image.focus();
		notnull=false;
		}
	return notnull;
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">�޸���Ϣ</font></strong></td>
        </tr>
      </table>
      <br>
      <form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit();">
        <table width="100%" height="25" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td align="center">����:</td>
            <td><%dim rs1
	  set rts=server.CreateObject("adodb.recordset")
			   		set rs1=server.CreateObject("adodb.recordset")
					rs1.open "select * from jiedai_down where id="&id,conn,1,1
					rts.open "select * from fwclass order by anclassidorder",conn,1,1
					if rts.eof and rts.bof then
					response.write "���������Ŀ��"
					response.end
					else
				%>
                <select name="anclassid" size="1" id="anclassid" onChange="changelocation(document.myform.anclassid.options[document.myform.anclassid.selectedIndex].value)">
                  <%do while not rts.eof%>
                  <option value="<%=rts("anclassid")%>" <%if rs1("anclassid")=rts("anclassid") then%>selected<%end if%>><%=trim(rts("anclass"))%></option>
                  <%
					rts.movenext
					loop
					end if
					rts.close
				%>
              </select></td>
          </tr>
          <tr>
            <td width="100" align="center">��Ϣ���ƣ�</td>
            <td><input name="title" type="text" id="title" value="<%= rs("title") %>" size="40" style="height:30">
                <font color="#ff0000">*[20������]</font></td>
          </tr>
          <tr>
            <td align="center">��Ϣ��ַ��</td>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%"><input name="image" type="text" id="image" value="<%= rs("daddress") %>" size="40" style="height:30"></td>
                <td width="67%"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
          
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="�޸�" style="width:80; height:30; cursor:hand">
              &nbsp;
            <input type="reset" name="submit2" value="����" onClick="history.go(-1)" style="width:80; height:30; cursor:hand"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <% 
if trim(request.querystring("action"))="del" then
	id=trim(request.querystring("id"))
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Down where id="&id
	rs.open sql,conn,1,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�޸ĳɹ�!');location='?action=list'</script>"
end if
 %>
 
 
 
 
 
 <%
if trim(request.querystring("action"))="list" then
	classid=trim(request.querystring("cid"))
	set rs=server.createobject("adodb.recordset")
	on error resume next
	sql="select * from jiedai_Down where 1=1 "
	
	if classid<>"" then
	sql=sql+" and anclassid="&classid&""
	
	end if
	
	sql=sql+" order by id desc"
	
	rs.open sql,conn,1,1
	rs.pagesize=10
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
 
 
 
 
 
 
 
      <table width="100%" height="59" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="17%" align="center" valign="top"><table width="100%" height="107" border="0" cellpadding="8" cellspacing="1" bgcolor="#DFEFFF">
            <tr>
              <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                  <tr>
                    <td align="center" bgcolor="#DFEFFF" style="font-size:16px"><b>��������</b></td>
                  </tr>
                </table>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="7"></td>
                    </tr>
                  </table>
                <% 	  
sql_classid="select * from fwclass"
set rs_classid=conn.execute(sql_classid)  
 %>
                  <% do while not rs_classid.eof %>
                  <a href="?action=list&cid=<%= rs_classid("anclassid") %>"><%= rs_classid("anclass") %></a>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="7"></td>
                    </tr>
                  </table>
                <%
		 rs_classid.movenext
		loop
		rs_classid.close
		 %>
              </td>
            </tr>
          </table></td>
          <td width="83%" valign="top">
            <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
              <tr align="center">
                <td width="50" align="center" bgcolor="#DFEFFF">���</td>
                <td bgcolor="#DFEFFF" align="left">��Ϣ���� </td>
                <td width="70" align="center" bgcolor="#DFEFFF">��ַ</td>
                <td width="70" align="center" bgcolor="#DFEFFF">�Ƿ���ʾ</td>
                <td width="50" align="center" bgcolor="#DFEFFF">ɾ��</td>
                <td width="50" align="center" bgcolor="#DFEFFF">�༭</td>
              </tr>
            </table>
            <br>
            <%
if rs.bof then
response.write("<center><br><br><br><br><br><br>")
response.write("<font color=red>����</font>��Ϣ��")
response.write("<br><br><br><br><br><br></center>")
end if

 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
            <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
              <tr>
                <td width="50" align="center" class="line"><%= rs("id") %></td>
                <td class="line">&nbsp;<%= rs("title") %></td>
                <td width="70" align="center" class="line"><a href="../uploadfile/<%=rs("daddress")%>" target="_blank"><img src="images/Search.gif" alt="�鿴��ַ,���һ�ѡ��-���ƿ�ݷ�ʽ!" width="20" height="20" border="0"></a></td>
                <td width="70" align="center" class="line"><%if rs("dhide")=0 then%>
                    <a href="?Action=list&zhiding=zdyes&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/Ok.gif" alt="����ʾ" width="16" height="16" border="0" /></a>
                    <%else%>
                    <a href="?Action=list&zhiding=zdno&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/err.gif" alt="������" width="12" height="11" border="0" /></a>
                    <%end if%></td>
                <td width="50" align="center" class="line"><a href="?action=del&id=<%= rs("id") %>" onClick="return window.confirm('ȷ��ɾ����?');">[ɾ��]</a> </td>
                <td width="50" align="center" class="line"><a href="?action=edit&id=<%= rs("id") %>">[�༭]</a> </td>
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
%>
            <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
              <tr>
                <td align="center">��<%= page %>ҳ&nbsp;
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
                  <% if page<rs.recordcount then %>
                  <a href="?action=list&page=<%= rs.recordcount %>&classid=<%= classid %>">ĩҳ</a>
                  <% else %>
                  ĩҳ
                  <% end if %>
                  &nbsp;����<%= rs.recordcount %>��</td>
              </tr>
            </table>
          <% end if %></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>