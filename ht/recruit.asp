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
on error resume next
if trim(request.querystring("action"))="list" then
%>

<table width="100%" height="116" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
      <tr>
        <td align="left" bgcolor="#D3E5FA" style="padding-left:15"><strong><font color="#215dc6">��Ƹ��Ϣ</font></strong></td>
      </tr>
    </table>
      <br>
      <%
dim rs,sql
set rs=server.createobject("adodb.recordset")
    sql="select * from recruit order by id desc"
    rs.open sql,conn,1,1
	rs.pagesize=4
	page=cint(trim(request.querystring("page")))
	if page<1 then
    page=1
	elseif page>rs.pagecount then
	    page=rs.pagecount
	end if
	rs.absolutepage=page
if rs.bof then 
response.write("<center><br><br><br><br><br><br>")
response.write("<font color=red>����</font>��Ƹ��Ϣ��")
response.write("<br><br><br><br><br><br></center>")
end if
for i=1 to rs.pagesize
if rs.eof then exit for 
 %>
      <table width="500" height="140" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td width="120" height="20" align="center">������λ</td>
          <td>&nbsp;<%= rs("��Ƹְλ") %></td>
        </tr>
        <tr>
          <td height="20" align="center">��Ƹ����</td>
          <td>&nbsp;<%= rs("��Ƹ����") %></td>
        </tr>
        <tr>
          <td height="20" align="center">�����ص�</td>
          <td>&nbsp;<%= rs("�����ص�") %></td>
        </tr>
        <tr>
          <td height="20" align="center">���ʴ���</td>
          <td>&nbsp;<%= rs("���ʴ���") %></td>
        </tr>
        <tr>
          <td height="20" align="center">��Ч����</td>
          <td>&nbsp;<%= rs("��ֹ����") %> </td>
        </tr>
        <tr>
          <td height="20" align="center">�������</td>
          <td>&nbsp;<%= rs("��ƸҪ��") %></td>
        </tr>
        <tr>
          <td height="20" align="center"><a href="?action=del&id=<%= rs("id") %>"><b>[ɾ��]</b></a>&nbsp;&nbsp;<a href="?action=edit&id=<%= rs("id") %>"><b>[�޸�]</b></a></td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="1" background="../images/dot.jpg"></td>
        </tr>
      </table>
      <br>
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
            &nbsp;����<%= rs.recordcount %>�� </td>
        </tr>
      </table></td>
  </tr>
</table>
<% end if %>

 
 
 
 
<% 
if trim(request.querystring("action"))="add" then
if trim(request.form("submit"))="�ύ"then
	dim gw,r,dd,dr,da,body
	gw=trim(request.form("gw"))
	r=trim(request.form("r"))
	dd=trim(request.form("dd"))
	dr=trim(request.form("dr"))
	da=trim(request.form("da"))
	body=trim(request.form("body"))
	body=replace(body,chr(13),"<br>")
	set rs=server.createobject("adodb.recordset")
	sql="select * from recruit"
	rs.open sql,conn,1,3
	rs.addnew
	rs("��Ƹְλ")=gw
	rs("��Ƹ����")=r
	rs("�����ص�")=dd
	rs("cnen")=request.form("cnen")
	rs("���ʴ���")=dr
	rs("��ֹ����")=da
	rs("��ƸҪ��")=body
	rs("stype")=request.form("stype")
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
	if (document.form1.gw.value==""){
		alert("������λ����Ϊ�գ�");
		document.form1.gw.focus();
		notnull=false;
		}
	else
	if (document.form1.r.value==""){
		alert("��Ƹ��������Ϊ�գ�");
		document.form1.r.focus();
		notnull=false;
		}
	else
	if (document.form1.dd.value==""){
		alert("�����ص㲻��Ϊ�գ�");
		document.form1.dd.focus();
		notnull=false;
		}
	else
	if (document.form1.dr.value==""){
		alert("���ʴ�������Ϊ�գ�");
		document.form1.dr.focus();
		notnull=false;
		}
		
	else
	if (document.form1.da.value==""){
		alert("��ֹ���ڲ���Ϊ�գ�");
		document.form1.da.focus();
		notnull=false;
		}		
	else
	if (document.form1.body.value==""){
		alert("��ƸҪ����Ϊ�գ�");
		document.form1.body.focus();
		notnull=false;
		}		
	return notnull;
	}
</script>
<form name="form1" method="post" action="?action=add" onSubmit="return check_add();">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
    <tr>
      <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td align="left" bgcolor="#D3E5FA" style="padding-left:15"><strong><font color="#215dc6">�����Ƹ</font></strong></td>
        </tr>
      </table>
        <br>
        <table width="500" height="140" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td height="20" align="center">�������</td>
            <td><select name="stype" id="stype">
              <option value="1">�����Ƹ</option>
              <option value="2">У԰��Ƹ</option>
            </select>
            </td>
          </tr>
          <tr>
            <td width="120" height="20" align="center">������λ</td>
            <td><input name="gw" type="text" id="gw" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">��Ƹ����</td>
            <td><input name="r" type="text" id="r" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">�����ص�</td>
            <td><input name="dd" type="text" id="dd2" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">���ʴ���</td>
            <td><input name="dr" type="text" id="dr2" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">��Ч����</td>
            <td><input name="da" type="text" id="da2" size="25">
              ��ʽ:2009-5-1</td>
          </tr>
          <tr>
            <td height="20" align="center">�������</td>
            <td><textarea name="body" cols="40" rows="6" id="body2"></textarea></td>
          </tr>
        </table>
        <br>
        <table width="500" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
          <tr>
            <td align="center"><input name="submit" type="submit" id="submit" value="�ύ">
              &nbsp;
              <input type="reset" name="submit3" value="����">            </td>
          </tr>
        </table></td>
    </tr>
  </table>
</form> 
<% end if %>
 
 
 
<% 
if trim(request.querystring("action"))="edit" then
if trim(request.form("submit"))="�޸�" then
	id=trim(request.querystring("id"))
	gw=trim(request.form("gw"))
	r=trim(request.form("r"))
	dd=trim(request.form("dd"))
	dr=trim(request.form("dr"))
	da=trim(request.form("da"))
	body=trim(request.form("body"))
	body=replace(body,chr(13),"<br>")
	set rs=server.createobject("adodb.recordset")
	sql="select * from recruit where id="&id
	rs.open sql,conn,1,3
	rs("��Ƹְλ")=gw
	rs("��Ƹ����")=r
	rs("�����ص�")=dd
	rs("cnen")=request.form("cnen")
	rs("���ʴ���")=dr
	rs("��ֹ����")=da
	rs("��ƸҪ��")=body
	rs("stype")=request.form("stype")
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�޸ĳɹ�!');location='?action=list'</script>"
end if
	id=trim(request.querystring("id"))
	sql="select * from recruit where id="&id
	set rs=conn.execute(sql)
 %>
<script language="javascript" type="text/javascript">
// ��֤�û���������
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.gw.value==""){
		alert("������λ����Ϊ�գ�");
		document.form1.gw.focus();
		notnull=false;
		}
	else
	if (document.form1.r.value==""){
		alert("��Ƹ��������Ϊ�գ�");
		document.form1.r.focus();
		notnull=false;
		}
	else
	if (document.form1.dd.value==""){
		alert("�����ص㲻��Ϊ�գ�");
		document.form1.dd.focus();
		notnull=false;
		}
	else
	if (document.form1.dr.value==""){
		alert("���ʴ�������Ϊ�գ�");
		document.form1.dr.focus();
		notnull=false;
		}
		
	else
	if (document.form1.da.value==""){
		alert("��ֹ���ڲ���Ϊ�գ�");
		document.form1.da.focus();
		notnull=false;
		}		
	else
	if (document.form1.body.value==""){
		alert("��ƸҪ����Ϊ�գ�");
		document.form1.body.focus();
		notnull=false;
		}		
	return notnull;
	}
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td valign="top"><br>
      <form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit();">
        <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
          <tr>
            <td align="left" bgcolor="#D3E5FA" style="padding-left:15"><strong><font color="#215dc6">�޸���Ƹ</font></strong></td>
          </tr>
        </table>
        <br>
        <table width="100%" height="140" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td height="20" align="center">�������</td>
            <td><select name="stype" id="stype">
                <option value="1" <%if rs("stype")=1 then response.write"selected"%>>�����Ƹ</option>
                <option value="2" <%if rs("stype")=2 then response.write"selected"%>>У԰��Ƹ</option>
              </select>
            </td>
          </tr>
          <tr>
            <td width="120" height="20" align="center">������λ</td>
            <td><input name="gw" type="text" value="<%= rs("��Ƹְλ") %>" id="gw" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">��Ƹ����</td>
            <td><input name="r" type="text" value="<%= rs("��Ƹ����") %>" id="r" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">�����ص�</td>
            <td><input name="dd" type="text"  value="<%= rs("�����ص�") %>"  id="dd" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">���ʴ���</td>
            <td><input name="dr" type="text"  value="<%= rs("���ʴ���") %>" id="dr" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">��Ч����</td>
            <td><input name="da" type="text" value="<%= rs("��ֹ����") %>" id="da" size="25">
              ��ʽ:2005-5-1</td>
          </tr>
          <tr>
            <td height="20" align="center">�������</td>
            <td><textarea name="body" cols="40" rows="6" id="body"><%= rs("��ƸҪ��") %></textarea></td>
          </tr>
        </table>
        <br>
        <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
          <tr>
            <td align="center"><input name="submit" type="submit" id="submit" value="�޸�">
              &nbsp;
              <input type="reset" name="submit4" value="����" onClick="javascript:history.go(-1)">            </td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<% end if %>

<% 
if trim(request.querystring("action"))="del" then
if trim(request.form("submit"))="ɾ ��" then
	id=trim(request.querystring("id"))
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from recruit where id="&id
	rs.open sql,conn,1,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('ɾ���ɹ�!');location='?action=list'</script>"
end if
 %>
<br><br><br><br><br><br><br><br><br><br><br>
ȷ��<font color="#ff0000">ɾ����Ƹ��Ϣ</font>��<br>
<form name="form1" method="post" action="?action=del&id=<%= trim(request.querystring("id")) %>">
  <table width=21% border=0 align="center" cellpadding=0 
cellspacing=0 bordercolor=#9cacd0 class=table_out>
    <tr align="center">  
      <td height=15><input type="submit" name="submit" value="ɾ ��"></td>
      <td><input type="reset" name="submit2" value="ȡ ��" onClick="javascript:history.go(-1)"></td>
    </tr>
  </table>
</form>  
<% end if %>
  
</body>
</html>
                                                                                                                          