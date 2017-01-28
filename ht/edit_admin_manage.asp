<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<//!--#include file="session.asp" -->
<!--#include file="md5.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<%
	admin=trim(request.form("admin"))
	password1=trim(request.form("password1"))
	
'''''''''''''''''''''''''''''''''''''''''''''''	
if trim(request.form("submit"))="修 改" then

jid=request.Form("jid")
	set rs=server.createobject("adodb.recordset")
		sql="select * from Jiedai_admin where id="&jid&""
		rs.open sql,conn,1,3
		admin=admin
		password1=md5(password1)
		rs("admin")=admin
		rs("password")=password1
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		
		response.write "<script>alert('修改成功!');location='admin_Manage.asp'</script>"
		response.end
	end if
''''''''''''''''''''''''''''''''''''''''''''''
 %>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_add(){
	var notnull;
	notnull=true;
	if (document.form1.admin.value==""){
		alert("用户名不能为空！");
		document.form1.admin.focus();
		notnull=false;
		}
	else
	if (document.form1.password1.value==""){
		alert("密码不能为空！");
		document.form1.password1.focus();
		notnull=false;
		}
	return notnull;
	}
</script>

<%

jid=request.QueryString("jid")
user_ename=request.QueryString("user_ename")
set res=server.createobject("adodb.recordset")
esql="select * from jiedai_Admin where id="&jid&""
res.open esql,conn,3,2

%>


<table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr>
    <td><table width="100%" height="181" border="0" cellpadding="6" cellspacing="0">
        <form name="form1" method="post" action="?action=add" onSubmit="return check_add();">
          <tr align="center"> 
            <td height="25" colspan="2" bgcolor="#DFEFFF"><font color="#215dc6"><strong>修改管理员</strong></font></td>
          </tr>
          <tr> 
            <td width="80" align="center">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td align="center">用户名：</td>
            <td><input name="admin" type="text" id="admin" size="15" style="width:150px; height:28px" value="<%=res("admin")%>" ></td>
          </tr>
          <tr> 
            <td align="center">密&nbsp;&nbsp;码：</td>
            <td><input name="password1" type="password" id="password1" size="15" style="width:150px; height:28px" value="<%=res("password")%>" ></td>
          </tr>
          <tr> 
            <td colspan="2" align="center">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" style="padding-left:55px"><input type="submit" name="submit" value="修 改" style="width:80px; height:35px"> 
              &nbsp; <input type="reset" name="submit2" value="重 置" style="width:80px; height:35px">
              <input name="jid" type="hidden" id="jid" value="<%=request.querystring("jid")%>"></td>
          </tr>
          <tr> 
            <td colspan="2" align="center">&nbsp;</td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>


</body>
</html>
