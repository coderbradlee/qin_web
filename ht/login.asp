<%@LANGUAGE="VBSCRIPT" CODEPAGE=936%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="conn.asp" -->
<!--#include file="functions.asp" -->
<!--#include file="md5.asp" -->
<% 
if trim(request.QueryString("action"))="login" then
	dim rs,sql
	dim admin,password,checkcode
	admin=trim(request.form("admin"))
	password=trim(request.form("password"))
	checkcode=trim(request.form("checkcode"))
	if checkcode<>session("checkcode") then
		response.write("<script>alert('验证码不正确！');history.go(-1)</script>") 
		response.end
	end if 
	admin=replacebadchar(admin)
	password=replacebadchar(password)
	admin=admin
	password=md5(password)
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_Admin where password='"&password&"' and admin='"&admin&"'"
	rs.open sql,conn,1,3
		
 	if not(rs.bof and rs.eof) then
			rs("jsq")=rs("jsq")+1
			rs("dltime")=now()
			rs.update
			rs.requery
			session("admin")=admin
			session("password")=password
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			response.redirect("index.asp")
	else
		response.write "<script>alert('非法操作:无此用户!');history.go(-1)</script>"
	end if
end if
 %>

<html><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<TITLE>管理员登陆</TITLE>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<STYLE type=text/css>BODY {
MARGIN: 0px; BACKGROUND-COLOR: #fafafa
}
.inc01 {
BORDER-RIGHT: #d2e4ee 1px solid; BORDER-TOP: #d2e4ee 1px solid; BORDER-LEFT: #d2e4ee 1px solid; WIDTH: 153px; BORDER-BOTTOM: #d2e4ee 1px solid; HEIGHT: 20px; BACKGROUND-COLOR: #f3f8fc
}
.inc02 {
BORDER-RIGHT: #d2e4ee 1px solid; BORDER-TOP: #d2e4ee 1px solid; BORDER-LEFT: #d2e4ee 1px solid; WIDTH: 80px; BORDER-BOTTOM: #d2e4ee 1px solid; HEIGHT: 20px; BACKGROUND-COLOR: #f3f8fc
}
.STYLE6 {
COLOR: #0066ff
}
.STYLE7 {
COLOR: #003584
}
.STYLE8 {
COLOR: #09155e
}
.STYLE9 {
FONT-WEIGHT: bold; COLOR: #ff6600
}
.STYLE10 {
FONT-WEIGHT: bold; COLOR: #c43309
}
.STYLE12 {
COLOR: #000033
}
.STYLE13 {
COLOR: #333333
}
</style>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check(){
	var notnull;
	notnull=true;
	if (document.form1.admin.value==""){
		alert("用户名不能为空！");
		document.form1.admin.focus();
		notnull=false;
		}
	else
	if (document.form1.password.value==""){
		alert("密码不能为空！");
		document.form1.password.focus();
		notnull=false;
		}
	else
	if (document.form1.checkcode.value==""){
		alert("验证码不能为空！");
		document.form1.checkcode.focus();
		notnull=false;
		}		
	return notnull;
	}
</script>
</HEAD>
<BODY onLoad="document.form1.admin.focus()">
<FORM id="form1" name="form1" method="post" action="?action=login"  onsubmit="return check();">
<DIV>
<TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
<TBODY>
<TR>
<TD vAlign=top background=images/logins_02.jpg height=604>
<TABLE cellSpacing=0 cellPadding=0 width=499 align=center border=0>
<TBODY>
<TR>
<TD height=35></TD></TR>
<TR>
<TD>&nbsp;</TD>
</TR>
<TR>
<TD><IMG height=126 alt="" src="images/logins_13.jpg" 
width=500></TD></TR>
<TR>
<TD vAlign=top background=images/logins_18.jpg height=368>
<TABLE cellSpacing=0 cellPadding=0 width="96%" align=center 
border=0>
<TBODY>
<TR>
<TD vAlign=top height=74>
<TABLE cellSpacing=0 cellPadding=0 width="96%" align=center 
border=0>
<TBODY>
<TR>
<TD vAlign=top>
<TABLE height="68%" cellSpacing=0 cellPadding=0 
width="100%" align=center border=0>
<TBODY>
<TR>
<TD>
<TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
<TBODY>
<TR>
<TD>
<TABLE cellSpacing=0 cellPadding=0 width=350 
align=center border=0>
<TBODY>
<TR>
<TD align=middle width=322 valign="middle">
<TABLE cellSpacing=0 cellPadding=0 width=320 
border=0>
<TBODY>
<TR>
<TD width="13%" height=30><IMG height=30 alt="" 
src="images/User_Login_0_15.jpg" 
width=35></TD>
<TD align=left width="21%">用户名称：</TD>
<TD align=left colSpan=2><SPAN 
style="HEIGHT: 20px"><INPUT class=inc01 
id=admin name=admin> 
</SPAN></TD></TR>
<TR>
<TD height=30><IMG height=30 alt="" 
src="images/User_Login_0_19.jpg" 
width=35></TD>
<TD align=left>用户密码：</TD>
<TD align=left colSpan=2><SPAN 
style="HEIGHT: 20px"><INPUT class=inc01 
id=password type=password name=password> 
</SPAN></TD></TR>
<TR>
<TD height=30><IMG height=30 alt="" 
src="images/User_Login_0_23.jpg" 
width=35></TD>
<TD align=left>验 证 码：</TD>
<TD align=left width="33%"><SPAN 
style="HEIGHT: 20px"><INPUT class=inc02 id=checkcode 
name=checkcode> 
</SPAN></TD>
<TD align=left width="33%"><img src="check/code.asp"></TD></TR>
<TR>
<TD align=middle colSpan=4 height=50><SPAN 
style="HEIGHT: 51px"><INPUT id=imgbtn 
style="WIDTH: 132px; HEIGHT: 32px" type=image 
src="images/User_Login_0_13.gif" 
name=imgbtn value="提 交"> 
</SPAN>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
<TR>
<TD height=13>&nbsp;</TD>
</TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=499 align=center 
border=0><TBODY>
<TR>
<TD>&nbsp;</TD></TR>
<TR>
<TD>&nbsp;</TD></TR>
<TR>
<TD>&nbsp;</TD>
</TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></FORM></BODY></html>
