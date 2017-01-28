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
        <td align="left" bgcolor="#D3E5FA" style="padding-left:15"><strong><font color="#215dc6">招聘信息</font></strong></td>
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
response.write("<font color=red>暂无</font>招聘信息！")
response.write("<br><br><br><br><br><br></center>")
end if
for i=1 to rs.pagesize
if rs.eof then exit for 
 %>
      <table width="500" height="140" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td width="120" height="20" align="center">工作岗位</td>
          <td>&nbsp;<%= rs("招聘职位") %></td>
        </tr>
        <tr>
          <td height="20" align="center">招聘人数</td>
          <td>&nbsp;<%= rs("招聘人数") %></td>
        </tr>
        <tr>
          <td height="20" align="center">工作地点</td>
          <td>&nbsp;<%= rs("工作地点") %></td>
        </tr>
        <tr>
          <td height="20" align="center">工资待遇</td>
          <td>&nbsp;<%= rs("工资待遇") %></td>
        </tr>
        <tr>
          <td height="20" align="center">有效日期</td>
          <td>&nbsp;<%= rs("截止日期") %> </td>
        </tr>
        <tr>
          <td height="20" align="center">相关需求</td>
          <td>&nbsp;<%= rs("招聘要求") %></td>
        </tr>
        <tr>
          <td height="20" align="center"><a href="?action=del&id=<%= rs("id") %>"><b>[删除]</b></a>&nbsp;&nbsp;<a href="?action=edit&id=<%= rs("id") %>"><b>[修改]</b></a></td>
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
            <% if page<rs.recordcount then %>
            <a href="?action=list&page=<%= rs.recordcount %>&classid=<%= classid %>">末页</a>
            <% else %>
            末页
            <% end if %>
            &nbsp;总数<%= rs.recordcount %>条 </td>
        </tr>
      </table></td>
  </tr>
</table>
<% end if %>

 
 
 
 
<% 
if trim(request.querystring("action"))="add" then
if trim(request.form("submit"))="提交"then
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
	rs("招聘职位")=gw
	rs("招聘人数")=r
	rs("工作地点")=dd
	rs("cnen")=request.form("cnen")
	rs("工资待遇")=dr
	rs("截止日期")=da
	rs("招聘要求")=body
	rs("stype")=request.form("stype")
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('添加成功!');location='?action=list'</script>"
end if
 %>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_add(){
	var notnull;
	notnull=true;
	if (document.form1.gw.value==""){
		alert("工作岗位不能为空！");
		document.form1.gw.focus();
		notnull=false;
		}
	else
	if (document.form1.r.value==""){
		alert("招聘人数不能为空！");
		document.form1.r.focus();
		notnull=false;
		}
	else
	if (document.form1.dd.value==""){
		alert("工作地点不能为空！");
		document.form1.dd.focus();
		notnull=false;
		}
	else
	if (document.form1.dr.value==""){
		alert("工资待遇不能为空！");
		document.form1.dr.focus();
		notnull=false;
		}
		
	else
	if (document.form1.da.value==""){
		alert("截止日期不能为空！");
		document.form1.da.focus();
		notnull=false;
		}		
	else
	if (document.form1.body.value==""){
		alert("招聘要求不能为空！");
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
          <td align="left" bgcolor="#D3E5FA" style="padding-left:15"><strong><font color="#215dc6">添加招聘</font></strong></td>
        </tr>
      </table>
        <br>
        <table width="500" height="140" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td height="20" align="center">所属类别</td>
            <td><select name="stype" id="stype">
              <option value="1">社会招聘</option>
              <option value="2">校园招聘</option>
            </select>
            </td>
          </tr>
          <tr>
            <td width="120" height="20" align="center">工作岗位</td>
            <td><input name="gw" type="text" id="gw" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">招聘人数</td>
            <td><input name="r" type="text" id="r" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">工作地点</td>
            <td><input name="dd" type="text" id="dd2" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">工资待遇</td>
            <td><input name="dr" type="text" id="dr2" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">有效日期</td>
            <td><input name="da" type="text" id="da2" size="25">
              格式:2009-5-1</td>
          </tr>
          <tr>
            <td height="20" align="center">相关需求</td>
            <td><textarea name="body" cols="40" rows="6" id="body2"></textarea></td>
          </tr>
        </table>
        <br>
        <table width="500" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
          <tr>
            <td align="center"><input name="submit" type="submit" id="submit" value="提交">
              &nbsp;
              <input type="reset" name="submit3" value="重置">            </td>
          </tr>
        </table></td>
    </tr>
  </table>
</form> 
<% end if %>
 
 
 
<% 
if trim(request.querystring("action"))="edit" then
if trim(request.form("submit"))="修改" then
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
	rs("招聘职位")=gw
	rs("招聘人数")=r
	rs("工作地点")=dd
	rs("cnen")=request.form("cnen")
	rs("工资待遇")=dr
	rs("截止日期")=da
	rs("招聘要求")=body
	rs("stype")=request.form("stype")
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('修改成功!');location='?action=list'</script>"
end if
	id=trim(request.querystring("id"))
	sql="select * from recruit where id="&id
	set rs=conn.execute(sql)
 %>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.gw.value==""){
		alert("工作岗位不能为空！");
		document.form1.gw.focus();
		notnull=false;
		}
	else
	if (document.form1.r.value==""){
		alert("招聘人数不能为空！");
		document.form1.r.focus();
		notnull=false;
		}
	else
	if (document.form1.dd.value==""){
		alert("工作地点不能为空！");
		document.form1.dd.focus();
		notnull=false;
		}
	else
	if (document.form1.dr.value==""){
		alert("工资待遇不能为空！");
		document.form1.dr.focus();
		notnull=false;
		}
		
	else
	if (document.form1.da.value==""){
		alert("截止日期不能为空！");
		document.form1.da.focus();
		notnull=false;
		}		
	else
	if (document.form1.body.value==""){
		alert("招聘要求不能为空！");
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
            <td align="left" bgcolor="#D3E5FA" style="padding-left:15"><strong><font color="#215dc6">修改招聘</font></strong></td>
          </tr>
        </table>
        <br>
        <table width="100%" height="140" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td height="20" align="center">所属类别</td>
            <td><select name="stype" id="stype">
                <option value="1" <%if rs("stype")=1 then response.write"selected"%>>社会招聘</option>
                <option value="2" <%if rs("stype")=2 then response.write"selected"%>>校园招聘</option>
              </select>
            </td>
          </tr>
          <tr>
            <td width="120" height="20" align="center">工作岗位</td>
            <td><input name="gw" type="text" value="<%= rs("招聘职位") %>" id="gw" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">招聘人数</td>
            <td><input name="r" type="text" value="<%= rs("招聘人数") %>" id="r" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">工作地点</td>
            <td><input name="dd" type="text"  value="<%= rs("工作地点") %>"  id="dd" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">工资待遇</td>
            <td><input name="dr" type="text"  value="<%= rs("工资待遇") %>" id="dr" size="25"></td>
          </tr>
          <tr>
            <td height="20" align="center">有效日期</td>
            <td><input name="da" type="text" value="<%= rs("截止日期") %>" id="da" size="25">
              格式:2005-5-1</td>
          </tr>
          <tr>
            <td height="20" align="center">相关需求</td>
            <td><textarea name="body" cols="40" rows="6" id="body"><%= rs("招聘要求") %></textarea></td>
          </tr>
        </table>
        <br>
        <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
          <tr>
            <td align="center"><input name="submit" type="submit" id="submit" value="修改">
              &nbsp;
              <input type="reset" name="submit4" value="返回" onClick="javascript:history.go(-1)">            </td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
<% end if %>

<% 
if trim(request.querystring("action"))="del" then
if trim(request.form("submit"))="删 除" then
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
	response.write "<script>alert('删除成功!');location='?action=list'</script>"
end if
 %>
<br><br><br><br><br><br><br><br><br><br><br>
确认<font color="#ff0000">删除招聘信息</font>吗？<br>
<form name="form1" method="post" action="?action=del&id=<%= trim(request.querystring("id")) %>">
  <table width=21% border=0 align="center" cellpadding=0 
cellspacing=0 bordercolor=#9cacd0 class=table_out>
    <tr align="center">  
      <td height=15><input type="submit" name="submit" value="删 除"></td>
      <td><input type="reset" name="submit2" value="取 消" onClick="javascript:history.go(-1)"></td>
    </tr>
  </table>
</form>  
<% end if %>
  
</body>
</html>
                                                                                                                          