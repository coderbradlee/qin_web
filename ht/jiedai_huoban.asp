<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<%
sjid=request("sjid")
sid=request("sid")
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
if trim(request.querystring("action"))="list" then

OrderBy = request("OrderBy") '当前记录的排序序号 at_no ，at_id是记录的id 我这个表里id和序号是分开的 
at_ID = request("at_ID") '当前记录id 
action = request("action") '移动方向 
paixu=request("paixu")

if paixu="up" then '上移 
		'首先判断是不是已经移动到最前 
		sql="select top 1 id from jiedai_huoban where flag<"&OrderBy&" and typeclass='"&request.QueryString("typecss")&"' order by flag desc" 
		set rs=server.createobject("adodb.recordset") 
		rs.open sql,conn,1,3 
		if rs.eof then '前边没有记录 了 
		rs.close 
		set rs=nothing 
		response.write "<script>alert('错误，该条信息已经位于首位了');window.history.back();</script>" 
		response.end 
		end if 
		before_id=rs("id") '前一条记录的id 
		rs.close 
		set rs=nothing 

		'修改前一条记录的id 
		sql2="select flag from jiedai_huoban where typeclass='"&request.QueryString("typecss")&"' and id="&before_id '查询前一条记录 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		before_Orderby=rs2("flag") '记录该记录的排序序号 
		rs2("flag")=OrderBy '更改该记录的排序序号 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_huoban where typeclass='"&request.QueryString("typecss")&"' and ID="&at_ID '查询当前记录（要移动的记录） 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("flag")=before_Orderby '修改当前记录的排序序号为前一记录的序号，到此 两条记录序号交换 实现上移 
		rs3.update 
		rs3.close 
		set rs3=nothing 

elseif paixu="Down" then '下移与上移道理一样 
		sql="select top 1 id from jiedai_huoban where typeclass='"&request.QueryString("typecss")&"' and  flag>"&OrderBy&" order by flag asc" 
		set rs=server.createobject("adodb.recordset") 
		rs.open sql,conn,1,3 
		if rs.eof then 
		rs.close 
		set rs=nothing 
		response.write "<script>alert('错误，该条信息已经位于最后一位了');window.history.back();</script>" 
		response.end 
		end if 
		end_id=rs("ID") 
		'response.Write(end_id) 
		'response.End() 
		rs.close 
		set rs=nothing 
		
		sql2="select flag from jiedai_huoban where typeclass='"&request.QueryString("typecss")&"' and  ID="&end_id 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		end_Orderby=rs2("flag") 
		rs2("flag")=OrderBy 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_huoban where typeclass='"&request.QueryString("typecss")&"' and  ID="&at_ID 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("flag")=end_Orderby 
		rs3.update 
		rs3.close 
		set rs3=nothing 
end if 


















%>



<%
	set res=server.createobject("adodb.recordset")
	sqle="select * from jiedai_paijiexi where id="&request.QueryString("sid")&" "
	res.open sqle,conn,1,1


%>





<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td>
 
 
 
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" bgcolor="#D3E5FA">
  <tr> 
    <td style="padding-left:15"><strong><font color="#215dc6"><%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>--<font color="#CC6600"><%=res("jgname")%></font>--优惠信息管理
    </font></strong></td>
  </tr>
</table><br>


<% 
classid=trim(request.querystring("classid"))
set rs=server.createobject("adodb.recordset")
sql="select * from jiedai_huoban where sid="&sid&" order by flag asc"
rs.open sql,conn,1,1
rs.pagesize=15

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
 
 




<%
if rs.bof then response.write("<div align=center><br><br><br><br><br><br><font color=red>暂无</font>信息！<br><br><br><br><br><br></div>")
 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr align="center">
    <td width="50" align="center" class="line"><%= rs("flag") %></td>
    <td width="950" align="left" class="line"><span class="style1">&nbsp;<%= rs("title") %></span>　<font color="#FF0000"><b><%=rs("typeclass")%></b></font></td>
    <td width="132" align="center" class="line">&nbsp;</td>
    <td width="50" align="center" class="line"><a href="?action=edit&id=<%= rs("id") %>&sid=<%=sid%>&sjid=<%=sjid%>">[修改]</a></td>
    <td width="50" align="center" class="line"><a href="?action=del&id=<%= rs("id") %>&sid=<%=sid%>&sjid=<%=sjid%>" onClick="return window.confirm('确定删除吗?');">[删除]</a></td>
  </tr>
</table>
<br>
<% 
rs.movenext
next
else
response.write"<div align=center><br><font color=red>暂无</font>信息！<br><br></div>"
end if
%>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr> 
    <td width="962" align="center">第<%= page %>页&nbsp; 
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
    <td width="270" align="center">转到第 
      <select name="select" onchange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%for m = 1 to rs.pagecount%>
        <option value="?action=list&page=<%=m%>&classid=<%= classid %>"><%=m%></option>
        <% next %>
      </select>
    页</td>
  </tr>
</table>
</td>
  </tr>
</table>
<% end if %>
<% if trim(request.querystring("action"))="add" then
	if trim(request.form("add"))="add" then
	
		sid=request.Form("sid")
		sjid=request.Form("sjid")
		
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_huoban"
		rs.open sql,conn,1,3
		rs.addnew
		rs("sid")=request.Form("sid")
		rs("sjid")=request.Form("sjid")
		rs("title")=request.Form("title")
		rs("typeclass")=request.Form("typeclass")
		rs("content")=request.Form("content")
		
		if request.form("flag")<>"" then
		rs("flag")=request.Form("flag")
		else
		rs("flag")=6
		end if
		
		rs.update
		rs.requery
		
		response.write("<script>alert('添加成功');location='?action=add&sid="&sid&"&sjid="&sjid&"';</script>")
		response.end
		
		
		rs.close
		set rs=nothing
	end if
%>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_add(){
if(document.form1.classid.value==""){alert("请填入您的姓名！");document.form1.classid.focus();return false;}
return true
}
</script>



<%
	set res=server.createobject("adodb.recordset")
	sqle="select * from jiedai_paijiexi where id="&request.QueryString("sid")&" "
	res.open sqle,conn,1,1


%>


<form name="myform" method="post" action="?action=add" onSubmit="return check_add()">
  <table width="100%" height="180" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr>
      <td height="30" colspan="2" align="left"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D3E5FA">
  <tr>
    <td width="9%" bgcolor="#F4FBFF" style="padding-left:12px;"><font color="#3399CC"><%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>
      --</font></td>
    <td width="91%" bgcolor="#F4FBFF" style="padding-left:12px; font-size:18px; font-weight:bold; color:#FF0000; padding-top:4px"><font color="#3399CC">给</font><%=res("jgname")%><font color="#3399CC">添加优惠券</font></td>
  </tr>
</table>
</td>
    </tr>
    <tr>
      <td height="30" align="center">类型:</td>
      <td><select name="typeclass">
        <option value="金额优惠">金额优惠</option>
        <option value="折扣优惠">折扣优惠</option>
        <option value="特色优惠">特色优惠</option>
        <option value="团体优惠">团体优惠</option>
        
        
      </select>      </td>
    </tr>
    <tr> 
      <td width="50" height="30" align="center">名称:</td>
      <td><input name="title" type="text" id="classid" size="40"> 
      　排序:
        <input name="flag" type="text" id="flag" size="8"></td>
    </tr>
    





    <tr>
      <td height="30" align="center">说明:</td>
      <td>
	  <input type="hidden" name="content" value="">
	   <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="200"></iframe>	  </td>
    </tr>
    <tr>
      <td height="30" colspan="2" background="images/bg_title.gif" style="padding-left:50"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="9"></td>
        </tr>
      </table>
        <%if request.form("add")="add" then
		 response.write"<img src=images/cms-ico7.gif width=12 height=11><font color=#ff0000><b></b>信息已添加成功</font>"
		 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table><%end if%><input type="image" name="imageField2" src="images/submit-bt.gif">
<input name="add" type="hidden" id="add" value="add">
<input type="hidden" name="sid" id="sid" value="<%=request.QueryString("sid")%>">
<input type="hidden" name="sjid" id="sjid" value="<%=request.QueryString("sjid")%>"></td>
    </tr>
  </table>
</form>
<% end if %>








<%

if trim(request.querystring("action"))="edit" then

%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td align="center" valign="top">
<% 
	if trim(request.form("add"))="add" then
		id=trim(request.querystring("id"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from jiedai_huoban where id="&id
		rs.open sql,conn,1,3


		'rs("sid")=request.Form("sid")
		'rs("sjid")=request.Form("sjid")
		rs("title")=request.Form("title")
		rs("typeclass")=request.Form("typeclass")
		rs("content")=request.Form("content")
		rs("flag")=request.Form("flag")
		

		rs.update
		rs.requery
		
	response.write("<script>alert('修改成功');location='?action=list&sid="&sid&"&sjid="&sjid&"';</script>")
		response.end

		
		
		rs.close
		set rs=nothing
		
		
		
	end if
	
		id=trim(request.querystring("id"))
		sql="select * from jiedai_huoban where id="&id
		set rs=conn.execute(sql)

%>




<%
	set res=server.createobject("adodb.recordset")
	sqle="select * from jiedai_paijiexi where id="&request.QueryString("sid")&" "
	res.open sqle,conn,1,1


%>




<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_edit(){
	var notnull;
	notnull=true;
	if (document.myform.title.value==""){
		alert("标题不能为空！");
		document.myform.title.focus();
		notnull=false;
		}
		return notnull;
	}
</script>
<form name="myform" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>&sid=<%=sid%>&sjid=<%=sjid%>" onSubmit="return check_edit()">
  <table width="100%" height="417" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr>
      <td height="323" valign="top">
      
      	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
      	  
          <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D3E5FA">
            <tr>
              <td width="9%" bgcolor="#F4FBFF" style="padding-left:12px;"><font color="#3399CC">
                <%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>
                --</font></td>
              <td width="91%" bgcolor="#F4FBFF" style="padding-left:12px; font-size:18px; font-weight:bold; color:#FF0000; padding-top:4px"><font color="#3399CC">给</font><%=res("jgname")%><font color="#3399CC">添加优惠券</font></td>
            </tr>
          </table>
          <table width="100%" height="193" border="0" cellpadding="0" cellspacing="0">
            <tr bordercolor="#cccccc">
              <td height="30" align="center">类型:</td>
              <td><select name="typeclass" id="typeclass">
                <option value="金额优惠" <%if rs("typeclass")="金额优惠" then response.write "selected"%>>金额优惠</option>
                <option value="折扣优惠" <%if rs("typeclass")="折扣优惠" then response.write "selected"%>>折扣优惠</option>
                <option value="特色优惠" <%if rs("typeclass")="特色优惠" then response.write "selected"%>>特色优惠</option>
                <option value="团体优惠" <%if rs("typeclass")="团体优惠" then response.write "selected"%>>团体优惠</option>
              </select></td>
            </tr>

            <tr bordercolor="#cccccc">
              <td width="50" height="30" align="center">名称:</td>
              <td><input name="title" type="text" id="title" value="<%=rs("title")%>" size="40">
                排序:
                <input name="flag" type="text" id="flag" value="<%=rs("flag")%>" size="8"></td>
            </tr>



            <tr>
              <td align="center">简介:</td>
              <td>


<textarea name="content" style="display:none"><%=rs("content")%></textarea>
      <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="240"></iframe></td>
            </tr>
          </table>
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table></td>
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
        <input type="hidden" name="sid" id="sid" value="<%=request.QueryString("sid")%>">
        <input type="hidden" name="sjid" id="sjid" value="<%=request.QueryString("sjid")%>">
        
        <input type="hidden" name="add" id="sjid" value="add">
        
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
	sid=trim(request.querystring("sid"))
	sjid=trim(request.querystring("sjid"))
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_huoban where id="&id
	rs.open sql,conn,2,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
		response.write("<script>alert('删除成功');location='?action=list&sid="&sid&"&sjid="&sjid&"';</script>")
end if
 %>
</body>
</html>                                                                             