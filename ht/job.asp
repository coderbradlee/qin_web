

<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->

<!--#include file="Function_Page.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">

<style>




/*列表分页*/
	#page {width:100%;padding:5px 0;}
	#page ul{margin:0 auto; display:table;text-align:center;}
	#page li{float:left !important; float:none;margin-right: 2px; height:17px;line-height:17px;display:inline; zoom:1;}
	#page span{display: block; padding: 2px 5px; background: #F5FBFF; border: 1px solid #CCC; color: #999999; }
	#page a{font-size:12px;display: block; text-decoration: none; margin:0px; color: #ffffff;padding: 2px 5px 2px 5px;background:#91CF40;}
	#page a:link,#page a:visited {border: 1px solid #CCCCCC; }
	#page a:hover {background:#F5FBFF;color:#135C86;}
	#page #span1 {background:#FFFFFF;display: block;}

</style>


</head>
<script language="javascript">
 function sel(a){ 
  o=document.getElementsByName(a) 
  for(i=0;i<o.length;i++) 
  o[i].checked=event.srcElement.checked 
 }
</script>
<%
function cutstr(thestr,strlen)
thestr=trim(thestr) '忽略字符串前后的空白
thestr_length= len(thestr) '求字符串的长度
if thestr_length > strlen then   '判断字符串的长度
   response.Write left(thestr,strlen)&" ..."  
else
   response.write thestr
end if
end function


Function RemoveHTML(strText) 
Dim RegEx 
Set RegEx = New RegExp 
RegEx.Pattern = "<[^>]*>" 
RegEx.Global = True 
RemoveHTML = RegEx.Replace(strText, "") 
End Function


%>
<body>



<%ss=request("ss")%>


<%if ss="list" or ss="" then%>





<%

	if request("action")="savenew" then
		call savenew()
	elseif request("action")="savedit" then
		call savedit()
	elseif request("action")="tg1" then
		call tg1()
		elseif request("action")="tg2" then
		call tg2()
	elseif request("action")="del"  then
		call del()
	elseif request("action")="delAll" then
		call delAll()
	else
		call List()
	end if

sub List()
%>
<form id="form2" name="form2" method="post" action="?action=delAll" class="border-padd0">
  <table width="98%" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td width="13" bgcolor="#DAECF5">&nbsp;</td>
      <td align="center" bgcolor="#DAECF5" class="fontb">管理招聘信息</td>
      <td width="13" bgcolor="#DAECF5">&nbsp;</td>
    </tr>
  </table>
  <%
listnum=request("listnum")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from job"
mysql=mysql&" where id order by listnum asc"
mypage.getsql=mysql
mypage.pagesize=3
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%> 
 <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#C7D5E0">
    

    <tr bgcolor="#D7F9EC" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="0" colspan="3" align="right" ><%=i%>.应聘职位：</td>
      <td width="84%" align="left" ><%=rs("Title")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="1" colspan="3" align="right" >人数： </td>
      <td align="left" ><%=rs("renci")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="3" colspan="3" align="right" >工作地点：</td>
      <td align="left" ><%=rs("gzdd")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="7" colspan="3" align="right" >工资待遇： </td>
      <td align="left" ><%=rs("gzdy")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="15" colspan="3" align="right" >截止日期：</td>
      <td align="left" ><%=rs("jzrq")%></td>
    </tr>
    
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >发布时间：</td>
      <td align="left" ><%=rs("Times")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >招聘要求：</td>
      <td align="left" ><%if rs("content")<>"" then response.write cutstr(removehtml(rs("content")),100) else response.write "没有输入要求" end if%></td>
    </tr><tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="24" colspan="3" align="right" >显示排序：</td>
      <td align="left" ><input name="listt" type="text" class="form2" id="listt" value="<%=rs("listnum")%>" size="6" />
&nbsp;&nbsp;      <a href="?ss=edit&action=edit&amp;id=<%=rs("ID")%>&page=<%=page%>">编辑</a> | <a  onclick='{if(confirm("您确定删除吗?此操作将不能恢复!")){return true;}return false;}' href="?ac=del&amp;id=<%=rs("ID")%>&amp;page=<%=page%>">删除</a>&nbsp;&nbsp;&nbsp;&nbsp;选择
        <input type="checkbox" value="<%=rs("ID")%>" name="ID" id="chk" style="border:0;" />
        <input name="ID2" type="hidden" id="ID2" value="<%=rs("ID")%>" /></td>
    </tr>
 
  </table>
  <table width="98%" height="10" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td></td>
    </tr>
  </table> 
    <%
        rs.movenext
    else
         exit for
    end if
next
%>
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CFDCE4">
    <tr>
      <td align="center" bgcolor="#F9F9F9">全选/取消：
        <input name="Action" type="hidden"  value="Del" />
        <input name="chkAll" type="checkbox" id="chkAll" onclick=sel('chk') value="checkbox" style="border:0" />
        &nbsp;
        <input name="del" type="submit" class="admintable1" id="del" value="排序" />
&nbsp;&nbsp;&nbsp;
<input name="Del" type="submit" class="admintable" id="Del" value="更新时间" />
&nbsp;&nbsp;
<input name="Del" type="submit" class="admintable1"  onclick='{if(confirm("您确定删除吗?此操作将不能恢复!")){return true;}return false;}'  id="Del" value="删除" />
更新时间、批量删除必须先选取</td>
    </tr>
    <tr>
      <td bgcolor="#F9F9F9"><div id=page><ul style="text-align:center;">
        <%=mypage.showpage()%>
      </ul>  </div>    </td>
    </tr>
  </table>
  <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<%
	rs.close
end sub

%>
<%

Sub delAll
ID=Trim(Request("ID"))
ids = Request.Form("id2") '获取同名控件的值，如：1,2,3,4,5
listt = Request.Form("listt") '也是获取同名控件的值，如你上面的：9,10,11,12,13
listtTmp = Split(listt,",")
idTmp = Split(ids,",")

If ID="" and Request("Del")<>"排序"  Then
	  Response.Write("<script language=javascript>alert('您没有选择!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="更新时间" Then
   set rs=conn.execute("update job set times = now() where ID In(" & ID & ")")
   Response.Write("<script>alert('操作成功!');location='?ss=list'</script>")
ElseIf Request("Del")="排序" Then
For i = 0 To UBound(idTmp)
  ' conn.execute("update job set list ="& listtTmp(i) &" where ID In(" & ID & ")")
    conn.execute("update job set listnum=" & listtTmp(i) & " where id=" & idTmp(i))
Next
 Response.Write("<script>alert('操作成功!');location='?ss=list'</script>")

ElseIf Request("Del")="删除" Then
	'set rs=conn.execute("delete from news where ID In(" & ID & ")")
			for i=1 to request("ID").count
				if request("ID").count=1 then
				newsID=request("ID")
				else
				newsID=replace(request("id")(i),"'","")
				end if
				
				'删除文章
				Conn.Execute("Delete from [job] where ID = "&newsID&"")
				
			next
			response.write "<script>alert('删除成功!');location='?ss=list'</script>"
End If
End Sub
if trim(request.querystring("ac"))="del" then
ID=request("ID")
Conn.Execute("Delete from [job] where ID = "&ID&"")
response.write "<script>alert('删除成功!');location='?ss=list'</script>"
end if

%>

</div> 


<%end if%>









<%if ss="add" then%>






	<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_add(){
var notnull;
notnull=true;
if (document.form1.title.value==""){
alert("标题不能为空！");
document.form1.title.focus();
notnull=false;
}


return notnull;
}
</script>

<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>

<form id="form1" name="form1" method="post" action="?action=add" class="border-padd0" onSubmit="return check_add()" >
<table width="98%" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="13" bgcolor="#DAEDFC">&nbsp;</td>
    <td align="center" bgcolor="#DAEDFC" class="fontb">添加招聘信息</td>
    <td width="13" bgcolor="#DAEDFC">&nbsp;</td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#D7E3F2" class="border1">
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">招聘职位:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="title" type="text" class="form1" id="title" size="30" maxlength="50" /><input name="TColor" type="hidden" id="TitleFontColor" onClick="Getcolor(ColorBG,'TitleFontColor');" Readonly></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">招聘人数: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="renci" type="text" class="form1" id="renci" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">工作地点:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdd" type="text" class="form1" id="gzdd" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">工资待遇:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdy" type="text" class="form1" id="gzdy" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">截止日期: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="jzrq" type="text" class="form1" id="jzrq" size="30" maxlength="50"  />
      </td>
    </tr>
  
    

    
    <tr>
      <td align="right" valign="top" bgcolor="#F9F9F9">招聘要求：</td>
      <td colspan="2" bgcolor="#F9F9F9">
<textarea name="content" style="display:none"></textarea>
		<IFRAME ID="qi500" SRC="../qi500@lm_webe/qi500@edit.htm?id=content&style=blue" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="500"></IFRAME></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#F9F9F9">发布时间：</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="times" type="text" class="form1" id="times" value="<%=now%>" size="30" /></td>
    </tr>

    <tr>
      <td colspan="3" align="center" bgcolor="#F9F9F9"><input type="submit" name="Submit" value="确认添加" />
      &nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="Submit2" value="清空重写" /></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="30">&nbsp;</td>
    </tr>
  </table>
</form>

<%end if%>


<%

if trim(request.querystring("action"))="add" then
if trim(request.form("submit"))="确认添加" then

    Title		=trim(request.form("Title"))
	content			=request.form("content")
	renci	=request.form("renci")
	TColor	=request.form("TColor")
	gzdd			=trim(request.form("gzdd"))
	gzdy			=trim(request.form("gzdy"))
	jzrq			=trim(request.form("jzrq"))
	times		=trim(request.form("times"))
	%>
	
	<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from job"
	rs.open sql,conn,1,3
	rs.addnew
	
		rs("Title")				=title
		rs("times")			=times
		rs("TColor")	=TColor
		rs("renci")			=renci
        rs("gzdd")=gzdd
        rs("gzdy")=gzdy
		rs("jzrq")=jzrq
		rs("content")=content
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	
    response.write "<script>alert('添加成功');location='?ss=list'</script>"
end if
end if
%>






















<%if ss="edit" then%>




<%
id=request("id")
page=request("page")
set rs=server.createobject("adodb.recordset")	
sql="select * from  job  where id="&id
rs.open sql,conn,1,1
%>
<form id="form1" name="form1" method="post" action="?action=edit&amp;id=<%=rs("id")%>&amp;page="&page&"" class="border-padd0" onSubmit="return check_add()" >
<table width="98%" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="13">&nbsp;</td>
    <td align="center" background="../images/bn_bg.jpg" class="fontb">修改招聘信息</td>
    <td width="13">&nbsp;</td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#D7E3F2" class="border1">
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">招聘职位:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="title" type="text" class="form1" id="title" value="<%=rs("title")%>" size="30" maxlength="50" /><input name="TColor" type="hidden" id="TitleFontColor" onClick="Getcolor(ColorBG,'TitleFontColor');" Readonly></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">招聘人数: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="renci" type="text" class="form1" id="renci" value="<%=rs("renci")%>" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">工作地点:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdd" type="text" class="form1" id="gzdd" value="<%=rs("gzdd")%>" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">工资待遇:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdy" type="text" class="form1" id="gzdy" value="<%=rs("gzdy")%>" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">截止日期: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="jzrq" type="text" class="form1" id="jzrq" value="<%=rs("jzrq")%>" size="30" maxlength="50" /></td>
    </tr>
  
    

    
    <tr>
      <td align="right" valign="top" bgcolor="#F9F9F9">招聘要求：</td>
      <td colspan="2" bgcolor="#F9F9F9">
	  
	<textarea id="content" name="content" style="display:none;"><%=rs("content")%></textarea>
		<IFRAME ID="qi500" SRC="../qi500@lm_webe/qi500@edit.htm?id=content&style=blue" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="500"></IFRAME></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#F9F9F9">发布时间：</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="times" type="text" class="form1" id="times" value="<%=rs("times")%>" size="30" /></td>
    </tr>

    <tr>
      <td colspan="3" align="center" bgcolor="#F9F9F9"><input type="submit" name="Submit" value="确认修改" />
      &nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="Submit2" value="清空重写" /></td>
    </tr>
  </table>
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="30">&nbsp;</td>
    </tr>
  </table>
</form>
<%end if%>

<%

if trim(request.querystring("action"))="edit" then
if trim(request.form("submit"))="确认修改" then

    Title		=trim(request.form("Title"))
	content			=request.form("content")
	renci	=request.form("renci")
	TColor	=request.form("TColor")
	gzdd			=trim(request.form("gzdd"))
	gzdy			=trim(request.form("gzdy"))
	jzrq			=trim(request.form("jzrq"))
	times		=trim(request.form("times"))
	%>
	
	<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from job"
	rs.open sql,conn,1,3
		rs("Title")				=title
		rs("times")			=times
		rs("TColor")	=TColor
		rs("renci")			=renci
        rs("gzdd")=gzdd
        rs("gzdy")=gzdy
		rs("jzrq")=jzrq
		rs("content")=content
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	
    response.write "<script>alert('修改成功');location='?ss=list'</script>"
end if
end if
%>




























</body>
</html>
