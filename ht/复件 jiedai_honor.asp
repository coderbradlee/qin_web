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
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td><%
if trim(request.querystring("action"))="add" then
if trim(request.form("submit"))="添加" then
	
   weburl=request.form("wblink")
  if weburl<>"" then
        if left(weburl,7)<>"http://" then
         response.write "<script>alert('链接格式错误，必须以 HTTP:// 开头，此项也可以填写!');location='?action=list'</script>"
        response.end
        end if
  end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_honor"
	rs.open sql,conn,1,3
	rs.addnew
	rs("title")=request.form("title")
rs("e_title")=request.form("e_title")
rs("e_content")=request.form("e_content")
	rs("addtime")=request.form("addtime")
		rs("tuijian")=request.form("xse")
	rs("tupian")=request.form("image")
rs("content")=request.form("content")
	rs("wblink")=weburl

	rs.update

	'rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('添加成功!');location='?action=list'</script>"
end if
%>
      <script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_add(){
	
		
	if (form1.xse.checked==true)
	{
		if (document.form1.image.value==""){
		
		alert("图片已设置显示,请上传图片！");
		document.form1.image.focus();
				    return false;

		}	
	}
	
	
	if (document.form1.image.value==""){
		alert("请上传图片");
		document.form1.image.focus();
		
		    return false;
		
		
		}
	return true;
	
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">添加信息</font></strong></td>
        </tr>
      </table>
      <br>
      <form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td width="100" align="right">中文名称：</td>
            <td><input name="title" type="text" id="title" size="40" style="height:20; width:100%"></td>
          </tr>
          
          
		    <tr>
            <td width="100" align="right">英文名称：</td>
            <td><input name="e_title" type="text" id="e_title" size="40" style="height:20; width:100%"></td>
          </tr>
          
		  
          
          
          <tr>
            <td align="right">发布时间：</td>
            <td><input name="addtime" type="text" id="addtime" size="23" style="height:20;width:100%"  value="<%=now()%>"></td>
          </tr>
          <!--<tr >
            <td align="right">链接：</td>
            <td><input name="wblink" type="text" id="wblink" style="height:20;width:100%" value="http://" size="40"></td>
          </tr> -->
          


          <tr >
            <td align="right">首页显示：</td>
            <td><input type="checkbox" name="xse"  id="led" style="cursor:hand" value="1"></td>
          </tr>
		  
		  
		  
		  
		  
          <tr  id="alle">
            <td align="right">上传图片：</td>
            <td><table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><input name="image" type="text" id="image" size="40" style="height:20"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
       
          
		  
		  
		  
		  
		  
		  <tr>
            <td align="center">中文内容</td>
            <td>
			 <textarea name="content" cols="" rows="" style="display:none"></textarea><iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		  
		  	  
		  <tr>
            <td align="center">英文内容</td>
            <td>
			 <textarea name="e_content" cols="" rows="" style="display:none"></textarea><iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		  
		  
		  
		  
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="添加" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit2" value="重置" style="width:80; height:30; cursor:hand"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <%
if trim(request.querystring("action"))="edit" then
if trim(request.form("submit"))="修改" then
	id=trim(request.querystring("id"))
	
	   weburl=request.form("wblink")
  if weburl<>"" then
        if left(weburl,7)<>"http://" then
         response.write "<script>alert('链接格式错误，必须以 HTTP:// 开头，此项也可以填写!');location='?action=list'</script>"
         response.end
        end if
  end if
  
	dim rs,sql
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_honor where id="&id
	rs.open sql,conn,1,3
	rs("title")=request.form("title")
	'rs("classid")=request.form("newsclass")
	'rs("xxfrom")=request.form("xxfrom")
	'rs("author")=request.form("author")
	rs("e_title")=request.form("e_title")
rs("e_content")=request.form("e_content")
	rs("addtime")=request.form("addtime")
		rs("tuijian")=request.form("xse")
	rs("tupian")=request.form("image")


	rs("wblink")=weburl
	'rs("jianjie")=request.form("jianjie")
	'rs("titlecolor")=request.form("titlecolor")
	rs("content")=request.form("content")
	rs.update
	rs.requery
	rs.close
	set rs=nothing

	response.write "<script>alert('修改成功!');location='?action=list'</script>"
	response.end
end if
id=trim(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_honor where id="&id
	rs.open sql,conn,1,1
%>
      <script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.image.value==""){
		alert("图片不能为空！");
		document.form1.image.focus();
		notnull=false;
		}
	return notnull;
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">修改信息</font></strong></td>
        </tr>
      </table>
      <br>
      <form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td width="100" align="right">中文名称：</td>
            <td><input name="title" type="text" id="title2" size="40" style="height:20; width:100%" value="<%=rs("title")%>"></td>
          </tr>
          
             <tr>
            <td width="100" align="right">英文名称：</td>
            <td><input name="e_title" type="text" id="e_title" size="40" style="height:20; width:100%" value="<%=rs("e_title")%>"></td>
          </tr>
          
		  
          
          
          <tr>
            <td align="right">发布时间：</td>
            <td><input name="addtime" type="text" id="addtime2" size="23" style="height:20;width:100%" value="<%=rs("addtime")%>"></td>
          </tr>
          <!--<tr >
            <td align="right">链接：</td>
            <td><input name="wblink" type="text" id="wblink2" size="40" style="height:20;width:100%" value="<%=rs("wblink")%>"></td>
          </tr> -->
          
		  
	



<tr align="right">

		  
     
		  
		
		  
            <td align="right">上传图片：</td>
            <td align="left"><table width="50%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>
				  
			
				  
				  <input name="image" type="text" id="image" size="40" style="height:20" value="<%=rs("tupian")%>">
			
				  
				  </td>
                  <td style="padding-left:8px"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                </tr>
            </table></td>
          </tr>
		  
	
		  
		  
		  	  <tr>
            <td align="center">中文内容</td>
            <td>
			 <textarea name="content" cols="" rows="" style="display:none"><%=rs("content")%></textarea><iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		  
		  	  
		  <tr>
            <td align="center">英文内容</td>
            <td>
			 <textarea name="e_content" cols="" rows="" style="display:none"><%=rs("e_content")%></textarea><iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		  
		  
		  
		  
		  
          
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="修改" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit22" value="重置" style="width:80; height:30; cursor:hand"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <% 
if trim(request.querystring("action"))="del" then
	id=trim(request.querystring("id"))
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_honor where id="&id
	rs.open sql,conn,1,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('删除成功!');location='?action=list'</script>"
end if
 %>
 
 
      <% 
if trim(request.querystring("zhiding"))="zdyes" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_honor where id="&id
	rs.open sql,conn,1,3
	rs("tuijian")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('已取消首页显示!');location='?Action=list&ynpage="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
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
	sql="select * from jiedai_honor where id="&id
	rs.open sql,conn,3,2
	
	if rs("tupian")="" then
		response.write "<script>alert('失败了，此条没有添加图片，不能设置!');location='?Action=list&ynpage="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	end if
	rs("tuijian")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('首页显示成功!');location='?Action=list&ynpage="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	response.end
end if
 %>
 
 
 
 
 
 
<%
if trim(request.querystring("action"))="list"  or trim(request.querystring("action"))="" then
 %>
 
 
 

 
 
 
 <table width="100%" border="0" cellspacing="0" cellpadding="6">
   <tr>
     
     <td width="82%" valign="top"><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
       <tr align="center">
         <td align="center" bgcolor="#DFEFFF"><strong>荣誉资质</strong></td>
         </tr>
     </table>
       <br>
     
 
 
 <% 

	on error resume next

set rs=server.createobject("adodb.recordset")

	sql="select * from jiedai_honor order by id desc "

	rs.open sql,conn,1,1
	rs.pagesize=8
	
	if trim(request.querystring("ynpage")<>"") then 
	
	
	
			if isnumeric(trim(request.querystring("ynpage")))=false then
			page=1
			else
			page=cint(trim(request.querystring("ynpage")))
			end if
	
	else
	
	page=1
	
	end if

	
	if page<1 then
		page=1
	elseif page>rs.pagecount then
		page=rs.pagecount
	end if
	rs.absolutepage=page
	
if rs.bof and rs.eof then
 response.write "<br><br>暂无产品"
 response.write "<br><br>"
end if
rowcount = rs.pagesize
do while not rs.eof and rowcount>0
%>
                      <table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td ><table width="181" border="0" cellpadding="0" cellspacing="3" class="zw">
                            <tr>
                              <%
i=0
do while i<4 and not rs.eof  and rowcount>0
%>






                              <td width="175">
							  
							  
							  
                              
                              
                              
                              
                              
                     <table width="162" border="0" cellpadding="0" cellspacing="0" style="margin:3px 5px;">
         <tr>
           <td bgcolor="#D3E9FC"><img src="../uploadfile/<%=rs("tupian")%>" height="155" width="162" border="0"  style="padding:1px; border:1px solid #B6D7EF; "></td>
         </tr>
         <tr>
           <td height="28" bgcolor="#D3E9FC" style="padding-left:5px;"><span class="line"><%if rs("e_title")<>"" then 
		   response.Write"<img src=""images/en.jpg"" />" 
		   else
		   response.Write"<img src=""images/noen.jpg"" />" 
		   end if
		   %> <%= rs("title") %> </span></td>
         </tr>
         <tr>
           <td height="28" align="center" bgcolor="#F3F3F3">&nbsp;&nbsp; <a href="?action=del&id=<%= rs("id") %>" onClick="return window.confirm('确定删除吗?');">[删除]</a> &nbsp;&nbsp; <a href="?action=edit&id=<%= rs("id") %>">[编辑]</a> &nbsp;&nbsp;  &nbsp;
             <%if rs("tuijian")=1 then%>
             <a href="?Action=list&zhiding=zdyes&jid=<%=rs("id")%>&ynpage=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/Ok.gif" alt="已显示" width="16" height="16" border="0" /></a>
             <%else%>
             <a href="?Action=list&zhiding=zdno&jid=<%=rs("id")%>&ynpage=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/err.gif" alt="未显示" width="12" height="11" border="0" /></a>
             <%end if%></td>
         </tr>
       </table>         
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
                              
							  
							  
							  </td>
<%
    i=i+1
	rowcount=rowcount-1
    rs.movenext
    loop
%>
                            </tr>
                            </table></td>
                        </tr>
                      </table>
                      <%
loop
%>
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 </td>
   </tr>
 </table> 
 <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
  <tr>
    <td align="center">第<%= page %>页&nbsp;
        <% if page<>1 then %>
        <a href="?action=list&ynpage=1&cid=<%= cid %>">首页</a>
        <% else %>
      首页
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&ynpage=<%= page-1 %>&cid=<%= cid %>">上一页</a>
      <% else %>
      上一页
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&ynpage=<%= page+1 %>&cid=<%= cid %>">下一页</a>
      <% else %>
      下一页
      <% end if %>
      &nbsp;
      <% if page<rs.recordcount then %>
      <a href="?action=list&ynpage=<%= rs.recordcount %>&cid=<%= cid %>">末页</a>
      <% else %>
      末页
      <% end if %>
      &nbsp;总数<%= rs.recordcount %>条</td>
  </tr>
</table>
<% end if  
%></td>
  </tr>
</table>
</body>
</html>