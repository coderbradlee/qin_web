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
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td><%
if trim(request.querystring("action"))="add" then
if trim(request.form("submit"))="���" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from jd_case"
	rs.open sql,conn,3,2
	rs.addnew
	rs("title")=request.form("title")
	rs("classid")=request.form("newsclass")
	rs("xxfrom")=request.form("xxfrom")
	
	rs("tpcheck")=request.form("xse")
	rs("tupian")=request.form("image")


	
	rs("author")=request.form("author")
	rs("jianjie")=request.form("jianjie")
	rs("addtime")=request.form("addtime")
	rs("wblink")=request.form("wblink")
	rs("titlecolor")=request.form("titlecolor")
	rs("content")=request.form("content")
	
	rs("e_title")=request.form("e_title")
	rs("e_content")=request.form("e_content")
	
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
	if (document.form1.title.value==""){
		alert("��Ϣ���Ʋ���Ϊ�գ�");
		document.form1.title.focus();
		
		    return false;
		
		
		}
		
	if (form1.xse.checked==true)
	{
		if (document.form1.image.value==""){
		
		alert("ͼƬ����������,���ϴ�ͼƬ��");
		document.form1.image.focus();
				    return false;

		}	
	}
	
	return true;
	
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">�����Ϣ</font></strong></td>
        </tr>
      </table>
      <br>
      <form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
           <tr>
            <td width="100" align="right">�������ƣ�</td>
            <td><input name="title" type="text" id="title" size="40" style="height:20; width:100%"></td>
          </tr>
          
          
		    <tr>
            <td width="100" align="right">Ӣ�����ƣ�</td>
            <td><input name="e_title" type="text" id="e_title" size="40" style="height:20; width:100%"></td>
          </tr>
          <tr>
            <td align="center">��Ϣ���ͣ�</td>
            <td>
<% 	  
sql="select * from jd_caseclass order by flag asc"
set rs=conn.execute(sql)  
 %>
                  <select name="newsclass" id="newsclass" style="width:100%">
                    <% do while not rs.eof %>
                    <option value="<%= rs("id") %>" <%if rs("id")=classid then response.Write("selected='selected'")%>><%= rs("classname") %>  [[<%= rs("e_classname") %>]]</option>
                    <%
		 rs.movenext
		loop
		rs.close
		set rs=nothing
		 %>
                  </select>            </td>
          </tr>
    
  
          <tr>
            <td align="center">����ʱ�䣺</td>
            <td><input name="addtime" type="text" id="addtime" size="23" style="height:20;width:100%"  value="<%=now()%>"></td>
          </tr>
          <tr>
            <td align="center">�ⲿ���ӣ�</td>
            <td><input name="wblink" type="text" id="wblink" size="40" style="height:20;width:100%"></td>
          </tr>
  
		  
		  
		  
		  
          <tr >
            <td align="center">ͼƬ��</td>
            <td><table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><input name="image" type="text" id="image" size="40" style="height:20"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
          

		  <tr>
            <td align="center">��������</td>
            <td>
			 <textarea name="content" cols="" rows="" style="display:none"></textarea><iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		  
		  	  
		  <tr>
            <td align="center">Ӣ������</td>
            <td>
			 <textarea name="e_content" cols="" rows="" style="display:none"></textarea><iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
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
	dim rs,sql
	set rs=server.createobject("adodb.recordset")
	sql="select * from jd_case where id="&id
	rs.open sql,conn,1,3
	rs("title")=request.form("title")
	rs("classid")=request.form("newsclass")
	rs("xxfrom")=request.form("xxfrom")
	rs("author")=request.form("author")
	rs("addtime")=request.form("addtime")
		rs("tpcheck")=request.form("xse")
	rs("tupian")=request.form("image")


	rs("wblink")=request.form("wblink")
	rs("jianjie")=request.form("jianjie")
	rs("titlecolor")=request.form("titlecolor")
	rs("content")=request.form("content")
	
	rs("e_title")=request.form("e_title")
	rs("e_content")=request.form("e_content")
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�޸ĳɹ�!');location='?action=list'</script>"
	response.end
end if
id=trim(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
	sql="select * from jd_case where id="&id
	rs.open sql,conn,1,1
%>
      <script language="javascript" type="text/javascript">
// ��֤�û���������
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.title.value==""){
		alert("���ⲻ��Ϊ�գ�");
		document.form1.title.focus();
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
      <form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
         
		 
		 
		
		  
		      <tr>
            <td width="100" align="right">�������ƣ�</td>
            <td><input name="title" type="text" id="title2" size="40" style="height:20; width:100%" value="<%=rs("title")%>"></td>
          </tr>
          
             <tr>
            <td width="100" align="right">Ӣ�����ƣ�</td>
            <td><input name="e_title" type="text" id="e_title" size="40" style="height:20; width:100%" value="<%=rs("e_title")%>"></td>
          </tr>  
		  
		  
		  
		  
		  
		  
          <tr>
            <td align="center">��Ϣ���ͣ�</td>
            <td><% 	  
sql="select * from jd_caseclass order by flag asc"
set ras=conn.execute(sql)  
 %>
                <select name="newsclass" id="select" style="width:100%">
                  <% do while not ras.eof %>
                  <option value="<%= ras("id") %>" <%if ras("id")=rs("classid") then response.Write("selected='selected'")%>><%= ras("classname") %> [[<%= ras("e_classname") %>]]</option>
                  <%
		 ras.movenext
		loop
		ras.close
		set ras=nothing
		 %>
                </select>            </td>
          </tr>
   
		  
          <tr>
            <td align="center">����ʱ�䣺</td>
            <td><input name="addtime" type="text" id="addtime2" size="23" style="height:20;width:100%" value="<%=rs("addtime")%>"></td>
          </tr>
          <tr >
            <td align="center">�ⲿ���ӣ�</td>
            <td><input name="wblink" type="text" id="wblink2" size="40" style="height:20;width:100%" value="<%=rs("wblink")%>"></td>
          </tr>
         

		  

		  
          <tr>
	
		  
		  
            <td align="center">ͼƬ��</td>
            <td><table width="50%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>
				  
				
				  <input name="image" type="text" id="image" size="40" style="height:20" value="<%=rs("tupian")%>">
			
				  
				  </td>
                  <td style="padding-left:8px"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                </tr>
            </table></td>
          </tr>
          
		  
		  
		  
		  
		  
		  
		  
		  	  <tr>
            <td align="center">��������</td>
            <td>
			 <textarea name="content" cols="" rows="" style="display:none"><%=rs("content")%></textarea><iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		  
		  	  
		  <tr>
            <td align="center">Ӣ������</td>
            <td>
			 <textarea name="e_content" cols="" rows="" style="display:none"><%=rs("e_content")%></textarea><iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		    
		  
		  
		  
		  
		  
		  
		  
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="�޸�" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit22" value="����" style="width:80; height:30; cursor:hand"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <% 
if trim(request.querystring("action"))="del" then
	id=trim(request.querystring("id"))
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from jd_case where id="&id
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
if trim(request.querystring("zhiding"))="zdyes" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jd_case where id="&id
	rs.open sql,conn,1,3
	rs("tuijian")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('��ȡ���ö�!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
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
	sql="select * from jd_case where id="&id
	rs.open sql,conn,3,2
	rs("tuijian")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�ö��ɹ�!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	response.end
end if
 %>
 
 
 
      <% 
if trim(request.querystring("toutiao"))="ttyes" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jd_case where id="&id
	rs.open sql,conn,3,2
	rs("toutiao")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('ͷ����Ϣ��ȡ��!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	response.end
end if
 %>
 
 
      <% 
if trim(request.querystring("toutiao"))="ttno" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jd_case where id="&id
	rs.open sql,conn,3,2
	rs("toutiao")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('ͷ�����óɹ�!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	response.end
end if
 %>
 
 
 
 
 
 
 
 
 
 
<%
if trim(request.querystring("action"))="list" then
 %>
 
 
 
 
 
  
 
 
 
 
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:7px">
  <form name="form" method="post" action="?Action=list">
    <tr> 
      <td align="left">����Ϣ������
<% 	  
sql="select * from jd_caseclass order by flag asc"
set ras=conn.execute(sql)  
 %>
                  <select name="newsclass" id="newsclass" style="width:130px">
				  <option value="0">��������</option>
				  
                    <% do while not ras.eof %>
                    <option value="<%= ras("id") %>" <%if request("newsclass")=ras("id") then response.Write("selected")%>><%= ras("classname") %></option>
                    <%
		 ras.movenext
		loop
		ras.close
		set ras=nothing
		 %>
                  </select>      			
        <input name="keywords" type="text" class="input" id="keywords" style="width:150px;height:21px; padding-left:5px" onFocus='this.select()' onBlur="if (this.value ==''){this.value=this.defaultValue}" onClick="if(this.value=='������Ϣ�ؼ���')this.value=''" value="������Ϣ�ؼ���">
	  <input name="Submit" type="submit" class="bt" id="Submit" value="����">
      </td>
      <td align="right">&nbsp;</td>
    </tr>
  </form>
</table>
 
 
 
 
 
 <table width="100%" border="0" cellspacing="0" cellpadding="6">
   <tr>
     <td width="18%" valign="top"><table width="100%" height="107" border="0" cellpadding="8" cellspacing="1" bgcolor="#DFEFFF">
       <tr>
         <td align="center" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="5">
           <tr>
             <td align="center" bgcolor="#DFEFFF" style="font-size:16px"><b>��Ϣ����</b></td>
           </tr>
         </table>
		 
		 
		 
		 
		 
		 
		 
		 	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="7"></td>
                </tr>
              </table>
		 	  <% 	  
sql_classid="select * from jd_caseclass"
set rs_classid=conn.execute(sql_classid)  
 %>	  
	  <% do while not rs_classid.eof %>
        <a href="?action=list&cid=<%= rs_classid("id") %>"><%if request.QueryString("cid")=rs_classid("id") then response.write"<font color=#ff0000><b>"%><%= rs_classid("classname") %></font></a>   <%if rs_classid("e_classname")<>"" then 
		   response.Write"<img src=""images/en.jpg"" />" 
		   else
		   response.Write"<img src=""images/noen.jpg"" />" 
		   end if
		   %><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="7"></td>
  </tr>
</table>

		<%
		 rs_classid.movenext
		loop
		rs_classid.close
		 %>		 </td>
       </tr>
     </table></td>
     <td width="82%" valign="top"><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
       <tr align="center">
         <td width="50" align="center" bgcolor="#DFEFFF">���</td>
         <td align="left" bgcolor="#DFEFFF">��Ϣ���� </td>
         <td width="50" align="center" bgcolor="#DFEFFF">ɾ��</td>
         <td width="50" align="center" bgcolor="#DFEFFF">�༭</td>
       </tr>
     </table>
       <br>
       <%
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   	newsclass=request.form("newsclass")
	keywords=request.form("keywords")
	cid=request.querystring("cid")
	set rs=server.createobject("adodb.recordset")
		
	sql="select * from jd_case where 1=1 "
	if newsclass<>"0" and newsclass<>"" then
	sql=sql+" and classid="&newsclass&" and title like '%"&keywords&"%' or content like '%"&keywords&"%' "
	elseif newsclass="0" then
	sql=sql+" and title like '%"&keywords&"%' or content  like '%"&keywords&"%' "
	end if
	if cid<>"" then
	sql=sql+" and classid="&cid&" "
	end if
	sql=sql+" order by tuijian desc,id desc"
		
		
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
           <td class="line">&nbsp;<font color="<%=rs("titlecolor")%>"><%= rs("title") %></font>   <%if rs("e_title")<>"" then 
		   response.Write"<img src=""images/en.jpg"" />" 
		   else
		   response.Write"<img src=""images/noen.jpg"" />" 
		   end if
		   %></td>
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
else
response.write"<div align=center><br>������Ϣ<br><br></div>"
end if
%></td>
   </tr>
 </table> 
 <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
  <tr>
    <td align="center">��<%= page %>ҳ&nbsp;
        <% if page<>1 then %>
        <a href="?action=list&page=1&cid=<%= cid %>">��ҳ</a>
        <% else %>
      ��ҳ
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>&cid=<%= cid %>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>&cid=<%= cid %>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.recordcount then %>
      <a href="?action=list&page=<%= rs.recordcount %>&cid=<%= cid %>">ĩҳ</a>
      <% else %>
      ĩҳ
      <% end if %>
      &nbsp;����<%= rs.recordcount %>��</td>
  </tr>
</table>
<% end if %></td>
  </tr>
</table>
</body>
</html>