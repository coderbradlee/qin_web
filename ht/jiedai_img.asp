
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
	sql="select * from jiedai_img"
	rs.open sql,conn,1,3
	rs.addnew
	rs("title")=request.form("title")

	rs("addtime")=request.form("addtime")
		rs("tuijian")=request.form("xse")
	rs("tupian")=request.form("image")
rs("author")=request.form("author")
	rs("content")=request.form("content")
rs("toutiao")=request.form("toutiao")
	rs.update

	'rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('��ӳɹ�!');location='?'</script>"
end if
%>
      <script language="javascript" type="text/javascript">
// ��֤�û���������
function check_add(){
	
		
	if (form1.xse.checked==true)
	{
		if (document.form1.image.value==""){
		
		alert("ͼƬ��������ʾ,���ϴ�ͼƬ��");
		document.form1.image.focus();
				    return false;

		}	
	}
	
	
	if (document.form1.image.value==""){
		alert("���ϴ�ͼƬ");
		document.form1.image.focus();
		
		    return false;
		
		
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
           <td width="100" align="right">���ƣ�</td>
            <td><input name="title" type="text" id="title" size="40" style="height:20; width:100%"></td>
          </tr>
          
       <!--     <tr>
            <td width="100" align="right">���</td>
            <td><select name="toutiao" id="toutiao">
              <option value="1" selected>��˾�쵼</option>
              <option value="2">ְ�ܲ���</option>
            </select>
            </td>
          </tr>
		  
          <tr>
            <td width="100" align="right">ְλ��</td>
            <td><input name="author" type="text" id="author" size="40" style="height:20; width:100%"></td>
          </tr><tr>
            <td width="100" align="right">������</td>
            <td><textarea name="content" id="content"  style="height:90; width:100%"></textarea></td>
          </tr>
          
          <tr>
            <td align="right">����ʱ�䣺</td>
            <td><input name="addtime" type="text" id="addtime" size="23" style="height:20;width:100%"  value="<%'=now()%>"></td>
          </tr>
          <tr  style="display:none">
            <td align="right">���ӣ�</td>
            <td><input name="wblink" type="text" id="wblink" style="height:20;width:100%" value="http://" size="40"></td> 
          </tr>
          


          <tr >
            <td align="right">��ʾ��</td>
            <td><input type="checkbox" name="xse"  id="led" style="cursor:hand" value="1"></td>
          </tr>
		  -->
		  
		  
		  
		  
          <tr  id="alle">
            <td align="right">�ϴ�ͼƬ��</td>
            <td><table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><input name="image" type="text" id="image" size="40" style="height:20"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
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
 
	dim rs,sql
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_img where id="&id
	rs.open sql,conn,1,3
	'rs("title")=request.form("title")
	'rs("classid")=request.form("newsclass")
	'rs("xxfrom")=request.form("xxfrom")
	rs("author")=request.form("author")
	rs("addtime")=request.form("addtime")
		rs("tuijian")=request.form("xse")
	rs("tupian")=request.form("image")
rs("toutiao")=request.form("toutiao")

	'rs("wblink")=weburl
	'rs("jianjie")=request.form("jianjie")
	'rs("titlecolor")=request.form("titlecolor")
	rs("content")=request.form("content")
	rs.update
	rs.requery
	rs.close
	set rs=nothing

	response.write "<script>alert('�޸ĳɹ�!');location='?'</script>"
	response.end
end if
id=trim(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_img where id="&id
	rs.open sql,conn,1,1
%>
      <script language="javascript" type="text/javascript">
// ��֤�û���������
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.image.value==""){
		alert("ͼƬ����Ϊ�գ�");
		document.form1.image.focus();
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
           <td width="100" align="right">���ƣ�</td>
            <td><input name="title" type="text" disabled id="title2" style="height:20; width:100%" value="<%=rs("title")%>" size="40" readonly="readonly"></td>
          </tr>
         <!--    <tr>
            <td width="100" align="right">���</td>
            <td><select name="toutiao" id="toutiao">
              <option value="1" <%'if rs("toutiao")=1 then%>selected<%'end if%>>��˾�쵼</option>
              <option value="2" <%'if rs("toutiao")=2 then%>selected<%'end if%>>ְ�ܲ���</option>
            </select>
            </td>
          </tr>
          
         <tr>
            <td width="100" align="right">ְλ��</td>
            <td><input name="author" type="text" id="author" style="height:20; width:100%" value="<%'=rs("author")%>" size="40"></td>
          </tr><tr>
            <td width="100" align="right">������</td>
            <td><textarea name="content" id="content"  style="height:90; width:100%"><%'=rs("content")%></textarea></td>
          </tr> 
		  
		  
		  
		  
		  
		  
          
          <tr>
            <td align="right">����ʱ�䣺</td>
            <td><input name="addtime" type="text" id="addtime2" size="23" style="height:20;width:100%" value="<%'=rs("addtime")%>"></td>
          </tr>
          <tr  style="display:none">
            <td align="right">���ӣ�</td>
            <td><input name="wblink" type="text" id="wblink2" size="40" style="height:20;width:100%" value="<%'=rs("wblink")%>"></td> 
          </tr>
        
		  
	



<tr align="right">

		  
          <tr >
            <td align="right">��ʾ��</td>
            <td><input name="xse" type="checkbox" id="led" style="cursor:hand"  value="1" <%if rs("tuijian")="1" then response.write"checked"%>></td>
          </tr>
		  
		 --> 
		  
            <td align="right">�ϴ�ͼƬ��</td>
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
	sql="select * from jiedai_img where id="&id
	rs.open sql,conn,1,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('ɾ���ɹ�!');location='?'</script>"
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
	sql="select * from jiedai_img where id="&id
	rs.open sql,conn,1,3
	rs("tuijian")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('��ȡ����ʾ!');location='?Action=list&ynpage="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
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
	sql="select * from jiedai_img where id="&id
	rs.open sql,conn,3,2
	rs("tuijian")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�ö���ʾ�ɹ�!');location='?Action=list&ynpage="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
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
         <td align="left" bgcolor="#DFEFFF"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ͼƬ�б�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></td>
         </tr>
     </table>
       <br>
       
       
       
       
       
     <% 

	on error resume next
   toutiao=request("t")
set rs=server.createobject("adodb.recordset")
    if toutiao<>"" then
	sql="select * from jiedai_img where toutiao="&toutiao
	else
	sql="select * from jiedai_img order by id desc "
    end if
	rs.open sql,conn,1,1
	rs.pagesize=99999999
	
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
 response.write "<br><br>����"
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
do while i<2 and not rs.eof  and rowcount>0
%>






                              <td width="175">
							  
							  
							  
                              
                              
                              
                              
                              
                     <table width="162" border="0" cellpadding="0" cellspacing="0" style="margin:3px 5px;">
         <tr>
           <td colspan="2" bgcolor="#D3E9FC"><img src="../uploadfile/<%=rs("tupian")%>" height="164" width="354" border="0"  style="padding:1px; border:1px solid #B6D7EF; "></td>
         </tr>
         <tr >
           <td height="28" bgcolor="#D3E9FC" style="padding-left:5px;"><span class="line"><%= rs("id") %>. <%= rs("title") %></span>    </td>
           <td bgcolor="#D3E9FC" style="padding-left:5px;"><a href="?action=edit&id=<%= rs("id") %>">[�༭]</a></td>
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
    <td align="center">��<%= page %>ҳ&nbsp;
        <% if page<>1 then %>
        <a href="?action=list&ynpage=1&cid=<%= cid %>">��ҳ</a>
        <% else %>
      ��ҳ
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&ynpage=<%= page-1 %>&cid=<%= cid %>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&ynpage=<%= page+1 %>&cid=<%= cid %>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.recordcount then %>
      <a href="?action=list&ynpage=<%= rs.recordcount %>&cid=<%= cid %>">ĩҳ</a>
      <% else %>
      ĩҳ
      <% end if %>
      &nbsp;����<%= rs.recordcount %>��</td>
  </tr>
</table>
<% end if  
%></td>
  </tr>
</table>
</body>
</html>