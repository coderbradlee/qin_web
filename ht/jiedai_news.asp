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
	sql="select * from jiedai_News"
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
sql="select * from jiedai_newsclass order by flag asc"
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
          <tr style="display:none">
            <td align="center">��Ϣ��飺</td>
            <td><textarea name="jianjie" rows="4" id="jianjie" style="width:100%"></textarea></td>
          </tr>
          <tr style="display:none">
            <td align="center">��Ϣ��Դ��</td>
            <td><input name="xxfrom" type="text" id="imagebig" size="40" style="height:20;width:100%"></td>
          </tr>
          <tr style="display:none">
            <td align="center">��Ϣ���ߣ�</td>
            <td><input name="author" type="text" id="author" size="16" style="height:20;width:100%"></td>
          </tr>
          <tr>
            <td align="center">����ʱ�䣺</td>
            <td><input name="addtime" type="text" id="addtime" size="23" style="height:20;width:100%"  value="<%=now()%>"></td>
          </tr>
          <tr style="display:none">
            <td align="center">�ⲿ���ӣ�</td>
            <td><input name="wblink" type="text" id="wblink" size="40" style="height:20;width:100%"></td>
          </tr>
          <tr>
            <td align="center">������ɫ��</td>
            <td>
<select name=titlecolor size=1>
<option value="">������ɫ</option>
<option style="background-color:Black;color:Black" value=Black>�� ɫ</option>
<option style="background-color:Red;color:Red" value=Red>�� ɫ</option>
<option style="background-color:Yellow;color:Yellow" value=Yellow>�� ɫ</option>
<option style="background-color:Green;color:Green" value=Green>�� ɫ</option>
<option style="background-color:Orange;color:Orange" value=Orange>�� ɫ</option>
<option style="background-color:Purple;color:Purple" value=Purple>�� ɫ</option>
<option style="background-color:Blue;color:Blue" value=Blue>�� ɫ</option>
<option style="background-color:Brown;color:Brown" value=Brown>�� ɫ</option>
<option style="background-color:Teal;color:Teal" value=Teal>ī ��</option>
<option style="background-color:Navy;color:Navy" value=Navy>�� ��</option>
<option style="background-color:Maroon;color:Maroon" value=Maroon>�� ʯ</option>
<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">�� ��</option>
<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">�� ��</option>
<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">�� ��</option>
<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">�� ��</option>
<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">�� ��</option>
<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">ש ��</option>
<option style="background-color:#6495ED;color: #6495ED" value="#6495ED">�� ��</option>
<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">�� ��</option>
<option style="background-color:#FF1493;color: #FF1493" value="#FF1493">õ���</option>
<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">�� ��</option>
<option style="background-color:#FFD700;color: #FFD700" value="#FFD700">�� ��</option>
<option style="background-color:#DAA520;color: #DAA520" value="#DAA520">�� ��</option>
<option style="background-color:#808080;color: #808080" value="#808080">�� ��</option>
<option style="background-color:#778899;color: #778899" value="#778899">�� ��</option>
<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">�� ��</option>
</select></td>
          </tr>


<script language="javascript">

function xs(){

if (form1.xse.checked==true)
{
alle.style.display="block";
}
else{

alle.style.display="none";

}




}




</script>



          <tr >
            <td align="center">ͼƬ���ţ�</td>
            <td><input type="checkbox" name="xse" onClick="xs();" id="led" style="cursor:hand" value="1"></td>
          </tr>
		  
		  
		  
		  
		  
          <tr style="display:none" id="alle">
            <td align="center">��Ϣ��ַ��</td>
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
	sql="select * from jiedai_News where id="&id
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
	sql="select * from jiedai_News where id="&id
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
sql="select * from jiedai_newsclass order by flag asc"
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
   
          <tr style="display:none">
            <td align="center">��Ϣ��Դ��</td>
            <td><input name="xxfrom" type="text" id="xxfrom" size="40" style="height:20;width:100%" value="<%=rs("xxfrom")%>"></td>
          </tr>
		  
		      <tr style="display:none">
            <td align="center">Ӣ����Դ��</td>
            <td><input name="xxfrom" type="text" id="xxfrom" size="40" style="height:20;width:100%" value="<%=rs("e_xxfrom")%>"></td>
          </tr>
		  
          <tr style="display:none">
            <td align="center">�������ߣ�</td>
            <td><input name="author" type="text" id="author2" size="16" style="height:20;width:100%" value="<%=rs("author")%>"></td>
          </tr>
		  
		      <tr style="display:none">
            <td align="center">Ӣ�����ߣ�</td>
            <td><input name="author" type="text" id="author2" size="16" style="height:20;width:100%" value="<%=rs("e_author")%>"></td>
          </tr>
		  
		  
          <tr>
            <td align="center">����ʱ�䣺</td>
            <td><input name="addtime" type="text" id="addtime2" size="23" style="height:20;width:100%" value="<%=rs("addtime")%>"></td>
          </tr>
          <tr style="display:none">
            <td align="center">�ⲿ���ӣ�</td>
            <td><input name="wblink" type="text" id="wblink2" size="40" style="height:20;width:100%" value="<%=rs("wblink")%>"></td>
          </tr>
          <tr>
            <td align="center">������ɫ��</td>
            <td>
			
			
<select name=titlecolor size=1>
<option value="">ѡ����ɫ</option>
<option style="background-color:Black;color:Black" value=Black <%if rs("titlecolor")="Black" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Red;color:Red" value=Red <%if rs("titlecolor")="Red" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Yellow;color:Yellow" value=Yellow <%if rs("titlecolor")="Yellow" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Green;color:Green" value=Green <%if rs("titlecolor")="Green" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Orange;color:Orange" value=Orange <%if rs("titlecolor")="Orange" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Purple;color:Purple" value=Purple <%if rs("titlecolor")="Purple" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Blue;color:Blue" value=Blue <%if rs("titlecolor")="Blue" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Brown;color:Brown" value=Brown <%if rs("titlecolor")="Brown" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Teal;color:Teal" value=Teal <%if rs("titlecolor")="Teal" then response.write "selected" %>>ī ��</option>
<option style="background-color:Navy;color:Navy" value=Navy <%if rs("titlecolor")="Navy" then response.write "selected" %>>�� ��</option>
<option style="background-color:Maroon;color:Maroon" value=Maroon <%if rs("titlecolor")="Maroon" then response.write "selected" %>>�� ʯ</option>
<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF" <%if rs("titlecolor")="#00FFFF" then response.write "selected" %>>�� ��</option>
<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4" <%if rs("titlecolor")="#7FFFD4" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4" <%if rs("titlecolor")="#FFE4C4" then response.write "selected" %>>�� ��</option>
<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00" <%if rs("titlecolor")="#7FFF00" then response.write "selected" %>>�� ��</option>
<option style="background-color:#D2691E;color: #D2691E" value="#D2691E" <%if rs("titlecolor")="#D2691E" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50" <%if rs("titlecolor")="#FF7F50" then response.write "selected" %>>ש ��</option>
<option style="background-color:#6495ED;color: #6495ED" value="#6495ED" <%if rs("titlecolor")="#6495ED" then response.write "selected" %>>�� ��</option>
<option style="background-color:#DC143C;color: #DC143C" value="#DC143C" <%if rs("titlecolor")="#DC143C" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FF1493;color: #FF1493" value="#FF1493" <%if rs("titlecolor")="#FF1493" then response.write "selected" %>>õ���</option>
<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF" <%if rs("titlecolor")="#FF00FF" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FFD700;color: #FFD700" value="#FFD700" <%if rs("titlecolor")="#FFD700" then response.write "selected" %>>�� ��</option>
<option style="background-color:#DAA520;color: #DAA520" value="#DAA520" <%if rs("titlecolor")="#DAA520" then response.write "selected" %>>�� ��</option>
<option style="background-color:#808080;color: #808080" value="#808080" <%if rs("titlecolor")="#808080" then response.write "selected" %>>�� ��</option>
<option style="background-color:#778899;color: #778899" value="#778899" <%if rs("titlecolor")="#778899" then response.write "selected" %>>�� ��</option>
<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE" <%if rs("titlecolor")="#B0C4DE" then response.write "selected" %>>�� ��</option>
</select>			</td>
          </tr>
		  
		  
		  
<script language="javascript">

function xs(){

if (form1.xse.checked==true)
{
alle.style.display="block";
}
else{

alle.style.display="none";

}

}




</script>

		  
          <tr >
            <td align="center">ͼƬ���ţ�</td>
            <td><input type="checkbox" name="xse" onClick="xs();" id="led" style="cursor:hand" value="1" <%if rs("tpcheck")="1" then response.write"checked"%>></td>
          </tr>
		  
		  <%if rs("tpcheck")="1" then%>
		  
          <tr id="alle" style="display:none">
		  
		  <%else%>
		  
          <tr id="alle" style="display:none">
		  
		  <%end if%>
		  
		  
            <td align="center">��Ϣ��ַ��</td>
            <td><table width="50%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>
				  
				  <%if rs("tpcheck")="1" then%>
				  
				  <input name="image" type="text" id="image" size="40" style="height:20" value="<%=rs("tupian")%>">
				  
				  <%else%>
				  
				  <input name="image" type="text" id="image" size="40" style="height:20">
				  
				  <%end if%>
				  
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
	sql="select * from jiedai_News where id="&id
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
	sql="select * from jiedai_News where id="&id
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
	sql="select * from jiedai_News where id="&id
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
	sql="select * from jiedai_News where id="&id
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
	sql="select * from jiedai_News where id="&id
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
sql="select * from jiedai_newsclass order by flag asc"
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
             <td align="center" bgcolor="#DFEFFF" style="font-size:16px"><b>��������</b></td>
           </tr>
         </table>
		 
		 
		 
		 
		 
		 
		 
		 	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="7"></td>
                </tr>
              </table>
		 	  <% 	  
sql_classid="select * from jiedai_newsclass"
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
         <td bgcolor="#DFEFFF" align="left">�������� </td>
         <td width="50" bgcolor="#DFEFFF">ͷ��</td>
         <td width="50" bgcolor="#DFEFFF">�ö�</td>
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
		
	sql="select * from jiedai_News where 1=1 "
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
           <td width="50" align="center" class="line"><%if rs("toutiao")=1 then%>
               <a href="?Action=list&toutiao=ttyes&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/Ok.gif" alt="������ͷ��" width="16" height="16" border="0" /></a>
               <%else%>
               <a href="?Action=list&toutiao=ttno&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/err.gif" alt="δ����ͷ��" width="12" height="11" border="0" /></a>
               <%end if%></td>
           <td width="50" align="center" class="line"><%if rs("tuijian")=1 then%>
               <a href="?Action=list&zhiding=zdyes&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/Ok.gif" alt="���ö�" width="16" height="16" border="0" /></a>
               <%else%>
               <a href="?Action=list&zhiding=zdno&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>"><img src="images/err.gif" alt="δ�ö�" width="12" height="11" border="0" /></a>
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