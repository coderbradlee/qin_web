<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<%
sjid=request("sjid")

%>
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
	sql="select * from jiedai_paijiexi"
	rs.open sql,conn,3,2
	rs.addnew
	rs("jtype")=request.form("jtype")
	rs("diqu")=request.form("diqu")
	rs("jgname")=request.form("jgname")
	rs("jgdz")=request.form("jgdz")
	rs("jgyb")=request.form("jgyb")
	rs("jgtel")=request.form("jgtel")
	rs("yysj")=request.form("yysj")
	rs("zjdcz")=request.form("zjdcz")
	rs("zbxx")=request.form("zbxx")
	rs("pjys")=request.form("pjys")
	rs("jcolor")=request.form("jcolor")	
	rs("jcontent")=request.form("content")	
	rs("sjid")=request.form("sjid")	
	
	rs("sjpic")=request.form("image")	
	rs("sjditu")=request.form("ditu")	
	
	
	

	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('��ӳɹ�!');location='?action=list&sjid="&sjid&"'</script>"
end if

%>
      <script language="javascript" type="text/javascript">
// ��֤�û���������
function check_add(){
	var notnull;
	notnull=true;
	if (document.form1.title.value==""){
		alert("��Ϣ���Ʋ���Ϊ�գ�");
		document.form1.title.focus();
		notnull=false;
		}
	return notnull;
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">���--<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>--��Ϣ</font></strong></td>
        </tr>
      </table>
      <br>
      
      
      
      <script language="JavaScript">
<!--
function show(){
if (form1.jtype.value=="ҽҩ��ҵ��¼"){
 document.getElementById("z1").style.display='none';
 document.getElementById("z2").style.display='none';
 document.getElementById("z3").style.display='none';
 document.getElementById("z4").style.display='none';
}
else{
 document.getElementById("z1").style.display='';
 document.getElementById("z2").style.display='';
 document.getElementById("z3").style.display='';
 document.getElementById("z4").style.display='';
}

}

//-->
</script>
      
      
      
      <form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td width="75" align="center">�̼ҵ�����</td>
            <td>
            
            
            <% 	  
sql="select * from sh_sort order by anclassidorder asc"
set rs=conn.execute(sql)  
 %>
                  <select name="diqu" id="classid" style="width:100%">
                    <% do while not rs.eof %>
                    <option value="<%= rs("anclass") %>"><%= rs("anclass") %></option>
                    <%
		 rs.movenext
		loop
		 %>
                  </select>            </td>
          </tr>
          <tr>
            <td align="center">�̼����ͣ�</td>
            <td>
            
            
                        
						
                        
                        
                        
                      <input name="jtype" type="text" id="textfield2" size="50">
              
              <% 	  
sqle="select * from sjsort where sjid="&sjid&" order by anclassidorder asc"
set res=conn.execute(sqle)  
 %>
              <select name="diqu" id="diqu" onChange="(document.form1.jtype.value+=this.options[this.selectedIndex].value+',')">
              
                  <option>��ѡ��Ӫ����</option>
              
                  <% do while not res.eof %>
                  <option value="<%= res("anclass") %>"><%= res("anclass") %></option>
                  <%
		 res.movenext
		loop
		res.close
		set res=nothing

		 %>
              </select>            </td>
          </tr>
          <tr>
            <td align="center">�̼����ƣ�</td>
            <td><input type="text" name="jgname" id="textfield" style="width:100%"></td>
          </tr>
          <tr>
            <td align="center">�̼ҵ�ַ��</td>
            <td><input type="text" name="jgdz" id="textfield" style="width:60%"> �ʱ�:<input type="text" name="jgyb" id="textfield" style="width:30%"></td>
          </tr>
          <tr>
            <td align="center">��ϵ�绰��</td>
            <td><input type="text" name="jgtel" id="textfield" style="width:100%"></td>
          </tr>
          <tr>
            <td align="center">Ӫҵʱ�䣺</td>
            <td><input name="yysj" type="text" id="yysj" style="height:20; width:100%" value="10:00 �� 22:00" size="40"></td>
          </tr>
          <tr>
            <td align="center"><p>�ܱ���Ϣ��</p>            </td>
            <td><input name="zbxx" type="text" id="zbxx" size="40" style="height:20; width:100%"></td>
          </tr>
          <tr id="z5">
            <td align="right">����ĳ�վ��</td>
            <td><input name="zjdcz" type="text" id="zjdcz" size="40" style="height:20; width:100%"></td>
          </tr>
          <tr id="z3">
            <td align="center">ƽ��Ԥ�㣺</td>
            <td><input name="pjys" type="text" id="pjys" style="height:20; width:100%" size="40"></td>
          </tr>
          <tr>
            <td align="center">�ϴ�����ͼ��</td>
            <td><table width="553" border="0" cellspacing="1" cellpadding="0">
              <tr>
                <td width="323"><input name="image" type="text" id="image" size="40"></td>
                <td width="270"><iframe src="tongyongsc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td align="center">�ϴ���ͼ��</td>
            <td><table width="560" border="0" cellspacing="1" cellpadding="0">
              <tr>
                <td width="299"><input name="ditu" type="text" id="textfield4" size="40"></td>
                <td width="300"><iframe src="tongyongscc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
          <tr style="display:none">
            <td align="center">������ɫ��</td>
            <td>
<select name=jcolor size=1 id="jcolor">
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
          <tr>
            <td align="center">˵����</td>
            <td>
			
			
			
			
			
			 <textarea name="content" style="display:none"></textarea>
			 
			 
	   <iframe id="eWebEditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="250"></iframe>			</td>
          </tr>
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="���" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit2" value="����" style="width:80; height:30; cursor:hand">
              <input name="sjid" type="hidden" id="sjid" value="<%=request.querystring("sjid")%>"></td>
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
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,3
	rs("jtype")=request.form("jtype")
	rs("diqu")=request.form("diqu")
	rs("jgname")=request.form("jgname")
	rs("jgdz")=request.form("jgdz")
	rs("sjid")=request.form("sjid")
	rs("jgyb")=request.form("jgyb")
	rs("jgtel")=request.form("jgtel")
	rs("yysj")=request.form("yysj")
	rs("zjdcz")=request.form("zjdcz")
	rs("zbxx")=request.form("zbxx")
	rs("pjys")=request.form("pjys")
	rs("jcolor")=request.form("jcolor")	
	rs("jcontent")=request.form("content")	
	
	rs("sjpic")=request.form("image")	
	rs("sjditu")=request.form("ditu")	
	
	
	
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�޸ĳɹ�!');location='?action=list&sjid="&sjid&"'</script>"
	response.end
end if
id=trim(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,1
%>







      
      <script language="JavaScript">
<!--
function show(){
if (form1.jtype.value=="ҽҩ��ҵ��¼"){
 document.getElementById("z1").style.display='none';
 document.getElementById("z2").style.display='none';
 document.getElementById("z3").style.display='none';
 document.getElementById("z4").style.display='none';
}
else{
 document.getElementById("z1").style.display='';
 document.getElementById("z2").style.display='';
 document.getElementById("z3").style.display='';
 document.getElementById("z4").style.display='';
}

}

//-->
</script>
      










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
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">�޸�--<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>--��Ϣ</font></strong></td>
        </tr>
      </table>
      <br>
      <form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td width="78" align="right">�̼ҵ�����</td>
            <td><% 	  
sqle="select * from sh_sort order by anclassidorder asc"
set res=conn.execute(sqle)  
 %>
              <select name="diqu" id="diqu" style="width:100%">
                <% do while not res.eof %>
                <option value="<%= res("anclass") %>"<%if res("anclass")=rs("diqu") then response.Write("selected")%>><%= res("anclass") %></option>
                <%
		 res.movenext
		loop
		res.close
		set res=nothing
		 %>
                </select>            </td>
          </tr>
          <tr>
            <td width="78" align="right">��Ӫ���ͣ�</td>
            <td>
			
            
            
              <input name="jtype" type="text" id="textfield2" size="50" value="<%=rs("jtype")%>">
              
              <% 	  
sqle="select * from sjsort where sjid="&sjid&" order by anclassidorder asc"
set res=conn.execute(sqle)  
 %>
              <select name="11diqu" id="1diqu" onChange="(document.form1.jtype.value+=this.options[this.selectedIndex].value+',')">
              
                  <option>��ѡ��Ӫ����</option>
              
                  <% do while not res.eof %>
                  <option value="<%= res("anclass") %>"><%= res("anclass") %></option>
                  <%
		 res.movenext
		loop
		res.close
		set res=nothing

		 %>
              </select>            </td>
          </tr>
          <tr>
            <td width="78" align="right">�̼����ƣ�</td>
            <td><input type="text" name="jgname" id="textfield" style="width:100%" value="<%=rs("jgname")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right">�̼ҵ�ַ��</td>
            <td><input type="text" name="jgdz" id="textfield" style="width:60%" value="<%=rs("jgdz")%>"> �ʱ�:<input type="text" name="jgyb" id="textfield" style="width:30%" value="<%=rs("jgyb")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right">��ϵ�绰��</td>
            <td><input type="text" name="jgtel" id="textfield" style="width:100%" value="<%=rs("jgtel")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right">Ӫҵʱ�䣺</td>
            <td><input name="yysj" type="text" id="yysj" size="40" style="height:20; width:100%" value="<%=rs("yysj")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right"><p>�ܱ���Ϣ��</p></td>
            <td><input name="zbxx" type="text" id="yywz3" size="40" style="height:20; width:100%" value="<%=rs("zbxx")%>"></td>
          </tr>
          
          
         <tr id="z3">
            <td width="78" align="right">����ĳ�վ��</td>
            <td><input name="zjdcz" type="text" id="zjdcz" size="40" style="height:20; width:100%"  value="<%=rs("zjdcz")%>"></td>
          </tr>
          
            <td width="78" align="right">ƽ��Ԥ�㣺</td>
            <td><input name="pjys" type="text" id="pjys" size="40" style="height:20; width:100%" value="<%=rs("pjys")%>"></td>
          </tr>  <tr>
              <td align="center">�ϴ�����ͼ��</td>
              <td><table width="553" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="323"><input name="image" type="text" id="image2" size="40" value="<%=rs("sjpic")%>"></td>
                    <td width="270"><iframe src="tongyongsc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td align="center">�ϴ���ͼ��</td>
              <td><table width="560" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="299"><input name="ditu" type="text" id="textfield3" size="40" value="<%=rs("sjditu")%>"></td>
                    <td width="300"><iframe src="tongyongscc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                  </tr>
              </table></td>
            </tr>
          
          <tr>
            <td width="78" align="right">������ɫ��</td>
            <td>
			
			
<select name=jcolor size=1 id="jcolor">
<option value="">ѡ����ɫ</option>
<option style="background-color:Black;color:Black" value=Black <%if rs("jcolor")="Black" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Red;color:Red" value=Red <%if rs("jcolor")="Red" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Yellow;color:Yellow" value=Yellow <%if rs("jcolor")="Yellow" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Green;color:Green" value=Green <%if rs("jcolor")="Green" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Orange;color:Orange" value=Orange <%if rs("jcolor")="Orange" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Purple;color:Purple" value=Purple <%if rs("jcolor")="Purple" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Blue;color:Blue" value=Blue <%if rs("jcolor")="Blue" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Brown;color:Brown" value=Brown <%if rs("jcolor")="Brown" then response.write "selected" %>>�� ɫ</option>
<option style="background-color:Teal;color:Teal" value=Teal <%if rs("jcolor")="Teal" then response.write "selected" %>>ī ��</option>
<option style="background-color:Navy;color:Navy" value=Navy <%if rs("jcolor")="Navy" then response.write "selected" %>>�� ��</option>
<option style="background-color:Maroon;color:Maroon" value=Maroon <%if rs("jcolor")="Maroon" then response.write "selected" %>>�� ʯ</option>
<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF" <%if rs("jcolor")="#00FFFF" then response.write "selected" %>>�� ��</option>
<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4" <%if rs("jcolor")="#7FFFD4" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4" <%if rs("jcolor")="#FFE4C4" then response.write "selected" %>>�� ��</option>
<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00" <%if rs("jcolor")="#7FFF00" then response.write "selected" %>>�� ��</option>
<option style="background-color:#D2691E;color: #D2691E" value="#D2691E" <%if rs("jcolor")="#D2691E" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50" <%if rs("jcolor")="#FF7F50" then response.write "selected" %>>ש ��</option>
<option style="background-color:#6495ED;color: #6495ED" value="#6495ED" <%if rs("jcolor")="#6495ED" then response.write "selected" %>>�� ��</option>
<option style="background-color:#DC143C;color: #DC143C" value="#DC143C" <%if rs("jcolor")="#DC143C" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FF1493;color: #FF1493" value="#FF1493" <%if rs("jcolor")="#FF1493" then response.write "selected" %>>õ���</option>
<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF" <%if rs("jcolor")="#FF00FF" then response.write "selected" %>>�� ��</option>
<option style="background-color:#FFD700;color: #FFD700" value="#FFD700" <%if rs("jcolor")="#FFD700" then response.write "selected" %>>�� ��</option>
<option style="background-color:#DAA520;color: #DAA520" value="#DAA520" <%if rs("jcolor")="#DAA520" then response.write "selected" %>>�� ��</option>
<option style="background-color:#808080;color: #808080" value="#808080" <%if rs("jcolor")="#808080" then response.write "selected" %>>�� ��</option>
<option style="background-color:#778899;color: #778899" value="#778899" <%if rs("jcolor")="#778899" then response.write "selected" %>>�� ��</option>
<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE" <%if rs("jcolor")="#B0C4DE" then response.write "selected" %>>�� ��</option>
</select>			</td>
          </tr>
          <tr>
            <td width="78" align="right">˵����</td>
            <td><textarea name="content" style="display:none"><%=rs("jcontent")%></textarea>
                <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="250"></iframe></td>
          </tr>
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="�޸�" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit22" value="����" style="width:80; height:30; cursor:hand">
              <input name="sjid" type="hidden" id="sjid" value="<%=sjid%>"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <% 
if trim(request.querystring("action"))="del" then
	id=trim(request.querystring("id"))
	sjid=request.QueryString("sjid")
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�޸ĳɹ�!');location='?action=list&sjid="&sjid&"'</script>"
end if
 %>
 
 
      <% 
if trim(request.querystring("zhiding"))="zdyes" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,3
	rs("tuijian")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('��ȡ���ö�!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>


      <% 
if trim(request.querystring("zhiding"))="zdno" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,3,2
	rs("tuijian")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('�ö��ɹ�!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>
 
 
 
      <% 
if trim(request.querystring("toutiao"))="ttyes" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
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
	response.write "<script>alert('ͷ����Ϣ��ȡ��!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>
 
 
      <% 
if trim(request.querystring("toutiao"))="ttno" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
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
	response.write "<script>alert('ͷ�����óɹ�!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>
 
 
 
 
 
 
 
 
 
 
<%
if trim(request.querystring("action"))="list" then
 %>
 
 
 
 
 
  
 
 
 
 
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:7px">
  <form name="form" method="post" action="?Action=list&sjid=<%=sjid%>">
    <tr> 
      <td align="left">��<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>--��Ϣ������

        <input name="keywords" type="text" class="input" id="keywords" style="width:150px;height:21px; padding-left:5px" onFocus='this.select()' onBlur="if (this.value ==''){this.value=this.defaultValue}" onClick="if(this.value=='������Ϣ�ؼ���')this.value=''" value="������Ϣ�ؼ���">
	  <input name="Submit" type="submit" class="bt" id="Submit" value="����">
      </td>
      <td align="right">&nbsp;</td>
    </tr>
  </form>
</table>
 
 
 
 
 
 <table width="100%" border="0" cellspacing="0" cellpadding="6">
   <tr>
     <td width="82%" valign="top"><table width="99%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
       <tr align="center">
         <td width="26%" align="left" bgcolor="#DFEFFF" style="padding-left:10px">�������� 
           
           	  <% 	  
sql_classid="select * from sh_sort"
set rs_classid=conn.execute(sql_classid)  
 %>	  
	  <select name="classid" id="classid" onchange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
	<option value="?action=list&dq=&sjid=<%=sjid%>">ѡ�����</option>

	  <% do while not rs_classid.eof %>
        <option value="?action=list&dq=<%= rs_classid("anclass") %>&sjid=<%=sjid%>" <% if request.QueryString("dq")=rs_classid("anclass") then response.write"selected"%>><%= rs_classid("anclass") %></option>
		<%
		 rs_classid.movenext
		loop
		 %>
	<option value="?action=list&dq=&sjid=<%=sjid%>">���е���</option>
    </select>           </td>
         <td width="24%" align="left" bgcolor="#DFEFFF">�绰</td>
         <td width="11%" align="center" bgcolor="#DFEFFF">����ͼ|��ͼ</td>
         <td width="13%" align="center" bgcolor="#DFEFFF" style="display:none">�ֵ����</td>
         <td width="10%" align="center" bgcolor="#DFEFFF">�Ż�ȯ</td>
         <td width="5%" align="center" bgcolor="#DFEFFF">�ö�</td>
         <td width="5%" align="center" bgcolor="#DFEFFF">ɾ��</td>
         <td width="6%" align="center" bgcolor="#DFEFFF">�༭</td>
       </tr>
     </table>
       <br>
       <%
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   	newsclass=request.form("newsclass")
	keywords=request.form("keywords")
	cid=request.querystring("cid")
	dq=request.querystring("dq")
	set rs=server.createobject("adodb.recordset")
		
	sql="select * from jiedai_paijiexi where 1=1 "
	if keywords<>"" then
	sql=sql+" and jgname like '%"&keywords&"%' or jcontent  like '%"&keywords&"%' "
	end if
	
	if sjid<>"" then
	sql=sql+" and sjid="&sjid&" "
	end if
	
	
	
	if cid<>"" then
	sql=sql+" and jtype='"&cid&"' "
	end if

	if dq<>"" then
	sql=sql+" and diqu='"&dq&"' "
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
       <table width="99%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
         <tr>
           <td width="26%" class="line" style="padding-left:10px">��&nbsp;<font color="<%=rs("jcolor")%>"><%= rs("jgname") %></font></td>
           <td width="24%" class="line"><%= rs("jgtel") %></td>
           <td width="11%" align="center" class="line"><%if rs("sjpic")<>"" then%><a href="../uploadfile/<%=rs("sjpic")%>" target="_blank"><img src="images/arrow38.gif" width="18" height="18" border="0" align="absmiddle"></a><%else%><img src="images/arrow38.gif" width="18" height="18" border="0" align="absmiddle" style="filter:progid:DXImageTransform.Microsoft.BasicImage(grayScale=1)"><%end if%>| <%if rs("sjditu")<>"" then%><a href="../uploadfile/<%=rs("sjditu")%>" target="_blank"><img src="images/arrow38.gif" width="18" height="18" align="absmiddle" border="0"></a><%else%><img src="images/arrow38.gif" width="18" height="18" align="absmiddle" border="0" style="filter:progid:DXImageTransform.Microsoft.BasicImage(grayScale=1)"><%end if%></td>
           <td width="13%" align="center" class="line" style="display:none"><a href="jiedai_fendian.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=add">���</a> | <a href="jiedai_fendian.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=list">����</a></td>
           <td width="10%" align="center" class="line"><a href="jiedai_huoban.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=add">���</a> | <a href="jiedai_huoban.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=list">����</a></td>
           <td width="5%" align="center" class="line"><%if rs("tuijian")=1 then%>
               <a href="?Action=list&zhiding=zdyes&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>&dq=<%=dq%>&sjid=<%=sjid%>"><img src="images/Ok.gif" alt="���ö�" width="16" height="16" border="0" /></a>
               <%else%>
               <a href="?Action=list&zhiding=zdno&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>&dq=<%=dq%>&sjid=<%=sjid%>"><img src="images/err.gif" alt="δ�ö�" width="12" height="11" border="0" /></a>
               <%end if%></td>
           <td width="5%" align="center" class="line"><a href="?action=del&id=<%= rs("id") %>&sjid=<%=sjid%>" onClick="return window.confirm('ȷ��ɾ����?');">[ɾ��]</a> </td>
           <td width="6%" align="center" class="line"><a href="?action=edit&id=<%= rs("id") %>&sjid=<%=sjid%>">[�༭]</a> </td>
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
        <a href="?action=list&page=1&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">��ҳ</a>
        <% else %>
      ��ҳ
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">��һҳ</a>
      <% else %>
      ��һҳ
      <% end if %>
      &nbsp;
      <% if page<rs.recordcount then %>
      <a href="?action=list&page=<%= rs.recordcount %>&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">ĩҳ</a>
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