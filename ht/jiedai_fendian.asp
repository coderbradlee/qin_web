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

OrderBy = request("OrderBy") '��ǰ��¼��������� at_no ��at_id�Ǽ�¼��id ���������id������Ƿֿ��� 
at_ID = request("at_ID") '��ǰ��¼id 
action = request("action") '�ƶ����� 
paixu=request("paixu")

if paixu="up" then '���� 
		'�����ж��ǲ����Ѿ��ƶ�����ǰ 
		sql="select top 1 id from jiedai_fendian where flag<"&OrderBy&" and typeclass='"&request.QueryString("typecss")&"' order by flag desc" 
		set rs=server.createobject("adodb.recordset") 
		rs.open sql,conn,1,3 
		if rs.eof then 'ǰ��û�м�¼ �� 
		rs.close 
		set rs=nothing 
		response.write "<script>alert('���󣬸�����Ϣ�Ѿ�λ����λ��');window.history.back();</script>" 
		response.end 
		end if 
		before_id=rs("id") 'ǰһ����¼��id 
		rs.close 
		set rs=nothing 

		'�޸�ǰһ����¼��id 
		sql2="select flag from jiedai_fendian where typeclass='"&request.QueryString("typecss")&"' and id="&before_id '��ѯǰһ����¼ 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		before_Orderby=rs2("flag") '��¼�ü�¼��������� 
		rs2("flag")=OrderBy '���ĸü�¼��������� 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_fendian where typeclass='"&request.QueryString("typecss")&"' and ID="&at_ID '��ѯ��ǰ��¼��Ҫ�ƶ��ļ�¼�� 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("flag")=before_Orderby '�޸ĵ�ǰ��¼���������Ϊǰһ��¼����ţ����� ������¼��Ž��� ʵ������ 
		rs3.update 
		rs3.close 
		set rs3=nothing 

elseif paixu="Down" then '���������Ƶ���һ�� 
		sql="select top 1 id from jiedai_fendian where typeclass='"&request.QueryString("typecss")&"' and  flag>"&OrderBy&" order by flag asc" 
		set rs=server.createobject("adodb.recordset") 
		rs.open sql,conn,1,3 
		if rs.eof then 
		rs.close 
		set rs=nothing 
		response.write "<script>alert('���󣬸�����Ϣ�Ѿ�λ�����һλ��');window.history.back();</script>" 
		response.end 
		end if 
		end_id=rs("ID") 
		'response.Write(end_id) 
		'response.End() 
		rs.close 
		set rs=nothing 
		
		sql2="select flag from jiedai_fendian where typeclass='"&request.QueryString("typecss")&"' and  ID="&end_id 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		end_Orderby=rs2("flag") 
		rs2("flag")=OrderBy 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select flag from jiedai_fendian where typeclass='"&request.QueryString("typecss")&"' and  ID="&at_ID 
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
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>--<font color="#CC6600"><%=res("jgname")%></font>--�ֵ���Ϣ����
    </font></strong></td>
  </tr>
</table><br>


<% 
classid=trim(request.querystring("classid"))
set rs=server.createobject("adodb.recordset")
sql="select * from jiedai_fendian where sid="&sid&" order by id desc"
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
if rs.bof then response.write("<div align=center><br><br><br><br><br><br><font color=red>����</font>��Ϣ��<br><br><br><br><br><br></div>")
 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr align="center">
    <td width="50" align="center" class="line">��</td>
    <td width="950" align="left" class="line"><span class="style1">&nbsp;<%= rs("title") %></span>��<font color="#FF0000"><b><%=rs("tel")%> </b></font>��<%if rs("ditu")<>"" then%><a href="../uploadfile/<%=rs("ditu")%>" target="_blank"><img src="images/arrow38.gif" width="18" height="18" border="0" align="absmiddle"></a><%else%><img src="images/arrow38.gif" width="18" height="18" border="0" align="absmiddle" style="filter:progid:DXImageTransform.Microsoft.BasicImage(grayScale=1)"><%end if%></td>
    <td width="132" align="center" class="line">&nbsp;</td>
    <td width="50" align="center" class="line"><a href="?action=edit&id=<%= rs("id") %>&sid=<%=sid%>&sjid=<%=sjid%>">[�޸�]</a></td>
    <td width="50" align="center" class="line"><a href="?action=del&id=<%= rs("id") %>&sid=<%=sid%>&sjid=<%=sjid%>" onClick="return window.confirm('ȷ��ɾ����?');">[ɾ��]</a></td>
  </tr>
</table>
<br>
<% 
rs.movenext
next
else
response.write"<div align=center><br><font color=red>����</font>��Ϣ��<br><br></div>"
end if
%>
<table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
  <tr> 
    <td width="962" align="center">��<%= page %>ҳ&nbsp; 
      <% if page<>1 then %>
      <a href="?action=list&page=1&classid=<%= classid %>">��ҳ</a> 
      <% else %>
      ��ҳ 
      <% end if %>
      &nbsp; 
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>&classid=<%= classid %>">��һҳ</a> 
      <% else %>
      ��һҳ 
      <% end if %>
      &nbsp; 
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>&classid=<%= classid %>">��һҳ</a> 
      <% else %>
      ��һҳ 
      <% end if %>
      &nbsp; 
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%=rs.pagecount%>&classid=<%= classid %>">ĩҳ</a> 
      <% else %>
      ĩҳ 
      <% end if %>
      &nbsp;����<%= rs.recordcount %>��</td>
    <td width="270" align="center">ת���� 
      <select name="select" onchange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
        <%for m = 1 to rs.pagecount%>
        <option value="?action=list&page=<%=m%>&classid=<%= classid %>"><%=m%></option>
        <% next %>
      </select>
    ҳ</td>
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
		sql="select * from jiedai_fendian"
		rs.open sql,conn,1,3
		rs.addnew
		rs("sid")=request.Form("sid")
		rs("sjid")=request.Form("sjid")
		rs("title")=request.Form("title")
		rs("tel")=request.Form("tel")
		rs("yusuan")=request.Form("yusuan")
		rs("dizhi")=request.Form("dizhi")
		rs("yytime")=request.Form("yytime")
		rs("ditu")=request.form("image")
		
		rs.update
		rs.requery
		
		response.write("<script>alert('��ӳɹ�');location='?action=add&sid="&sid&"&sjid="&sjid&"';</script>")
		response.end
		
		
		rs.close
		set rs=nothing
	end if
%>
<script language="javascript" type="text/javascript">
// ��֤�û���������
function check_add(){
if(document.form1.title.value==""){alert("�������������ƣ�");document.form1.title.focus();return false;}
return true
}
</script>



<%
	set res=server.createobject("adodb.recordset")
	sqle="select * from jiedai_paijiexi where id="&request.QueryString("sid")&" "
	res.open sqle,conn,1,1


%>


<form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
  <table width="100%" height="240" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr>
      <td height="30" colspan="2" align="left"><table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#D3E5FA">
  <tr>
    <td width="9%" bgcolor="#F4FBFF" style="padding-left:12px;"><font color="#3399CC"><%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>
      --</font></td>
    <td width="91%" bgcolor="#F4FBFF" style="padding-left:12px; font-size:18px; font-weight:bold; color:#FF0000; padding-top:4px"><font color="#3399CC">��</font><%=res("jgname")%><font color="#3399CC">��ӷֵ�</font></td>
  </tr>
</table></td>
    </tr>
    <tr> 
      <td width="50" height="30" align="center">����:</td>
      <td><input name="title" type="text" id="classid" size="40"> 
      ��</td>
    </tr>
    <tr>
      <td height="30" align="center">�绰:</td>
      <td><input name="tel" type="text" id="classid2" size="40"></td>
    </tr>
    <tr>
      <td height="30" align="center">Ԥ��:</td>
      <td><input name="yusuan" type="text" id="classid3" size="40"></td>
    </tr>
    <tr>
      <td height="30" align="center">ʱ��:</td>
      <td><input name="yytime" type="text" id="classid5" value="10:00 �� 22:00" size="40"></td>
    </tr>
    <tr>
      <td height="30" align="center">��ַ:</td>
      <td><input name="dizhi" type="text" id="classid4" size="40"></td>
    </tr>
    <tr>
      <td align="center">��ͼ:</td>
      <td><table width="553" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td width="323"><input name="image" type="text" id="ditu" size="40"></td>
            <td width="270"><iframe src="tongyongsc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
          </tr>
      </table></td>
    </tr>
    
    <tr>
      <td height="30" colspan="2" background="images/bg_title.gif" style="padding-left:50"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="9"></td>
        </tr>
      </table>
        <%if request.form("add")="add" then
		 response.write"<img src=images/cms-ico7.gif width=12 height=11><font color=#ff0000><b></b>��Ϣ����ӳɹ�</font>"
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
		sql="select * from jiedai_fendian where id="&id
		rs.open sql,conn,1,3


		'rs("sid")=request.Form("sid")
		'rs("sjid")=request.Form("sjid")
		rs("title")=request.Form("title")
		rs("tel")=request.Form("tel")
		rs("yusuan")=request.Form("yusuan")
		rs("dizhi")=request.Form("dizhi")
		rs("yytime")=request.Form("yytime")
		rs("ditu")=request.form("image")

		rs.update
		rs.requery
		
	response.write("<script>alert('�޸ĳɹ�');location='?action=list&sid="&sid&"&sjid="&sjid&"';</script>")
		response.end

		
		
		rs.close
		set rs=nothing
		
		
		
	end if
	
		id=trim(request.querystring("id"))
		sql="select * from jiedai_fendian where id="&id
		set rs=conn.execute(sql)

%>




<%
	set res=server.createobject("adodb.recordset")
	sqle="select * from jiedai_paijiexi where id="&request.QueryString("sid")&" "
	res.open sqle,conn,1,1


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
<form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>&sid=<%=sid%>&sjid=<%=sjid%>" onSubmit="return check_edit()">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr>
      <td valign="top">
      
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
	 response.write"<font color=#ff0000>��ʳ</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>����</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>�ͷ�</font>" 
	   end if%>
                --</font></td>
              <td width="91%" bgcolor="#F4FBFF" style="padding-left:12px; font-size:18px; font-weight:bold; color:#FF0000; padding-top:4px"><font color="#3399CC">��</font><%=res("jgname")%><font color="#3399CC">�޸ķֵ���Ϣ</font></td>
            </tr>
          </table>
          <table width="100%" height="180" border="0" cellpadding="0" cellspacing="0">
            <tr bordercolor="#cccccc">
              <td width="50" height="30" align="center">����:</td>
              <td><input name="title" type="text" id="classid10" size="40" value="<%=rs("title")%>">
              </td>
            </tr>
            <tr bordercolor="#cccccc">
              <td height="30" align="center">�绰:</td>
              <td><input name="tel" type="text" id="classid9" value="<%=rs("tel")%>" size="40"></td>
            </tr>
            <tr bordercolor="#cccccc">
              <td height="30" align="center">Ԥ��:</td>
              <td><input name="yusuan" type="text" id="classid8" size="40" value="<%=rs("yusuan")%>"></td>
            </tr>
            <tr bordercolor="#cccccc">
              <td height="30" align="center">ʱ��:</td>
              <td><input name="yytime" type="text" id="classid7" size="40" value="<%=rs("yytime")%>"></td>
            </tr>
            <tr bordercolor="#cccccc">
              <td height="30" align="center">��ַ:</td>
              <td><input name="dizhi" type="text" id="classid6" size="40"  value="<%=rs("dizhi")%>"></td>
            </tr>
            <tr bordercolor="#cccccc">
              <td align="center">��ͼ:</td>
              <td><table width="553" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="323"><input name="image" type="text" id="ditu2" size="40" value="<%=rs("ditu")%>"></td>
                    <td width="270"><iframe src="tongyongsc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                  </tr>
              </table></td>
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
		 response.write"<img src=images/cms-ico7.gif width=12 height=11><font color=#ff0000><b>"&rs("classid")&"-</b>��Ϣ���޸ĳɹ�</font>"
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
	sql="select * from jiedai_fendian where id="&id
	rs.open sql,conn,2,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
		response.write("<script>alert('ɾ���ɹ�');location='?action=list&sid="&sid&"&sjid="&sjid&"';</script>")
end if
 %>
</body>
</html>                                                                             