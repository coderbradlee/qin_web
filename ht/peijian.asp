<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
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
		sql="select top 1 id from peijian where jhpx<"&OrderBy&" order by jhpx desc" 
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
		sql2="select jhpx from peijian where id="&before_id '��ѯǰһ����¼ 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		before_Orderby=rs2("jhpx") '��¼�ü�¼��������� 
		rs2("jhpx")=OrderBy '���ĸü�¼��������� 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select jhpx from peijian where ID="&at_ID '��ѯ��ǰ��¼��Ҫ�ƶ��ļ�¼�� 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("jhpx")=before_Orderby '�޸ĵ�ǰ��¼���������Ϊǰһ��¼����ţ����� ������¼��Ž��� ʵ������ 
		rs3.update 
		rs3.close 
		set rs3=nothing 

elseif paixu="Down" then '���������Ƶ���һ�� 
		sql="select top 1 id from peijian where jhpx>"&OrderBy&" order by jhpx asc" 
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
		
		sql2="select jhpx from peijian where ID="&end_id 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		end_Orderby=rs2("jhpx") 
		rs2("jhpx")=OrderBy 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select jhpx from peijian where ID="&at_ID 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("jhpx")=end_Orderby 
		rs3.update 
		rs3.close 
		set rs3=nothing 
end if 


















%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td valign="top"><table width="173" height="148" border="0" cellpadding="7" cellspacing="0">
          <tr>
            <td align="center" valign="top"><table width="73%" height="33" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" class="typecss">
                      <tr>
                        <td align="center" style="padding-left:12px;"><b>�������</b></td>
                      </tr>
                    </table>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="1" bgcolor="#ECECEC"></td>
                        </tr>
                      </table>
                    <%
		  		 if request.QueryString("Caseid")<>"" then
		 weburl=int(request.QueryString("Caseid"))
		 end if
		 
		set rs=server.createobject("adodb.recordset")
		sql="select * from peijian_class order by anclassidorder asc"
	rs.open sql,conn,1,1
do while not rs.eof
		 
		  %>
                      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" class="dzt">
                        <tr>
                          <%if weburl=int(rs("anclassid")) then%>
                          <td align="center" style="padding-left:12px;" bgcolor="#F0F0F0"><a href="?Caseid=<%=rs("anclassid")%>&action=list" class="dzt"><%=rs("anclass")%></a></td>
                          <%
				else
				%>
                          <td align="center" style="padding-left:12px;"><a href="?Caseid=<%=rs("anclassid")%>&action=list" class="dzt"><%=rs("anclass")%></a></td>
                          <%
				 end if
				 %>
                        </tr>
                      </table>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="1" bgcolor="#ECECEC"></td>
                        </tr>
                      </table>
                    <%
		  
		  rs.movenext
		  loop
		  rs.close
		  set rs=nothing
		  
		  
		  
		  %>                  </td>
                </tr>
            </table></td>
            </tr>
        </table></td>
        <td valign="top"><% 
'classid=trim(request.querystring("classid"))
set rs=server.createobject("adodb.recordset")
	Caseid=trim(request.querystring("Caseid"))
	Nclassid=trim(request.querystring("Nclassid"))
	sql="select * from peijian where 1=1 "
		
		if Caseid<>"" then
		sql=sql+" and anclassid="&Caseid&""
		end if
		
		if Nclassid<>"" then
		sql=sql+" and nclassid="&Nclassid&""
		end if
		
		sql=sql+"  order by id desc"




rs.open sql,conn,1,1
rs.pagesize=10

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
          <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" bgcolor="#D3E5FA">
            <tr>
              <td style="padding-left:15"><strong><font color="#215dc6">��Ϣ����</font></strong> </td>
            </tr>
          </table>
          <br>
          <%
if rs.bof then response.write("<center><br><br><br><br><br><br><font color=red>����</font>��Ϣ��<br><br><br><br><br><br></center>")
 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
          <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
            <tr align="center">
              <td align="center" class="line">&nbsp;</td>
              <td align="left" class="line"><span class="style1"><b><%= rs("��Ʒ����") %></b></span></td>
              <td align="center" class="line">&nbsp;</td>
              <td align="center" class="line">&nbsp;</td>
              <td align="center" class="line">&nbsp;</td>
            </tr>
            <tr align="center">
              <td width="50" align="center" class="line"><%= rs("jhpx") %></td>
              <td align="left" class="line"><span class="style1">&nbsp;<img src="../uploadfile/<%=rs("��ƷͼƬ")%>" width="100" height="100"></span></td>
              <td width="132" align="center" class="line">&nbsp;<table width="60" border="0" cellspacing="0" cellpadding="0" style="display:none">
                  <tr>
                    <td width="30"><a href="?action=list&paixu=up&OrderBy=<%=rs("jhpx")%>&at_id=<%=rs("id")%>"><img src="images/up.gif" width="15" height="16" border="0"></a> ��</td>
                    <td width="30"><a href="?action=list&paixu=Down&OrderBy=<%=rs("jhpx")%>&at_id=<%=rs("id")%>"><img src="images/down.gif" width="15" height="16" border="0"></a></td>
                  </tr>
              </table></td>
              <td width="50" align="center" class="line"><a href="?action=edit&id=<%= rs("id") %>&Caseid=<%=Caseid%>&Nclassid=<%=Nclassid%>">[�޸�]</a></td>
              <td width="50" align="center" class="line"><a href="?action=del&id=<%= rs("id") %>&Caseid=<%=Caseid%>&Nclassid=<%=Nclassid%>" onClick="return window.confirm('ȷ��ɾ����?');">[ɾ��]</a></td>
            </tr>
          </table>
          <br>
          <% 
rs.movenext
next
else
response.write"������Ϣ"
end if
%>
          <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
            <tr>
              <td width="766" align="center">��<%= page %>ҳ&nbsp;
                  <% if page<>1 then %>
                  <a href="?action=list&page=1&classid=<%= classid %>&Caseid=<%=request.QueryString("Caseid")%>&Nclassid=<%=request.QueryString("Nclassid")%>">��ҳ</a>
                  <% else %>
                ��ҳ
                <% end if %>
                &nbsp;
                <% if page>1 then %>
                <a href="?action=list&page=<%= page-1 %>&classid=<%= classid %>&Caseid=<%=request.QueryString("Caseid")%>&Nclassid=<%=request.QueryString("Nclassid")%>">��һҳ</a>
                <% else %>
                ��һҳ
                <% end if %>
                &nbsp;
                <% if page<rs.pagecount then %>
                <a href="?action=list&page=<%= page+1 %>&classid=<%= classid %>&Caseid=<%=request.QueryString("Caseid")%>&Nclassid=<%=request.QueryString("Nclassid")%>">��һҳ</a>
                <% else %>
                ��һҳ
                <% end if %>
                &nbsp;
                <% if page<rs.pagecount then %>
                <a href="?action=list&page=<%=rs.pagecount%>&classid=<%= classid %>&Caseid=<%=request.QueryString("Caseid")%>&Nclassid=<%=request.QueryString("Nclassid")%>">ĩҳ</a>
                <% else %>
                ĩҳ
                <% end if %>
                &nbsp;����<%= rs.recordcount %>��</td>
              <td width="217" align="center">ת����
                <select name="select" onchange='javascript:window.open(this.options[this.selectedindex].value,"_self")'>
                    <%for m = 1 to rs.pagecount%>
                    <option value="?action=list&page=<%=m%>&classid=<%= classid %>"><%=m%></option>
                    <% next %>
                  </select>
                ҳ</td>
            </tr>
          </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<% end if %>
<% if trim(request.querystring("action"))="add" then
	if trim(request.form("add"))="add" then
		classid=trim(request.form("classid"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from peijian"
		rs.open sql,conn,1,3
		rs.addnew
		rs("��ƷͼƬ")=request("image")
		if request.form("flag")<>"" then
		rs("jhpx")=request.form("flag")
		else
		rs("jhpx")=100
		end if
		rs("��Ʒ����")=request("content")
		rs("��Ʒ����")=request("title") '����
		rs("anclassid")=int(request("anclassid")) '����
rs("nclassid")=int(request("nclassid")) 'С��		
		rs.update
		rs.requery
		rs.close
		set rs=nothing
	end if
%>
<script language="javascript" type="text/javascript">
// ��֤�û���������
function check_add(){
if(document.form1.classid.value==""){alert("���������ı��⣡");document.form1.classid.focus();return false;}
return true
}
</script>








<form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
  <table width="100%" height="120" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr> 
      <td width="50" height="30" align="center">����:</td>
      <td width="1184"><input name="title" type="text" id="title" size="40"> 
      ������:
        <input name="flag" type="text" id="flag" size="8"></td>
    </tr>
    <tr>
      <td align="center">ͼƬ:</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="25%"><input name="image" type="text" id="image" size="40" style="height:25"></td>
            <td width="75%"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
          </tr>
      </table></td>
    </tr>
    <tr height="30">
      <td align="center">����:</td>
      <td><%
	  set rs=server.CreateObject("adodb.recordset")
     rs.open "select * from peijian_class order by anclassidorder",conn,1,1
	if rs.eof and rs.bof then
	response.write "���������Ŀ��"
	response.end
	else
  %>
        <select name="anclassid" size="1" id="anclassid">
          <option selected value="<%=rs("anclassid")%>"><%=trim(rs("anclass"))%></option>
          <%
        dim selclass
         selclass=rs("anclassid")
        rs.movenext
        do while not rs.eof
	%>
          <option value="<%=rs("anclassid")%>"><%=trim(rs("anclass"))%></option>
          <%
        rs.movenext
        loop
		end if
        rs.close
	%>
        </select></td>
    </tr>
    <tr>
      <td height="30" align="center">
	  ����:</td>
      <td>
	  <textarea name="content" cols="" rows="" style="display:none"></textarea>
	   <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="350"></iframe>	  </td>
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
<input name="add" type="hidden" id="add" value="add"></td>
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
		classid=trim(request.form("classid"))
		
		Caseid=trim(request.QueryString("Caseid"))
		Nclassid=trim(request.QueryString("Nclassid"))
		
		 scontent =request.form("content")
		set rs=server.createobject("adodb.recordset")
		sql="select * from peijian where id="&id
		rs.open sql,conn,1,3
'		rs("classid")=classid
		rs("��Ʒ����")=scontent
		
		if request.form("flag")<>"" then
		rs("jhpx")=request.form("flag")
		else
		rs("jhpx")=100
		end if
		
		rs("��Ʒ����")=request("title")
		rs("��ƷͼƬ")=request("image")
		rs("anclassid")=int(request("anclassid")) '����
rs("nclassid")=int(request.form("nclassid")) 'С��
		
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		
		response.Redirect("Jiedai_dypm.asp?Action=list&&Caseid="&Caseid&"&Nclassid="&Nclassid&"")
		
	end if
	
		id=trim(request.querystring("id"))
		sql="select * from peijian where id="&id
		set ras=conn.execute(sql)

%>
<script language="javascript" type="text/javascript">
// ��֤�û���������
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.classid.value==""){
		alert("���ⲻ��Ϊ�գ�");
		document.form1.classid.focus();
		notnull=false;
		}
		return notnull;
	}
</script>





<form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>&Caseid=<%=request.QueryString("Caseid")%>&Nclassid=<%=request.QueryString("Nclassid")%>" onSubmit="return check_edit()">
  <table width="100%" height="417" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr>
      <td height="20" bgcolor="#D3E5FA" style="padding-left:15"><b></b>&nbsp;
        <input name="add" type="hidden" id="add" value="add"></td>
      </tr>
    <tr>
      <td height="323" valign="top">
      
      	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
      	  <table width="100%" height="53" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
            <tr>
              <td width="5%" align="center">����:</td>
              <td width="95%"><input name="title" type="text" id="title" size="40" value="<%=ras("��Ʒ����")%>">
������:
  <input name="flag" type="text" id="flag" value="<%=ras("jhpx")%>" size="8"></td>
            </tr>
            <tr>
              <td align="center">����:</td>
              <td>
			  
			  
			  
			  
			  
			  
			  
			  <%dim rs1
	  set rs=server.CreateObject("adodb.recordset")
			   		set rs1=server.CreateObject("adodb.recordset")
					rs1.open "select * from peijian where id="&id,conn,1,1
					rs.open "select * from peijian_class order by anclassidorder",conn,1,1
					if rs.eof and rs.bof then
					response.write "���������Ŀ��"
					response.end
					else
				%>
        <select name="anclassid" size="1" id="anclassid" onChange="changelocation(document.myform.anclassid.options[document.myform.anclassid.selectedIndex].value)">
          <%do while not rs.eof%>
          <option value="<%=rs("anclassid")%>" <%if rs1("anclassid")=rs("anclassid") then%>selected<%end if%>><%=trim(rs("anclass"))%></option>
          <%
					rs.movenext
					loop
					end if
					rs.close
				%>
        </select></td>
            </tr>
            <tr>
              <td align="center">ͼƬ:</td>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="33%"><input name="image" type="text" id="image" style="height:24" value="<%=ras("��ƷͼƬ")%>" size="40"></td>
                    <td width="67%"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                  </tr>
              </table></td>
            </tr>
          </table>
	  <textarea name="content" cols="" rows="" style="display:none"><%=ras("��Ʒ����")%></textarea>
      	  <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="340"></iframe>
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table>      </td>
    </tr>
    <tr>
      <td height="30" align="left" valign="top" background="images/bg_title.gif" style="padding-left:50">        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table>
        <input type="image" name="imageField" id="imageField" src="images/submit-bt.gif">
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
	set rs=server.createobject("adodb.recordset")
	sql="select * from peijian where id="&id
	rs.open sql,conn,2,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write("<script>alert('ɾ���ɹ�');location='?action=list';</script>")
end if
 %>
</body>
</html>                                                                             