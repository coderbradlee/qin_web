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
if trim(request.querystring("zhiding"))="zdyes" then
	id=trim(request.querystring("jid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	
		set rs=server.createobject("adodb.recordset")
		sql="select * from product where id="&id
		rs.open sql,conn,1,3
		'rs.addnew
		
		rs("�Ƿ��Ƽ�")="��"	
	
	rs.update
	'rs.requery
	rs.close
	set rs=nothing	

	response.write "<script>alert('��ȡ����ҳ��ʾ!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
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
		sql="select * from product where id="&id
		rs.open sql,conn,1,3
		'rs.addnew
		
		rs("�Ƿ��Ƽ�")="��"
	rs.update
'	rs.requery
	rs.close
	set rs=nothing

	response.write "<script>alert('������ҳ��ʾ�ɹ�!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"'</script>"
	response.end
end if
 %>
 
 













<%
if trim(request.querystring("action"))="list" then












OrderBy = request("OrderBy") '��ǰ��¼��������� at_no ��at_id�Ǽ�¼��id ���������id������Ƿֿ��� 
at_ID = request("at_ID") '��ǰ��¼id 
action = request("action") '�ƶ����� 
paixu=request("paixu")




if paixu="up" then '���� 
		'�����ж��ǲ����Ѿ��ƶ�����ǰ 
		sql="select top 1 id from product where jhpx<"&OrderBy&" order by jhpx desc" 
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
		sql2="select jhpx from product where id="&before_id '��ѯǰһ����¼ 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		before_Orderby=rs2("jhpx") '��¼�ü�¼��������� 
		rs2("jhpx")=OrderBy '���ĸü�¼��������� 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select jhpx from product where ID="&at_ID '��ѯ��ǰ��¼��Ҫ�ƶ��ļ�¼�� 
		set rs3=server.createobject("adodb.recordset") 
		rs3.open sql3,conn,1,3 
		rs3("jhpx")=before_Orderby '�޸ĵ�ǰ��¼���������Ϊǰһ��¼����ţ����� ������¼��Ž��� ʵ������ 
		rs3.update 
		rs3.close 
		set rs3=nothing 

elseif paixu="Down" then '���������Ƶ���һ�� 
		sql="select top 1 id from product where jhpx>"&OrderBy&" order by jhpx asc" 
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
		
		sql2="select jhpx from product where ID="&end_id 
		set rs2=server.createobject("adodb.recordset") 
		rs2.open sql2,conn,1,3 
		end_Orderby=rs2("jhpx") 
		rs2("jhpx")=OrderBy 
		rs2.update 
		rs2.close 
		set rs2=nothing 
		
		sql3="select jhpx from product where ID="&at_ID 
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
        <td valign="top"></td>
        <td valign="top">     <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" bgcolor="#D3E5FA">
            <tr>
              <td style="padding-left:15"><strong><font color="#215dc6">��Ϣ����</font></strong> </td>
            </tr>
          </table>
          <table width="100%" height="40" border="0" cellpadding="0" cellspacing="0" style="display:none">
            <tr>
              <td width="2%">&nbsp;</td>
              <td width="98%">
			  
			  ��Ʒ���ͣ�

			  
			  
<% 	
sql="select * from sh_sort order by anclassidorder"
set ras=conn.execute(sql)  
 %>

				  
                    <% do while not ras.eof %>
<a href="?Caseid=<%=ras("anclassid")%>&action=list" style="font-weight:bold;"><%= ras("anclass") %></a><%if ras("e_anclass")<>"" then 
		   response.Write"<img src=""images/en.jpg"" />" 
		   else
		   response.Write"<img src=""images/noen.jpg"" />" 
		   end if
		   %> &nbsp;&nbsp; | &nbsp;&nbsp;
                    <%
		 ras.movenext
		loop
		ras.close
		set ras=nothing
		 %>
			  
			  
			  
			  </td>
            </tr>
          </table>
          <br>
          <table width="100%" height="138" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td valign="top">
			  
			  
			  
			  
			  
			  
			 <% 

	on error resume next
   toutiao=request("t")
set rs=server.createobject("adodb.recordset")
	Caseid=trim(request.querystring("Caseid"))
	Nclassid=trim(request.querystring("Nclassid"))
	sql="select * from product where 1=1 "
		
		if Caseid<>"" then
		sql=sql+" and anclassid="&Caseid&""
		end if
		
		if Nclassid<>"" then
		sql=sql+" and nclassid="&Nclassid&""
		end if
		
		sql=sql+"  order by id desc"

	rs.open sql,conn,1,1
	rs.pagesize=8
	
	if trim(request.querystring("page")<>"") then 
	
	
	
			if isnumeric(trim(request.querystring("page")))=false then
			page=1
			else
			page=cint(trim(request.querystring("page")))
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
do while i<4 and not rs.eof  and rowcount>0
%>






                              <td width="175">
							  
							  
							  
                              
                              
                              
                              
                              
                     <table width="162" border="0" cellpadding="0" cellspacing="0" style="margin:3px 5px;">
         <tr>
           <td bgcolor="#D3E9FC"><img src="../uploadfile/<%=rs("��ƷͼƬ")%>" height="145" width="162" border="0"  style="padding:1px; border:1px solid #B6D7EF; "></td>
         </tr>
         <tr>
           <td height="28" bgcolor="#D3E9FC"  align="center"><%if rs("e_title")<>"" then 
		   response.Write"<img src=""images/en.jpg"" />" 
		   else
		   response.Write"<img src=""images/noen.jpg"" />" 
		   end if
		   %><span class="line"><%= rs("��Ʒ����") %></span></td>
         </tr>
         <tr>
           <td height="28" align="center" bgcolor="#F3F3F3">&nbsp;&nbsp; <a href="?action=del&id=<%= rs("id") %>&Caseid=<%=Caseid%>&Nclassid=<%=Nclassid%>" onClick="return window.confirm('ȷ��ɾ����?');">[ɾ��]</a> &nbsp;&nbsp; <a href="?action=edit&id=<%= rs("id") %>&Caseid=<%=Caseid%>&Nclassid=<%=Nclassid%>">[�޸�]</a></td>
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
		sql="select * from product"
		rs.open sql,conn,1,3
		rs.addnew
		rs("��ƷͼƬ")=request("image")
		rs("�Ƿ��Ƽ�")=request("tuijian")
		if request.form("flag")<>"" then
		rs("jhpx")=request.form("flag")
		else
		rs("jhpx")=100
		end if
		rs("��Ʒ����")=request("content")		
		rs("��Ʒ����")=request("title") 
		rs("anclassid")=int(request("anclassid")) '����
rs("nclassid")=int(request("nclassid")) 'С��	

rs("e_content")=request("e_content")
rs("e_title")=request("e_title")

rs("cptz")=request("cptz")
rs("e_cptz")=request("e_cptz")
rs("yyly")=request("yyly")
rs("e_yyly")=request("e_yyly")
rs("baozhuang")=request("baozhuang")
rs("e_baozhuang")=request("e_baozhuang")
rs("cpyy")=request("cpyy")
rs("e_cpyy")=request("e_cpyy")
rs("fjbz")=request("fjbz")
rs("e_fjbz")=request("e_fjbz")
rs("cpbz")=request("cpbz")
rs("e_cpbz")=request("e_cpbz")
	
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
            <td width="132" align="center" height="40">�������ƣ�</td>
            <td><input name="title" type="text" id="title" size="40" style="height:20; width:400"></td>
          </tr>
          
          
		    <tr>
            <td width="132" align="center"  height="40">Ӣ�����ƣ�</td>
            <td><input name="e_title" type="text" id="e_title" size="40" style="height:20; width:100%"></td>
          </tr>
	
	    <tr> 
      <td width="132" height="40" align="center">����:</td>
      <td width="861"> 
      ������:
      <input name="flag" type="text" id="flag" size="8"></td>
    </tr>
	
	
	
    <tr>
      <td height="40" align="center">ͼƬ:</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="25%"><input name="image" type="text" id="image" size="40" style="height:25"></td>
            <td width="75%"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
          </tr>
      </table></td>
    </tr>
    <tr height="30" style="display:none" >
      <td height="40" align="center">����:</td>
      <td>
	  
	  
	  
	  <%
	  set rs=server.CreateObject("adodb.recordset")
     rs.open "select * from sh_sort order by anclassidorder",conn,1,1
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
          <option value="<%=rs("anclassid")%>"><%=trim(rs("anclass"))%>  [[<%=trim(rs("e_anclass"))%>]]</option>
          <%
        rs.movenext
        loop
		end if
        rs.close
	%>
        </select>
	  
	  
      </td>
    </tr>
  
  
  
  
		  <tr>
            <td align="center">��Ʒ���ܣ��У�</td>
            <td>
			 <textarea name="cptz" id="cptz" cols="45" rows="5"  style="display:none"></textarea><iframe id="ewebeditor1" src="../qi500@lm_webe/qi500@edit.htm?id=cptz&style=blue" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
  
  
  
		  <tr>
            <td align="center">��Ʒ���ܣ�Ӣ��</td>
            <td>
			 <textarea name="e_cptz" id="e_cptz" cols="45" rows="5"  style="display:none"></textarea><iframe id="ewebeditor1" src="../qi500@lm_webe/qi500@edit.htm?id=e_cptz&style=blue1" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
  
  
  
  <%sub cn_zb()%>
  <TABLE cellSpacing=1 cellPadding=0 width=407 bgColor=#e6e6e6 border=0>
<TBODY>
<TR>
<TD class=font vAlign=top align=left bgColor=#ffffff colSpan=2 height=20>
<P>&nbsp;<IMG height=12 src="images/dot3.jpg" 
width=12>����ָ��</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;�������Ѻ�����%</P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>��93&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;��������g/100g</P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>��20&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;ɸ����(45��mɸ��)��%</P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>��0.01&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;��ɫ����R-930������ȣ�<BR>&nbsp;Ŀ�ӷ�<BR>&nbsp;ɫ�ȷ�����</P></TD>
<TD class=font vAlign=bottom align=right width=158 bgColor=#ffffff height=20>
<P align=right>������&nbsp;<BR>��0.2&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;���ɢ����(��R-930������)��%</P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>96&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;��ɫ������ŵ������</P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>��1800&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;ˮ��ȡҺ�����ʣ�����m�� </P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>��200&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;105��ӷ��% </P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>�� 0.5&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;��23��2�漰���ʪ�ȣ�50��5����%<BR>&nbsp;Ԥ����24h��105��ӷ��% </P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right><BR>�� 1.5&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;ˮ���% </P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>��0.2&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;ˮ����ҺpHֵ</P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>6.5��8.0&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;��ĥ��ɢ�ԣ��ڸ�����H�� </P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>��6.25&nbsp;</P></TD></TR>
<TR>
<TD class=font vAlign=top align=left width=246 bgColor=#ffffff height=20>
<P>&nbsp;�߽���ɢ�ԣ���m��</P></TD>
<TD class=font vAlign=top align=right width=158 bgColor=#ffffff height=20>
<P align=right>�� 30&nbsp;</P></TD></TR></TBODY></TABLE>

<%end sub%>

	
		  <tr>
            <td align="center">����ָ�꣨�У�</td>
            <td>
			 <textarea name="content" id="content" cols="" rows="" style="display:none"><%call cn_zb%></textarea><iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
	<%sub en_zb()%>	 
    <table width="407" border="0" cellpadding="0" cellspacing="1" bgcolor="#E6E6E6">
              <tbody>
                <tr>
                  <td height="20" colspan="2" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;<img width="12" height="12" src="img/cpyfw_clip_image001.jpg" />Specifications</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Titanium dioxide content,%</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&ge;93&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Oil absorption, g/100g</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&le;20&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Residue on sieve (45&mu;m sieve),%</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&le;0.01&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Color (with R-930 compared with standard sample)<br />
                    &nbsp;Visual method<br />
                    &nbsp;Law &Delta;&Epsilon; chromaticity</p></td>
                  <td width="158" height="20" align="right" valign="bottom" bgcolor="#FFFFFF"><p align="right">Similar&nbsp;<br />
                    &le;0.2&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Relative scattering power (with the R-930 standard sample),%</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">96&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Reducing power (Reynolds number)</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&ge;1800&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Water extract resistivity (&Omega; &middot; m)</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&ge;200&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;105 �� volatile matter,%</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&le; 0.5&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;After 23 &plusmn; 2 �� and relative humidity (50 &plusmn; 5),%<br />
                    &nbsp;After 24h pretreatment, 105 �� Volatile,%</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right"><br />
                    &le; 1.5&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Water soluble,%</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&le;0.2&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;PH value of aqueous suspension</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">6.5��8.0&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;Grinding and dispersion of (Hagerman number of H)</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&ge;6.25&nbsp;</p></td>
                </tr>
                <tr>
                  <td width="246" height="20" align="left" valign="top" bgcolor="#FFFFFF"><p>&nbsp;High stirring dispersion (&mu;m)</p></td>
                  <td width="158" height="20" align="right" valign="top" bgcolor="#FFFFFF"><p align="right">&le; 30&nbsp;</p></td>
                </tr>
              </tbody>
            </table>
            
            <%end sub%> 
		  	  
		  <tr>
            <td align="center">����ָ�꣨Ӣ��</td>
            <td>
			 <textarea name="e_content" id="e_content" cols="" rows="" style="display:none"><%call en_zb%></textarea><iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
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
        </table><%end if%><input type="image" name="imageField2" src="images/submit-bt.gif"><br>
<br>

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
		tuijian=trim(request.form("tuijian"))
		Caseid=trim(request.QueryString("Caseid"))
		Nclassid=trim(request.QueryString("Nclassid"))
		
		 scontent =request.form("content")
		set rs=server.createobject("adodb.recordset")
		sql="select * from product where id="&id
		rs.open sql,conn,1,3
'		rs("classid")=classid
		rs("��Ʒ����")=scontent
		rs("�Ƿ��Ƽ�")=tuijian
		if request.form("flag")<>"" then
		rs("jhpx")=request.form("flag")
		else
		rs("jhpx")=100
		end if
		
		rs("��Ʒ����")=request("title")
		rs("��ƷͼƬ")=request("image")
		rs("anclassid")=int(request("anclassid")) '����
rs("nclassid")=int(request.form("nclassid")) 'С��
		rs("e_content")=request("e_content")
rs("e_title")=request("e_title")


rs("cptz")=request("cptz")
rs("e_cptz")=request("e_cptz")
rs("yyly")=request("yyly")
rs("e_yyly")=request("e_yyly")
rs("baozhuang")=request("baozhuang")
rs("e_baozhuang")=request("e_baozhuang")
rs("cpyy")=request("cpyy")
rs("e_cpyy")=request("e_cpyy")
rs("fjbz")=request("fjbz")
rs("e_fjbz")=request("e_fjbz")
rs("cpbz")=request("cpbz")
rs("e_cpbz")=request("e_cpbz")

		rs.update
		rs.requery
		rs.close
		set rs=nothing
		
		'response.Redirect("Jiedai_dypm.asp?Action=list&&Caseid="&Caseid&"&Nclassid="&Nclassid&"")
		response.write "<script>alert('�޸ĳɹ�!');location='Jiedai_dypm.asp?Action=list&&Caseid="&Caseid&"&Nclassid="&Nclassid&"'</script>"
	end if
	
		id=trim(request.querystring("id"))
		sql="select * from product where id="&id
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
      	  <table width="100%" height="53" border="0" cellpadding="0" cellspacing="1" bordercolor="#cccccc">
           
		   
		   
		       <tr>
            <td height="40" align="center" bgcolor="#F0F0F0">�������ƣ�</td>
            <td><input name="title" type="text" id="title" size="40" style="height:20; width:100% "   value="<%=ras("��Ʒ����")%>"></td>
          </tr>
          
          
		    <tr>
            <td width="217"  height="40" align="center" bgcolor="#F0F0F0">Ӣ�����ƣ�</td>
            <td><input name="e_title" type="text" id="e_title" size="40" style="height:20; width:100%"  value="<%=ras("e_title")%>"></td>
          </tr>
		   
		   
		   
		   
		    <tr>
              <td width="217" height="40" align="center" bgcolor="#F0F0F0">����:</td>
              <td width="774" height="40">
������:
  <input name="flag" type="text" id="flag" value="<%=ras("jhpx")%>" size="8">   ��</td>
            </tr>
			
			
			
			
			
			
            <tr  style="display:none">
              <td height="40" align="center" bgcolor="#F0F0F0">����:</td>
              <td height="40">
			  
			  
			  
	  
			  
			  <%dim rs1
	  set rs=server.CreateObject("adodb.recordset")
			   		set rs1=server.CreateObject("adodb.recordset")
					rs1.open "select * from product where id="&id,conn,1,1
					rs.open "select * from sh_sort order by anclassidorder",conn,1,1
					if rs.eof and rs.bof then
					response.write "���������Ŀ��"
					response.end
					else
				%>
        <select name="anclassid" size="1" id="anclassid" onChange="changelocation(document.myform.anclassid.options[document.myform.anclassid.selectedIndex].value)">
          <%do while not rs.eof%>
          <option value="<%=rs("anclassid")%>" <%if rs1("anclassid")=rs("anclassid") then%>selected<%end if%>><%=trim(rs("anclass"))%>   [[<%=trim(rs("e_anclass"))%>]]</option>
          <%
					rs.movenext
					loop
					end if
					rs.close
				%>
        </select>
		
		
		
		
		</td>
            </tr>
            <tr>
              <td height="40" align="center" bgcolor="#F0F0F0">ͼƬ:</td>
              <td height="40"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="33%"><input name="image" type="text" id="image" style="height:24" value="<%=ras("��ƷͼƬ")%>" size="40"></td>
                    <td width="67%"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                  </tr>
              </table></td>
            </tr>
            
            
            
            
            	  <tr>
            <td align="center">��Ʒ���ܣ��У�</td>
            <td>
			 <textarea name="cptz" id="cptz" cols="45" rows="5"  style="display:none"><%=ras("cptz")%></textarea><iframe id="ewebeditor1" src="../qi500@lm_webe/qi500@edit.htm?id=cptz&style=blue" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
            
            
  
		  <tr>
            <td align="center">��Ʒ���ܣ�Ӣ��</td>
            <td>
			 <textarea name="e_cptz" id="e_cptz" cols="45" rows="5"  style="display:none"><%=ras("e_cptz")%></textarea><iframe id="ewebeditor1" src="../qi500@lm_webe/qi500@edit.htm?id=e_cptz&style=blue1" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  


	
		  <tr>
            <td align="center">����ָ�꣨�У�</td>
            <td>
			 <textarea name="content" id="content" cols="" rows="" style="display:none"><%=ras("��Ʒ����")%></textarea><iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
  
		  <tr>
            <td align="center">����ָ�꣨Ӣ��</td>
            <td>
			 <textarea name="e_content" id="e_content" cols="" rows="" style="display:none"><%=ras("e_content")%></textarea><iframe id="ewebeditor2" src="<%=webeden%>" frameborder="0" scrolling="no" width="100%" height="300"></iframe>			</td>
          </tr>
		  
		  
            
            
            
            
          </table>

		
		
		
		
		
		
		
		  </td>
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
	sql="select * from product where id="&id
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