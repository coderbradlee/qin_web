<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->

<!--#include file="Function_Page.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE2 {color: #E8F9ED}






/*�б��ҳ*/
	#page {width:100%;padding:5px 0;}
	#page ul{margin:0 auto; display:table;text-align:center;}
	#page li{float:left !important; float:none;margin-right: 2px; height:17px;line-height:17px;display:inline; zoom:1;}
	#page span{display: block; padding: 2px 5px; background: #F5FBFF; border: 1px solid #CCC; color: #999999; }
	#page a{font-size:12px;display: block; text-decoration: none; margin:0px; color: #ffffff;padding: 2px 5px 2px 5px;background:#91CF40;}
	#page a:link,#page a:visited {border: 1px solid #CCCCCC; }
	#page a:hover {background:#F5FBFF;color:#135C86;}
	#page #span1 {background:#FFFFFF;display: block;}
-->
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
thestr=trim(thestr) '�����ַ���ǰ��Ŀհ�
thestr_length= len(thestr) '���ַ����ĳ���
if thestr_length > strlen then   '�ж��ַ����ĳ���
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
      <td width="13" bgcolor="#DBF5FD">&nbsp;</td>
      <td align="center" bgcolor="#DBF5FD" class="fontb">����ӦƸ��Ϣ</td>
      <td width="13" bgcolor="#DBF5FD">&nbsp;</td>
    </tr>
  </table>
     <%
listnum=request("listnum")
Set mypage=new xdownpage
mypage.getconn=conn
mysql="select * from job1"
mysql=mysql&" where id order by times desc"
mypage.getsql=mysql
mypage.pagesize=2
set rs=mypage.getrs()
for i=1 to mypage.pagesize
    if not rs.eof then 
%> 
 <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#C7D5E0">
    

    <tr bgcolor="#E2FCF0" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="0" colspan="3" align="right" ><%=i%>.ӦƸְλ��</td>
      <td width="78%" align="left" ><%=rs("Title")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="1" colspan="3" align="right" >�� ���� </td>
      <td align="left" ><%=rs("tname")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="3" colspan="3" align="right" >�� ��</td>
      <td align="left" ><%=rs("sex")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="7" colspan="3" align="right" >�������£� </td>
      <td align="left" ><%=rs("sr")%></td>
    </tr>
	   <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="7" colspan="3" align="right" >���᣺ </td>
      <td align="left" ><%=rs("jg")%></td>
    </tr>
	
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="15" colspan="3" align="right" >��ϵ�绰��</td>
      <td align="left" ><%=rs("tel")%></td>
    </tr>
    
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >��ϵ���䣺</td>
      <td align="left" ><%=rs("email")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >��סַ��</td>
      <td align="left" ><%=rs("add")%></td>
    </tr>
	    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >ѧ����</td>
      <td align="left" ><%=rs("xl")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >�������飺</td>
      <td align="left" ><%=rs("jy")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >ӦƸʱ�䣺</td>
      <td align="left" ><%=rs("times")%></td>
    </tr>
	   <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >˵����</td>
      <td align="left" ><%if rs("body")<>"" then response.write rs("body") else response.write "û������Ҫ��" end if%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >���˼�����</td>
      <td align="left" ><%if rs("content")<>"" then response.write cutstr(removehtml(rs("content")),100) else response.write "û������Ҫ��" end if%></td>
    </tr>
    
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="24" colspan="3" align="right" >������</td>
      <td align="left" >&nbsp;&nbsp;      <a href="edit.asp?action=edit&amp;id=<%=rs("ID")%>&page=<%=page%>">�༭</a> | <a  onclick='{if(confirm("��ȷ��ɾ����?�˲��������ָܻ�!")){return true;}return false;}' href="?ac=del&amp;id=<%=rs("ID")%>&amp;page=<%=page%>">ɾ��</a>&nbsp;&nbsp;&nbsp;&nbsp;ѡ��
        <input type="checkbox" value="<%=rs("ID")%>" name="ID" id="chk" style="border:0;" />
        <input name="ID2" type="hidden" id="ID2" value="<%=rs("ID")%>" /></td>
    </tr>
  </table>
  <table width="80%" height="10" border="0" align="center" cellpadding="0" cellspacing="0">
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
      <td align="center" bgcolor="#F9F9F9">ȫѡ/ȡ����
        
        <input name="chkAll" type="checkbox" id="chkAll" onclick=sel('chk') value="checkbox" style="border:0" />
        &nbsp;&nbsp;&nbsp;
<input name="Del" type="submit" class="admintable1"  onclick='{if(confirm("��ȷ��ɾ����?�˲��������ָܻ�!")){return true;}return false;}'  id="Del" value="ɾ��" />
������ѡȡ</td>
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
ids = Request.Form("id2") '��ȡͬ���ؼ���ֵ���磺1,2,3,4,5
listt = Request.Form("listt") 'Ҳ�ǻ�ȡͬ���ؼ���ֵ����������ģ�9,10,11,12,13
listtTmp = Split(listt,",")
idTmp = Split(ids,",")

If ID="" and Request("Del")<>"����"  Then
	  Response.Write("<script language=javascript>alert('��û��ѡ��!');history.back(1);</script>")
	  Response.End
ElseIf Request("Del")="����ʱ��" Then
   set rs=conn.execute("update job1 set times = now() where ID In(" & ID & ")")
   Response.Write("<script>alert('�����ɹ�!');location='list.asp'</script>")
ElseIf Request("Del")="����" Then
For i = 0 To UBound(idTmp)
  ' conn.execute("update job set list ="& listtTmp(i) &" where ID In(" & ID & ")")
    conn.execute("update job1 set listnum=" & listtTmp(i) & " where id=" & idTmp(i))
Next
 Response.Write("<script>alert('�����ɹ�!');location='list1.asp'</script>")

ElseIf Request("Del")="ɾ��" Then
	'set rs=conn.execute("delete from news where ID In(" & ID & ")")
			for i=1 to request("ID").count
				if request("ID").count=1 then
				newsID=request("ID")
				else
				newsID=replace(request("id")(i),"'","")
				end if
				
				'ɾ������
				Conn.Execute("Delete from [job1] where ID = "&newsID&"")
				
			next
			response.write "<script>alert('ɾ���ɹ�!');location='list1.asp'</script>"
End If
End Sub
if trim(request.querystring("ac"))="del" then
ID=request("ID")
Conn.Execute("Delete from [job1] where ID = "&ID&"")
response.write "<script>alert('ɾ���ɹ�!');location='list1.asp'</script>"
end if

%>

</div> 

</body>
</html>
