

<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->

<!--#include file="Function_Page.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">

<style>




/*�б��ҳ*/
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
      <td align="center" bgcolor="#DAECF5" class="fontb">������Ƹ��Ϣ</td>
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
      <td height="0" colspan="3" align="right" ><%=i%>.ӦƸְλ��</td>
      <td width="84%" align="left" ><%=rs("Title")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="1" colspan="3" align="right" >������ </td>
      <td align="left" ><%=rs("renci")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="3" colspan="3" align="right" >�����ص㣺</td>
      <td align="left" ><%=rs("gzdd")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="7" colspan="3" align="right" >���ʴ����� </td>
      <td align="left" ><%=rs("gzdy")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="15" colspan="3" align="right" >��ֹ���ڣ�</td>
      <td align="left" ><%=rs("jzrq")%></td>
    </tr>
    
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >����ʱ�䣺</td>
      <td align="left" ><%=rs("Times")%></td>
    </tr>
    <tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td colspan="3" align="right" >��ƸҪ��</td>
      <td align="left" ><%if rs("content")<>"" then response.write cutstr(removehtml(rs("content")),100) else response.write "û������Ҫ��" end if%></td>
    </tr><tr bgcolor="#F9F9F9" onMouseOver="this.style.backgroundColor='#EAFCD5';" onMouseOut="this.style.backgroundColor='';this.style.color=''">
      <td height="24" colspan="3" align="right" >��ʾ����</td>
      <td align="left" ><input name="listt" type="text" class="form2" id="listt" value="<%=rs("listnum")%>" size="6" />
&nbsp;&nbsp;      <a href="?ss=edit&action=edit&amp;id=<%=rs("ID")%>&page=<%=page%>">�༭</a> | <a  onclick='{if(confirm("��ȷ��ɾ����?�˲��������ָܻ�!")){return true;}return false;}' href="?ac=del&amp;id=<%=rs("ID")%>&amp;page=<%=page%>">ɾ��</a>&nbsp;&nbsp;&nbsp;&nbsp;ѡ��
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
      <td align="center" bgcolor="#F9F9F9">ȫѡ/ȡ����
        <input name="Action" type="hidden"  value="Del" />
        <input name="chkAll" type="checkbox" id="chkAll" onclick=sel('chk') value="checkbox" style="border:0" />
        &nbsp;
        <input name="del" type="submit" class="admintable1" id="del" value="����" />
&nbsp;&nbsp;&nbsp;
<input name="Del" type="submit" class="admintable" id="Del" value="����ʱ��" />
&nbsp;&nbsp;
<input name="Del" type="submit" class="admintable1"  onclick='{if(confirm("��ȷ��ɾ����?�˲��������ָܻ�!")){return true;}return false;}'  id="Del" value="ɾ��" />
����ʱ�䡢����ɾ��������ѡȡ</td>
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
   set rs=conn.execute("update job set times = now() where ID In(" & ID & ")")
   Response.Write("<script>alert('�����ɹ�!');location='?ss=list'</script>")
ElseIf Request("Del")="����" Then
For i = 0 To UBound(idTmp)
  ' conn.execute("update job set list ="& listtTmp(i) &" where ID In(" & ID & ")")
    conn.execute("update job set listnum=" & listtTmp(i) & " where id=" & idTmp(i))
Next
 Response.Write("<script>alert('�����ɹ�!');location='?ss=list'</script>")

ElseIf Request("Del")="ɾ��" Then
	'set rs=conn.execute("delete from news where ID In(" & ID & ")")
			for i=1 to request("ID").count
				if request("ID").count=1 then
				newsID=request("ID")
				else
				newsID=replace(request("id")(i),"'","")
				end if
				
				'ɾ������
				Conn.Execute("Delete from [job] where ID = "&newsID&"")
				
			next
			response.write "<script>alert('ɾ���ɹ�!');location='?ss=list'</script>"
End If
End Sub
if trim(request.querystring("ac"))="del" then
ID=request("ID")
Conn.Execute("Delete from [job] where ID = "&ID&"")
response.write "<script>alert('ɾ���ɹ�!');location='?ss=list'</script>"
end if

%>

</div> 


<%end if%>









<%if ss="add" then%>






	<script language="javascript" type="text/javascript">
// ��֤�û���������
function check_add(){
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

<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>

<form id="form1" name="form1" method="post" action="?action=add" class="border-padd0" onSubmit="return check_add()" >
<table width="98%" height="32" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="13" bgcolor="#DAEDFC">&nbsp;</td>
    <td align="center" bgcolor="#DAEDFC" class="fontb">�����Ƹ��Ϣ</td>
    <td width="13" bgcolor="#DAEDFC">&nbsp;</td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#D7E3F2" class="border1">
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">��Ƹְλ:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="title" type="text" class="form1" id="title" size="30" maxlength="50" /><input name="TColor" type="hidden" id="TitleFontColor" onClick="Getcolor(ColorBG,'TitleFontColor');" Readonly></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">��Ƹ����: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="renci" type="text" class="form1" id="renci" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">�����ص�:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdd" type="text" class="form1" id="gzdd" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">���ʴ���:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdy" type="text" class="form1" id="gzdy" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">��ֹ����: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="jzrq" type="text" class="form1" id="jzrq" size="30" maxlength="50"  />
      </td>
    </tr>
  
    

    
    <tr>
      <td align="right" valign="top" bgcolor="#F9F9F9">��ƸҪ��</td>
      <td colspan="2" bgcolor="#F9F9F9">
<textarea name="content" style="display:none"></textarea>
		<IFRAME ID="qi500" SRC="../qi500@lm_webe/qi500@edit.htm?id=content&style=blue" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="500"></IFRAME></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#F9F9F9">����ʱ�䣺</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="times" type="text" class="form1" id="times" value="<%=now%>" size="30" /></td>
    </tr>

    <tr>
      <td colspan="3" align="center" bgcolor="#F9F9F9"><input type="submit" name="Submit" value="ȷ�����" />
      &nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="Submit2" value="�����д" /></td>
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
if trim(request.form("submit"))="ȷ�����" then

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
	
    response.write "<script>alert('��ӳɹ�');location='?ss=list'</script>"
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
    <td align="center" background="../images/bn_bg.jpg" class="fontb">�޸���Ƹ��Ϣ</td>
    <td width="13">&nbsp;</td>
  </tr>
</table>

  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#D7E3F2" class="border1">
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">��Ƹְλ:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="title" type="text" class="form1" id="title" value="<%=rs("title")%>" size="30" maxlength="50" /><input name="TColor" type="hidden" id="TitleFontColor" onClick="Getcolor(ColorBG,'TitleFontColor');" Readonly></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">��Ƹ����: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="renci" type="text" class="form1" id="renci" value="<%=rs("renci")%>" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">�����ص�:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdd" type="text" class="form1" id="gzdd" value="<%=rs("gzdd")%>" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">���ʴ���:</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="gzdy" type="text" class="form1" id="gzdy" value="<%=rs("gzdy")%>" size="30" maxlength="50" /></td>
    </tr>
    <tr>
      <td width="88" align="right" bgcolor="#F9F9F9">��ֹ����: </td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="jzrq" type="text" class="form1" id="jzrq" value="<%=rs("jzrq")%>" size="30" maxlength="50" /></td>
    </tr>
  
    

    
    <tr>
      <td align="right" valign="top" bgcolor="#F9F9F9">��ƸҪ��</td>
      <td colspan="2" bgcolor="#F9F9F9">
	  
	<textarea id="content" name="content" style="display:none;"><%=rs("content")%></textarea>
		<IFRAME ID="qi500" SRC="../qi500@lm_webe/qi500@edit.htm?id=content&style=blue" FRAMEBORDER="0" SCROLLING="no" WIDTH="100%" HEIGHT="500"></IFRAME></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#F9F9F9">����ʱ�䣺</td>
      <td colspan="2" bgcolor="#F9F9F9"><input name="times" type="text" class="form1" id="times" value="<%=rs("times")%>" size="30" /></td>
    </tr>

    <tr>
      <td colspan="3" align="center" bgcolor="#F9F9F9"><input type="submit" name="Submit" value="ȷ���޸�" />
      &nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="Submit2" value="�����д" /></td>
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
if trim(request.form("submit"))="ȷ���޸�" then

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
	
    response.write "<script>alert('�޸ĳɹ�');location='?ss=list'</script>"
end if
end if
%>




























</body>
</html>
