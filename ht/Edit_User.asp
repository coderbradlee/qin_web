<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="images/style.css" rel="stylesheet" type="text/css">


		
          <script>
function checkform(){
if(document.form.user_name.value==""){alert("�������û�����");document.form.user_name.focus();return false;}
if(document.form.user_truename.value==""){alert("��������ʵ������");document.form.user_truename.focus();return false;}
if(document.form.user_phone.value==""){alert("��������ϵ�绰��");document.form.user_phone.focus();return false;}
if(document.form.user_email.value==""){alert("������E-mail��ַ��");document.form.user_email.focus();return false;}



if (trim(document.form.country1.value) =="�й�") 
{ 
	if (trim(document.form.province.value) =="")
	{
	alert("��ѡ��ʡ�ݣ�"); 
	document.form.province.focus(); 
	return (false); 
	}
	if (trim(document.form.city.value) =="")
	{
	alert("��ѡ��ؼ����У�"); 
	document.form.city.focus(); 
	return (false); 
	}
} 

return true
}
</script>



</head>

<body>


<%
action=request.form("edit")
if action="edit" then
jid=request.Form("jid")
page=request.QueryString("page")
keywords=request.QueryString("keywords")
user_ename=request.QueryString("user_ename")
set ras=server.createobject("adodb.recordset")
wsql="select * from jiedai_User where id="&jid&""
ras.open wsql,conn,3,2
'ras("user_name")=Replace_Text(request.form("user_name"))
if request.form("user_pwd")<>"" then
ras("user_pwd")=request.form("user_pwd")
end if
ras("user_phone")=request.form("user_phone")
ras("user_truename")=request.form("user_truename")
ras("user_email")=request.form("user_email")
ras("user_address")=request.form("user_address")
ras("user_phone")=request.form("user_phone")
ras("user_question")=request.form("user_question")
ras("user_answer")=request.form("user_answer")

ras("User_country")=request.form("country1")
ras("User_province")=request.form("province")
ras("User_City")=request.form("city")

ras.update
ras.close:set ras=nothing
'session("jd_username")=request.form("user_name")
'session("jd_userid")=makefilename()
'session("jd_truename")=request.form("user_truename")
response.write"<script>alert('�޸ĳɹ���');location.href='user_Manage.asp?page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"';</script>"
response.end
end if
jid=request.QueryString("jid")
user_ename=request.QueryString("user_ename")
set res=server.createobject("adodb.recordset")
esql="select * from jiedai_User where id="&jid&""
res.open esql,conn,3,2


%>
          <table width="100%" height="410" border="0" cellpadding="4" cellspacing="1" style="font-size:12px;">
            <form name="form" method="post" OnSubmit="return checkform()" action="?page=<%=request.QueryString("page")%>&keywords=<%=request.QueryString("keywords")%>&user_ename=<%=request.QueryString("user_ename")%>">
              <tr class="black">
                <td width="20%" height="30" align="right">�����û���:</td>
                <td width="80%" height="30"><input name="user_name" type="text" class="inputb" id="user_name" size="20" maxlength="30" style="height:25px; padding-top:3px" value="<%=res("user_name")%>" disabled="disabled"> 
                  <font color="#FF0000">�����޸�</font> </td>
              </tr>
              <tr class="black">
                <td width="20%" height="30" align="right">��������:</td>
                <td width="80%" height="30"><input name="user_pwd" type="password" class="inputb" id="user_pwd" size="20" maxlength="30" style="height:25px; padding-top:3px" >
                  <font color="#FF0000">���޸�������</font></td>
              </tr>
              <tr class="black">
                <td width="20%" height="30" align="right">��ʵ����:</td>
                <td width="80%" height="30"><input name="user_truename" type="text" class="inputb" id="user_truename" style="height:25px; padding-top:3px" size="16" maxlength="24"  value="<%=res("user_truename")%>" >
                    <font color="#FF0000"> **</font></td>
              </tr>
              <tr class="black">
                <td width="20%" height="30" align="right">��ϵ�绰:</td>
                <td width="80%" height="30"><input name="user_phone" type="text" class="inputb" id="user_phone" style="height:25px; padding-top:3px"  value="<%=res("user_phone")%>"  >
                    <font color="#FF0000">**</font></td>
              </tr>
              <tr class="black">
                <td width="20%" height="30" align="right">��ϵ��ַ:</td>
                <td width="80%" height="30"><input name="user_address" type="text" class="inputb" id="user_address" style="height:25px; padding-top:3px"  value="<%=res("user_address")%>" ></td>
              </tr>
              <tr class="black">
                <td width="20%" height="30" align="right">E-mail:</td>
                <td width="80%" height="30"><input name="user_email" type="text" class="inputb" id="user_email" style="height:25px; padding-top:3px" size="30"  value="<%=res("user_email")%>" >
                    <font color="#FF0000"> **</font></td>
              </tr>
              <tr class="black">
                <td height="30" align="right">���ڵ���:</td>
                <td height="30">
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				<SELECT
            onchange=javascript:changeCountry1(this); name=country1 style="display:none">

                    </SELECT> <SELECT style="WIDTH: 60px" 
            onchange=changeProvince(this); name=province>
                      <OPTION value="" 
              selected>ʡ��</OPTION>
                    </SELECT> <SELECT name=city>
                      <OPTION value="" 
              selected>�ؼ�</OPTION>
                    </SELECT>
           <SCRIPT language=javascript src="../js/area.js"    
      purpose="INITIALIZER"></SCRIPT>
	  <SCRIPT language=javascript>
<!--
    //���ù���1��ȱʡֵ
    for (var i=0;i<5;i++){
	    if (country1Form1.options[i].value == '<%=res("User_country")%>'){
		    country1Form1.options[i].selected=true;
		}
	}

    for (var i=0;i<catArr1.length;i++) {
		catForm1.options[i+1]=new Option(catArr1[i].title,catArr1[i].id);
		//����ʡѡ����ѡ��ֵ
		if(catForm1.options[i+1].value == '<%=res("User_province")%>'){
	        	catForm1.options[i+1].selected=true;
	        }
	}
	changeProvince(catForm1);
	var len = boardForm1.options.length;
	for (var i=0;i<len;i++) {
		//���ó���ѡ����ѡ��ֵ
		if(boardForm1.options[i].value == '<%=res("User_City")%>') {
		    boardForm1.options[i].selected=true;
		}
	}
	
    if (country1Form1.options[country1Form1.selectedIndex].value!='�й�') {
	    catForm1.style.display = 'none';
	    boardForm1.style.display = 'none';
		}

-->
</SCRIPT>
                    
					
					
					
					
					
					
					
					
					
					
					
					
					
					
				  </td>
              </tr>
              <tr class="black">
                <td width="20%" height="30" align="right">������ʾ����:</td>
                <td width="80%" height="30"><input name="user_question" type="text" class="inputb" id="user_question" style="height:25px; padding-top:3px" size="20" maxlength="28"  value="<%=res("user_question")%>" >
                  ��Ҫ�һ������ʱ��,��ʾ�����⡰����ʲô���֣���</td>
              </tr>
              <tr class="black">
                <td width="20%" height="30" align="right">����ش�:</td>
                <td width="80%" height="30"><input name="user_answer" type="text" class="inputb" id="user_answer" style="height:25px; padding-top:3px" size="20" maxlength="28"  value="<%=res("user_answer")%>" >
                  ����������Ĵ�,�����Ĵ��ǡ�С�ơ�
                  <input name="edit" type="hidden" id="add" value="edit">
                  <input name="jid" type="hidden" id="jid" value="<%=request.querystring("jid")%>"></td>
              </tr>
              <tr class="black">
                <td height="30" colspan="2" align="center"><input name="Submit2" type="submit" id="Submit2" value=" �� �� " style="width:85px; height:35px">
                  
                  <input name="Submit2" type="button" id="Submit2" onClick="window.location='index.asp'" value=" ȡ �� "  style="width:85px; height:35px"></td>
              </tr>
            </form>
          </table>
      
</body>
</html>
