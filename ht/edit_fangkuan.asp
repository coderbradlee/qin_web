<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<link href="images/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE43 {color: #3e689a;
	font-size: 12px;
	font-weight: normal;
	line-height: 20px;
}
.STYLE46 {color: #666666;
	font-size: 12px;
}
.STYLE53 {color: #FF3600;
	font-size: 12px;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 24px;
}
.STYLE56 {color: #3E689A; font-family: Arial, Helvetica, sans-serif; font-size: 14px; font-weight: bold; }
.STYLE60 {font-size: 12px; font-family: "����"; color: #333333; text-decoration: none; }
-->
</style>
</head>

<body>








<table width="100%" border="0" cellspacing="0" cellpadding="8">
            <tr>
              <td>
			  
			  
			  
			  
			            <script>
function checkform(){
if(document.form.jiekuan_Title.value==""){alert("����������⣡");document.form.jiekuan_Title.focus();return false;}
if(document.form.jiekuan_Money.value==""){alert("���������");document.form.jiekuan_Money.focus();return false;}
if(document.form.jiekuan_Interest.value==""){alert("��������������ʣ�");document.form.jiekuan_Interest.focus();return false;}

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





          <%
action=request.form("add")
if action="add" then
jid=request.form("jid")
page=request.QueryString("page")
keywords=request.QueryString("keywords")
user_ename=request.QueryString("user_ename")

set ras=server.createobject("adodb.recordset")
wsql="select * from jiedai_Fangkuan where id="&jid&""
ras.open wsql,conn,3,2
'ras("user_uid")=jd_userid
'ras("user_name")=jd_username
ras("Fangkuan_Title")=Replace_Text(request.form("jiekuan_Title"))
ras("Fangkuan_Money")=request.form("jiekuan_Money")
ras("Fangkuan_Time")=request.form("jiekuan_Time")
ras("Fangkuan_Interest")=request.form("jiekuan_Interest")
ras("Fangkuan_terval")=request.form("jiekuan_terval")
ras("Fangkuan_ExpiredTime")=request.form("jiekuan_ExpiredTime")
ras("Fangkuan_country")=request.form("country1")
ras("Fangkuan_province")=request.form("province")

ras("Fangkuan_city")=request.form("city")
ras("Fangkuan_dyw")=request.form("jiekuan_dyw")
ras("Fangkuan_Purpose")=request.form("jiekuan_Purpose")
ras("AssetsStatus")=request.form("AssetsStatus")
ras("user_truename")=Replace_Text(request.form("user_name"))
ras("user_phone")=Replace_Text(request.form("user_phone"))
if Replace_Text(request.form("user_phone_check"))="1" then
ras("user_phone_check")=true
end if
if Replace_Text(request.form("user_address_check"))="1" then
ras("user_address_check")=true
end if
ras("user_address")=Replace_Text(request.form("user_address"))

ras.update
ras.close:set ras=nothing
'session("jd_username")=request.form("user_name")
'session("jd_userid")=makefilename()
'session("jd_truename")=request.form("user_truename")
response.write"<script>alert('�޸ĳɹ���');location.href='Fangkuan_Manage.asp?page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"';</script>"
'response.Redirect("User_add_Jiekuan.asp")
response.end
end if

jid=Replace_Text(request.QueryString("jid"))
set ras=server.createobject("adodb.recordset")
wsql="select * from jiedai_Fangkuan where id="&jid&""
ras.open wsql,conn,3,2

%>

			  
			  
			  <form id="form" name="form" method="post" action="?page=<%=request.QueryString("page")%>&keywords=<%=request.QueryString("keywords")%>&user_ename=<%=request.QueryString("user_ename")%>" OnSubmit="return checkform()">

			  <table width="100%" height="221" border="0" cellpadding="5" cellspacing="0" class="zw">
			    <tr>
                  <td align="right"><span class="STYLE60">�Ŵ����⣺��</span></td>
			      <td><input name="jiekuan_Title" type="text" id="jiekuan_Title" size="30" value="<%=ras("Fangkuan_Title")%>" />
                      <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"><span id="ctl00_MainContent_lblName">��</span><span class="STYLE29">�����⽫��ʾ��Ҫλ�ã��õı���������������ṩ���</span></span></span></td>
			      </tr>
                <tr>
                  <td width="21%" align="right"><span class="STYLE60">�Ŵ�����</span></td>
                  <td width="79%"><input name="jiekuan_Money" type="text" id="jiekuan_Money" size="15"  value="<%=ras("Fangkuan_Money")%>" />
                    <span class="STYLE46"><span class="STYLE53">*</span></span><span class="STYLE53">��Ԫ�����</span> <span class="STYLE29">��������� ����Ľ���벻Ҫ��С���㡣</span></td>
                </tr>
                <tr>
                  <td align="right"><span class="STYLE60">�Ŵ����ޣ���</span></td>
                  <td><select name="jiekuan_Time" class="STYLE43" id="jiekuan_Time">
                    <option value="1" <%if ras("Fangkuan_Time")="1" then response.write"selected"%>>1����</option>
                    <option value="2" <%if ras("Fangkuan_Time")="2" then response.write"selected"%>>2����</option>
                    <option value="3" <%if ras("Fangkuan_Time")="3" then response.write"selected"%>>3����</option>
                    <option value="4" <%if ras("Fangkuan_Time")="4" then response.write"selected"%>>4����</option>
                    <option value="5" <%if ras("Fangkuan_Time")="5" then response.write"selected"%>>5����</option>
                    <option value="6" <%if ras("Fangkuan_Time")="6" then response.write"selected"%>>6����</option>
                    <option value="7" <%if ras("Fangkuan_Time")="7" then response.write"selected"%>>7����</option>
                    <option value="8" <%if ras("Fangkuan_Time")="8" then response.write"selected"%>>8����</option>
                    <option value="9" <%if ras("Fangkuan_Time")="9" then response.write"selected"%>>9����</option>
                    <option value="10" <%if ras("Fangkuan_Time")="10" then response.write"selected"%>>10����</option>
                    <option value="11" <%if ras("Fangkuan_Time")="11" then response.write"selected"%>>11����</option>
                    <option value="12" <%if ras("Fangkuan_Time")="12" then response.write"selected"%>>12����</option>
                    <option value="18" <%if ras("Fangkuan_Time")="18" then response.write"selected"%>>18����</option>
                    <option value="24" <%if ras("Fangkuan_Time")="24" then response.write"selected"%>>24����</option>
                    <option value="30" <%if ras("Fangkuan_Time")="30" then response.write"selected"%>>30����</option>
                    <option value="36" <%if ras("Fangkuan_Time")="36" then response.write"selected"%>>36����</option>
                  </select>
                    <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"></span></span></td>
                </tr>
                <tr>
                  <td align="right"><span class="STYLE60">��������ʣ���</span></td>
                  <td><input name="jiekuan_Interest" type="text" id="jiekuan_Interest" size="15"  value="<%=ras("Fangkuan_Interest")%>" />
                    <span class="STYLE60">%</span> <span class="STYLE46"><span class="STYLE53">*</span></span><br />
                    ����д�����ܽ��ܵ���������ʡ���ע�⣺�����ʲ��������շŴ��������ʣ������ܻ����Ŵ���ľ���Ͷ��������ӡ�</td>
                </tr>
                <tr>
                  <td align="right"><span class="STYLE60">���ʽ����</span></td>
                  <td><select name="jiekuan_terval" class="STYLE43" id="jiekuan_terval">
                      <option value="0" <%if ras("Fangkuan_terval")="0" then response.write"selected"%>>ÿ�»���</option>
                      <option value="1" <%if ras("Fangkuan_terval")="1" then response.write"selected"%>>���ڻ���</option>
                    </select>
                    <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"></span></span><br />
                    <span class="STYLE29">ÿ�»�����ָ�����ߴӽ����һ������ÿ���»���Ϣ�<br />
���ڻ�����ָ�������ڽ�����޵���һ�α�Ϣ����</span></td>
                </tr>
                <tr>
                  <td align="right"><span class="STYLE60">��ֹ���ڣ���</span></td>
                  <td><input name="jiekuan_ExpiredTime" type="text" id="jiekuan_ExpiredTime" size="15" value="<%=ras("Fangkuan_ExpiredTime")%>" />
                    <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"></span></span><br />
                    <span class="STYLE29">���ô˴ν������Ľ�ֹ���ڡ������ֹʱ����3����20��(����7��)</span></td>
                </tr>
              </table>
                <table width="100%" height="209" border="0" cellpadding="5" cellspacing="0" class="zw">
                  <tr>
                    <td width="21%" align="right">��Ѻλλ�ã�</td>
                    <td width="79%">
					
			
										   
				
				<SELECT
            onchange=javascript:changeCountry1(this); name=country1 style="display:none">
                    </SELECT> <SELECT style="WIDTH: 60px" 
            onchange=changeProvince(this); name=province>
                      <OPTION value="" 
              selected>ʡ��</OPTION>
                    </SELECT>
                    ʡ <SELECT name=city>
                      <OPTION value="" 
              selected>�ؼ�</OPTION>
                    </SELECT>
                    ��
                    <SCRIPT language=javascript src="../js/area.js"    
      purpose="INITIALIZER"></SCRIPT>
	  <SCRIPT language=javascript>
<!--
    //���ù���1��ȱʡֵ
    for (var i=0;i<5;i++){
	    if (country1Form1.options[i].value == '<%=ras("Fangkuan_country")%>'){
		    country1Form1.options[i].selected=true;
		}
	}

    for (var i=0;i<catArr1.length;i++) {
		catForm1.options[i+1]=new Option(catArr1[i].title,catArr1[i].id);
		//����ʡѡ����ѡ��ֵ
		if(catForm1.options[i+1].value == '<%=ras("Fangkuan_province")%>'){
	        	catForm1.options[i+1].selected=true;
	        }
	}
	changeProvince(catForm1);
	var len = boardForm1.options.length;
	for (var i=0;i<len;i++) {
		//���ó���ѡ����ѡ��ֵ
		if(boardForm1.options[i].value == '<%=ras("Fangkuan_City")%>') {
		    boardForm1.options[i].selected=true;
		}
	}
	
    if (country1Form1.options[country1Form1.selectedIndex].value!='�й�') {
	    catForm1.style.display = 'none';
	    boardForm1.style.display = 'none';
		}

-->
</SCRIPT>					  		</td>
                  </tr>
                  <tr>
                    <td align="right">��Ѻ��Ҫ��</td>
                    <td><select name="jiekuan_dyw" class="STYLE43" id="jiekuan_dyw">
                      <option value="0" <%if ras("fangkuan_dyw")="0" then response.write"selected"%>>������Ѻ</option>
                      <option value="1" <%if ras("fangkuan_dyw")="1" then response.write"selected"%>>������Ѻ</option>
                      <option value="2" <%if ras("fangkuan_dyw")="2" then response.write"selected"%>>����Ѻ</option>
                      <option value="3" <%if ras("fangkuan_dyw")="3" then response.write"selected"%>>������Ѻ</option>
                    </select></td>
                  </tr>
                  <tr>
                    <td align="right"><span class="STYLE60">��������</span></td>
                    <td><span class="STYLE46"><span class="STYLE53">
                    <textarea name="AssetsStatus" cols="60" rows="10" id="AssetsStatus"><%=ras("AssetsStatus")%></textarea>
                    </span></span></td>
                  </tr>
                </table>
                <table width="100%" height="85" border="0" cellpadding="5" cellspacing="0" class="zw">
                  <tr>
                    <td width="21%" align="right"><span class="STYLE60">��ϵ�ˣ���</span></td>
                    <td width="79%"><input name="user_name" type="text" id="user_name" value="<%=ras("user_truename")%>"></td>
                  </tr>
                  <tr>
                    <td align="right"><span class="STYLE60">�绰��</span>��</td>
                    <td><input name="user_phone" type="text" id="user_phone" value="<%=ras("user_phone")%>">
                      <input name="user_phone_check" type="checkbox" id="user_phone_check" value="1" <%if ras("user_phone_check")=true then response.write"checked"%>>
                      �Ƿ񹫿�<font color="#999999">(��ѡ��ʾ������Ĭ��Ϊ������)</font></td>
                  </tr>
                  <tr>
                    <td align="right">��<span class="STYLE60">��ϵ��ַ��</span>��</td>
                    <td><input name="user_address" type="text" id="user_address" value="<%=ras("user_address")%>">
                      <input name="user_address_check" type="checkbox" id="user_address_check" value="1" <%if ras("user_address_check")=true then response.write"checked"%>>
                      �Ƿ񹫿�<font color="#999999">(��ѡ��ʾ������Ĭ��Ϊ������)</font></td>
                  </tr>
                </table>
                <table width="100%" height="30" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td>&nbsp;</td>
                  </tr>
                </table>
                <table width="100%" height="50" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="200">&nbsp;</td>
                    <td><span class="STYLE9">
                      <input name="Submit222" type="submit" class="STYLE56" value=" �� ��  " style="width:85px; height:35px; cursor:hand" />
                    </span></td>
                    <td>&nbsp;</td>
                    <td><input name="add" type="hidden" id="add" value="add">
                      <input name="jid" type="hidden" id="jid" value="<%=request.QueryString("jid")%>"></td>
                    <td width="200">&nbsp;</td>
                  </tr>
                </table>
				
				
				</form>
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				
				</td>
            </tr>
          </table>












</body>
</html>
