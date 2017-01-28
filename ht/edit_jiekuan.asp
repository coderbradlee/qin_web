<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
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
.STYLE60 {font-size: 12px; font-family: "宋体"; color: #333333; text-decoration: none; }
-->
</style>
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="8">
  <tr>
    <td><script>
function checkform(){
if(document.form.jiekuan_Title.value==""){alert("请填入借款标题！");document.form.jiekuan_Title.focus();return false;}
if(document.form.jiekuan_Money.value==""){alert("请填入借款金额！");document.form.jiekuan_Money.focus();return false;}
if(document.form.jiekuan_Interest.value==""){alert("请填入最高年利率！");document.form.jiekuan_Interest.focus();return false;}

if (trim(document.form.country1.value) =="中国") 
{ 
	if (trim(document.form.province.value) =="")
	{
	alert("请选择省份！"); 
	document.form.province.focus(); 
	return (false); 
	}
	if (trim(document.form.city.value) =="")
	{
	alert("请选择地级城市！"); 
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
wsql="select * from jiedai_Jiekuan where id="&jid&""
ras.open wsql,conn,3,2
'ras("user_uid")=jd_userid
'ras("user_name")=jd_username
ras("jiekuan_Title")=Replace_Text(request.form("jiekuan_Title"))
ras("jiekuan_Money")=request.form("jiekuan_Money")
ras("jiekuan_Time")=request.form("jiekuan_Time")
ras("jiekuan_Interest")=request.form("jiekuan_Interest")
ras("jiekuan_terval")=request.form("jiekuan_terval")
ras("jiekuan_ExpiredTime")=request.form("jiekuan_ExpiredTime")
ras("jiekuan_country")=request.form("country1")
ras("jiekuan_province")=request.form("province")

ras("jiekuan_city")=request.form("city")
ras("jiekuan_dyw")=request.form("jiekuan_dyw")
ras("jiekuan_Purpose")=request.form("jiekuan_Purpose")
ras("AssetsStatus")=request.form("AssetsStatus")
ras("MyAssets")=request.form("MyAssets")

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
response.write"<script>alert('修改成功！');location.href='Jiekuan_Manage.asp?page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"';</script>"
'response.Redirect("User_add_Jiekuan.asp")
response.end
end if

jid=Replace_Text(request.QueryString("jid"))
set ras=server.createobject("adodb.recordset")
wsql="select * from jiedai_Jiekuan where id="&jid&""
ras.open wsql,conn,3,2

%>
        <form id="form" name="form" method="post" action="?page=<%=request.QueryString("page")%>&keywords=<%=request.QueryString("keywords")%>&user_ename=<%=request.QueryString("user_ename")%>" OnSubmit="return checkform()">
          <table width="100%" height="221" border="0" cellpadding="5" cellspacing="0" class="zw">
            <tr>
              <td align="right"><span class="STYLE60">借款标题：　</span></td>
              <td><input name="jiekuan_Title" type="text" id="jiekuan_Title" size="30" value="<%=ras("jiekuan_Title")%>" />
                  <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"><span id="ctl00_MainContent_lblName">　</span><span class="STYLE29">借款标题将显示显要位置，好的标题吸引更多的人提供贷款。</span></span></span></td>
            </tr>
            <tr>
              <td width="21%" align="right"><span class="STYLE60">借款金额：　</span></td>
              <td width="79%"><input name="jiekuan_Money" type="text" id="jiekuan_Money" size="15"  value="<%=ras("jiekuan_Money")%>" />
                  <span class="STYLE46"><span class="STYLE53">*</span></span><span class="STYLE53">万元人民币</span> <span class="STYLE29">请输入借款金额！ 输入的金额请不要有小数点。</span></td>
            </tr>
            <tr>
              <td align="right"><span class="STYLE60">借款期限：　</span></td>
              <td><select name="jiekuan_Time" class="STYLE43" id="jiekuan_Time">
                  <option value="1" <%if ras("jiekuan_Time")="1" then response.write"selected"%>>1个月</option>
                  <option value="2" <%if ras("jiekuan_Time")="2" then response.write"selected"%>>2个月</option>
                  <option value="3" <%if ras("jiekuan_Time")="3" then response.write"selected"%>>3个月</option>
                  <option value="4" <%if ras("jiekuan_Time")="4" then response.write"selected"%>>4个月</option>
                  <option value="5" <%if ras("jiekuan_Time")="5" then response.write"selected"%>>5个月</option>
                  <option value="6" <%if ras("jiekuan_Time")="6" then response.write"selected"%>>6个月</option>
                  <option value="7" <%if ras("jiekuan_Time")="7" then response.write"selected"%>>7个月</option>
                  <option value="8" <%if ras("jiekuan_Time")="8" then response.write"selected"%>>8个月</option>
                  <option value="9" <%if ras("jiekuan_Time")="9" then response.write"selected"%>>9个月</option>
                  <option value="10" <%if ras("jiekuan_Time")="10" then response.write"selected"%>>10个月</option>
                  <option value="11" <%if ras("jiekuan_Time")="11" then response.write"selected"%>>11个月</option>
                  <option value="12" <%if ras("jiekuan_Time")="12" then response.write"selected"%>>12个月</option>
                  <option value="18" <%if ras("jiekuan_Time")="18" then response.write"selected"%>>18个月</option>
                  <option value="24" <%if ras("jiekuan_Time")="24" then response.write"selected"%>>24个月</option>
                  <option value="30" <%if ras("jiekuan_Time")="30" then response.write"selected"%>>30个月</option>
                  <option value="36" <%if ras("jiekuan_Time")="36" then response.write"selected"%>>36个月</option>
                </select>
                  <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"></span></span></td>
            </tr>
            <tr>
              <td align="right"><span class="STYLE60">最高年利率：　</span></td>
              <td><input name="jiekuan_Interest" type="text" id="jiekuan_Interest" size="15"  value="<%=ras("jiekuan_Interest")%>" />
                  <span class="STYLE60">%</span> <span class="STYLE46"><span class="STYLE53">*</span></span><br />
                  <span class="STYLE29">请填写您所能接受的最高年利率。请注意：此利率并非您最终还款的年利率，它可能会随着贷款的竞争投标会 
                    逐渐降低。 </span></td>
            </tr>
            <tr>
              <td align="right"><span class="STYLE60">还款方式：　</span></td>
              <td><select name="jiekuan_terval" class="STYLE43" id="jiekuan_terval">
                  <option value="0" <%if ras("jiekuan_terval")="0" then response.write"selected"%>>每月还款</option>
                  <option value="1" <%if ras("jiekuan_terval")="1" then response.write"selected"%>>到期还款</option>
                </select>
                  <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"></span></span><br />
                  <span class="STYLE29">每月还款是指借入者从借入第一个月起每个月还本息款；<br />
                    到期还款是指借入者在借款期限到后一次本息还清</span></td>
            </tr>
            <tr>
              <td align="right"><span class="STYLE60">截止日期：　</span></td>
              <td><input name="jiekuan_ExpiredTime" type="text" id="jiekuan_ExpiredTime" size="15" value="<%=ras("jiekuan_ExpiredTime")%>" />
                  <span class="STYLE46"><span class="STYLE53">*</span><span id="ctl00_MainContent_lblName"></span></span><br />
                  <span class="STYLE29">设置此次借款请求的截止日期。建议截止时间是3天至20天(建议7天)</span></td>
            </tr>
          </table>
          <table width="100%" height="209" border="0" cellpadding="5" cellspacing="0" class="zw">
            <tr>
              <td align="right">资产：</td>
              <td><select name="MyAssets" class="STYLE43" id="MyAssets">
                  <option value="有房有车" <%if ras("MyAssets")="有房有车" then response.write"selected"%>>有房有车</option>
                  <option value="有房" <%if ras("MyAssets")="有房" then response.write"selected"%>>有房</option>
                  <option value="有车" <%if ras("MyAssets")="有车" then response.write"selected"%>>有车</option>
                  <option value="其他" <%if ras("MyAssets")="其他" then response.write"selected"%>>其他</option>
                </select>
              </td>
            </tr>
            <tr>
              <td width="21%" align="right">资产位置：</td>
              <td width="79%"><SELECT
            onchange=javascript:changeCountry1(this); name=country1 style="display:none">
                </SELECT>
                  <SELECT style="WIDTH: 60px" 
            onchange=changeProvince(this); name=province>
                    <OPTION value="" 
              selected>省份</OPTION>
                  </SELECT>
                省
                <SELECT name=city>
                  <OPTION value="" 
              selected>地级</OPTION>
                </SELECT>
                市
                <SCRIPT language=javascript src="../js/area.js"    
      purpose="INITIALIZER"></SCRIPT>
                <SCRIPT language=javascript>
<!--
    //设置国家1的缺省值
    for (var i=0;i<5;i++){
	    if (country1Form1.options[i].value == '<%=ras("jiekuan_country")%>'){
		    country1Form1.options[i].selected=true;
		}
	}

    for (var i=0;i<catArr1.length;i++) {
		catForm1.options[i+1]=new Option(catArr1[i].title,catArr1[i].id);
		//设置省选择框的选择值
		if(catForm1.options[i+1].value == '<%=ras("jiekuan_province")%>'){
	        	catForm1.options[i+1].selected=true;
	        }
	}
	changeProvince(catForm1);
	var len = boardForm1.options.length;
	for (var i=0;i<len;i++) {
		//设置城市选择框的选择值
		if(boardForm1.options[i].value == '<%=ras("jiekuan_City")%>') {
		    boardForm1.options[i].selected=true;
		}
	}
	
    if (country1Form1.options[country1Form1.selectedIndex].value!='中国') {
	    catForm1.style.display = 'none';
	    boardForm1.style.display = 'none';
		}

-->
</SCRIPT>
              </td>
            </tr>
            <tr>
              <td align="right">资产价值：</td>
              <td><input name="jiekuan_dyw" type="text" id="jiekuan_dyw" value="<%=ras("jiekuan_dyw")%>" ></td>
            </tr>
            <tr>
              <td align="right"><span class="STYLE60">借款目的：　</span></td>
              <td><span class="STYLE29">　请利用下面的空间尽可能地提供详细的信息，这些信息将成为借出人的重要参考；一般来说，如果对于信息<br />
                不全或信息有疑问，借出人不会借出。最少</span><span class="STYLE53">10</span><span class="STYLE29">字，最多</span><span class="STYLE53">500</span><span class="STYLE29">字。</span> <br />
                <textarea name="jiekuan_Purpose" cols="60" rows="5" id="jiekuan_Purpose"><%=ras("jiekuan_Purpose")%></textarea>
                <span class="STYLE46"><span class="STYLE53">*<br />
                </span></span></td>
            </tr>
            <tr>
              <td align="right"><span class="STYLE60">经济状况描述：　</span></td>
              <td><span class="STYLE46"><span class="STYLE53">
                <textarea name="AssetsStatus" cols="60" rows="12" id="AssetsStatus"><%=ras("AssetsStatus")%></textarea>
                <br>
              </span></span></td>
            </tr>
          </table>
          <table width="100%" height="85" border="0" cellpadding="5" cellspacing="0" class="zw">
            <tr>
              <td width="21%" align="right"><span class="STYLE60">联系人：　</span></td>
              <td width="79%"><input name="user_name" type="text" id="user_name" value="<%=ras("user_truename")%>"></td>
            </tr>
            <tr>
              <td align="right"><span class="STYLE60">电话：</span>　</td>
              <td><input name="user_phone" type="text" id="user_phone" value="<%=ras("user_phone")%>">
                  <input name="user_phone_check" type="checkbox" id="user_phone_check" value="1" <%if ras("user_phone_check")=true then response.write"checked"%>>
                是否公开<font color="#999999">(勾选表示公开，默认为不公开)</font></td>
            </tr>
            <tr>
              <td align="right">　<span class="STYLE60">联系地址：</span>　</td>
              <td><input name="user_address" type="text" id="user_address" value="<%=ras("user_address")%>">
                  <input name="user_address_check" type="checkbox" id="user_address_check" value="1" <%if ras("user_address_check")=true then response.write"checked"%>>
                是否公开<font color="#999999">(勾选表示公开，默认为不公开)</font></td>
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
                <input name="Submit222" type="submit" class="STYLE56" value=" 提 交  " style="width:85px; height:35px; cursor:hand" />
              </span></td>
              <td>&nbsp;</td>
              <td><input name="add" type="hidden" id="add" value="add">
                  <input name="jid" type="hidden" id="jid" value="<%=request.QueryString("jid")%>"></td>
              <td width="200">&nbsp;</td>
            </tr>
          </table>
        </form></td>
  </tr>
</table>
</body>
</html>
