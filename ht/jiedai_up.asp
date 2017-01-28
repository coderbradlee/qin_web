<%@LANGUAGE="VBSCRIPT" CODEPAGE="936" %> 
<%Session.CodePage=936%>
<!--#include file="session.asp" -->
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<div id="lo" style="width:100%; height:100%; background:url(images/loading.gif) no-repeat left center; padding-left:18px; display:none; text-align:left">上传中...</div>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form action="saveupc.asp?ffs=<% =Trim(Request.QueryString("ffs")) %>&id=<% =Trim(Request.QueryString("id")) %>" method="post" enctype="multipart/form-data" name="form1">
  <tr>
    <td>  <input name="file" type="file" size="9">  
    <input type="submit" name="Submit" value="上传"  onclick="document.getElementById('lo').style.display=''" >

</td>
  </tr>
</form>
</table>
