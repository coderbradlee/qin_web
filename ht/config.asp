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

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td align="center" valign="top">
<% 
	if trim(request("action"))="edit" then
		id=trim(request.querystring("id"))
		webname=trim(request.form("webname"))
		webkeyword=trim(request.form("webkeyword"))
		webdes=trim(request.form("webdes"))
		webicp=trim(request.form("webicp"))
		bq=trim(request.form("content"))
		set rs=server.createobject("adodb.recordset")
		sql="select * from web_config where id=1"
		rs.open sql,conn,1,3
'		rs("classid")=classid
		rs("webname")=webname
		rs("webkeyword")=webkeyword
		rs("webdes")=webdes
		
		rs("e_webname")=trim(request.form("e_webname"))
		rs("e_webkeyword")=trim(request.form("e_webkeyword"))
		rs("e_webdes")=trim(request.form("e_webdes"))
		
		rs("webicp")=webicp
		rs("bq")=bq
		rs("image")=request("image")
		rs.update
		rs.requery
		rs.close
		set rs=nothing
		
		
		
	end if
	
		id=1
		sql="select * from web_config where id="&id
		set rs=conn.execute(sql)

%>
<script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.content1.value==""){
		alert("内容不能为空！");
		document.form1.content1.focus();
		notnull=false;
		}
		return notnull;
	}
</script>
<form name="form1" method="post" action="?action=edit&id=1" onSubmit="return check_edit()">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
    <tr>
      <td height="20" bgcolor="#D3E5FA" style="padding-left:15"><b></b>&nbsp;
        </td>
      </tr>
    <tr>
      <td valign="top">
      
      	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="5"></td>
            </tr>
          </table>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table>      
	  <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td align="right">中文网站名称：</td>
          <td><input name="webname" type="text" id="webname" style="width:100%" size="30" value="<%=rs("webname")%>"></td>
        </tr>
        <tr>
          <td align="right">中文关键字：</td>
          <td><textarea name="webkeyword" cols="40" rows="6" id="webkeyword" style="height:60; width:100%"><%=rs("webkeyword")%></textarea></td>
        </tr>
        <tr>
          <td width="107" align="right">中文网站说明：</td>
          <td><textarea name="webdes" cols="40" rows="6" id="webdes" style="height:60; width:100%"><%=rs("webdes")%></textarea></td>
        </tr>
		
		
		
		
		
		
		
		 <tr>
          <td align="right">英文名称：</td>
          <td><input name="e_webname" type="text" id="e_webname" style="width:100%" size="30" value="<%=rs("e_webname")%>"></td>
        </tr>
        <tr>
          <td align="right">英文关键字：</td>
          <td><textarea name="e_webkeyword" cols="40" rows="6" id="e_webkeyword" style="height:60; width:100%"><%=rs("e_webkeyword")%></textarea></td>
        </tr>
        <tr>
          <td width="107" align="right">英文网站说明：</td>
          <td><textarea name="e_webdes" cols="40" rows="6" id="webdes" style="height:60; width:100%"><%=rs("e_webdes")%></textarea></td>
        </tr>
		
		
		
		
		
        <tr style="display:none" >
          <td align="right">视频地址：</td>
          <td><table width="613" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="312"><input name="image" type="text" id="image" size="40" value="<%=rs("image")%>"></td>
              <td width="301"><iframe src="jiedai_up.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
            </tr>
          </table>
            <br>
            1mb以内可以上传，其他使用其他路径上传</td>
        </tr>
        <tr style="display:none">
          <td align="right">版权文字：</td>
          <td>		
			 <textarea name="content" cols="" rows="" style="display:none"><%=rs("bq")%></textarea>
			 
			 
	   <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="220"></iframe>	</td>
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
		 response.write"<img src=images/cms-ico7.gif width=12 height=11><font color=#ff0000><b>"&rs("classid")&"-</b>信息已修改成功</font>"
		 %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
        </table><%end if%>
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
</body>
</html>                                                                             