<%@language="vbscript" codepage="936"%>
<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<%
sjid=request("sjid")

%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title></title>
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="input">
  <tr>
    <td><%
if trim(request.querystring("action"))="add" then
if trim(request.form("submit"))="添加" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi"
	rs.open sql,conn,3,2
	rs.addnew
	rs("jtype")=request.form("jtype")
	rs("diqu")=request.form("diqu")
	rs("jgname")=request.form("jgname")
	rs("jgdz")=request.form("jgdz")
	rs("jgyb")=request.form("jgyb")
	rs("jgtel")=request.form("jgtel")
	rs("yysj")=request.form("yysj")
	rs("zjdcz")=request.form("zjdcz")
	rs("zbxx")=request.form("zbxx")
	rs("pjys")=request.form("pjys")
	rs("jcolor")=request.form("jcolor")	
	rs("jcontent")=request.form("content")	
	rs("sjid")=request.form("sjid")	
	
	rs("sjpic")=request.form("image")	
	rs("sjditu")=request.form("ditu")	
	
	
	

	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('添加成功!');location='?action=list&sjid="&sjid&"'</script>"
end if

%>
      <script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_add(){
	var notnull;
	notnull=true;
	if (document.form1.title.value==""){
		alert("信息名称不能为空！");
		document.form1.title.focus();
		notnull=false;
		}
	return notnull;
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">添加--<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>--信息</font></strong></td>
        </tr>
      </table>
      <br>
      
      
      
      <script language="JavaScript">
<!--
function show(){
if (form1.jtype.value=="医药企业名录"){
 document.getElementById("z1").style.display='none';
 document.getElementById("z2").style.display='none';
 document.getElementById("z3").style.display='none';
 document.getElementById("z4").style.display='none';
}
else{
 document.getElementById("z1").style.display='';
 document.getElementById("z2").style.display='';
 document.getElementById("z3").style.display='';
 document.getElementById("z4").style.display='';
}

}

//-->
</script>
      
      
      
      <form name="form1" method="post" action="?action=add" onSubmit="return check_add()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td width="75" align="center">商家地区：</td>
            <td>
            
            
            <% 	  
sql="select * from sh_sort order by anclassidorder asc"
set rs=conn.execute(sql)  
 %>
                  <select name="diqu" id="classid" style="width:100%">
                    <% do while not rs.eof %>
                    <option value="<%= rs("anclass") %>"><%= rs("anclass") %></option>
                    <%
		 rs.movenext
		loop
		 %>
                  </select>            </td>
          </tr>
          <tr>
            <td align="center">商家类型：</td>
            <td>
            
            
                        
						
                        
                        
                        
                      <input name="jtype" type="text" id="textfield2" size="50">
              
              <% 	  
sqle="select * from sjsort where sjid="&sjid&" order by anclassidorder asc"
set res=conn.execute(sqle)  
 %>
              <select name="diqu" id="diqu" onChange="(document.form1.jtype.value+=this.options[this.selectedIndex].value+',')">
              
                  <option>请选择经营类型</option>
              
                  <% do while not res.eof %>
                  <option value="<%= res("anclass") %>"><%= res("anclass") %></option>
                  <%
		 res.movenext
		loop
		res.close
		set res=nothing

		 %>
              </select>            </td>
          </tr>
          <tr>
            <td align="center">商家名称：</td>
            <td><input type="text" name="jgname" id="textfield" style="width:100%"></td>
          </tr>
          <tr>
            <td align="center">商家地址：</td>
            <td><input type="text" name="jgdz" id="textfield" style="width:60%"> 邮编:<input type="text" name="jgyb" id="textfield" style="width:30%"></td>
          </tr>
          <tr>
            <td align="center">联系电话：</td>
            <td><input type="text" name="jgtel" id="textfield" style="width:100%"></td>
          </tr>
          <tr>
            <td align="center">营业时间：</td>
            <td><input name="yysj" type="text" id="yysj" style="height:20; width:100%" value="10:00 ～ 22:00" size="40"></td>
          </tr>
          <tr>
            <td align="center"><p>周边信息：</p>            </td>
            <td><input name="zbxx" type="text" id="zbxx" size="40" style="height:20; width:100%"></td>
          </tr>
          <tr id="z5">
            <td align="right">最近的车站：</td>
            <td><input name="zjdcz" type="text" id="zjdcz" size="40" style="height:20; width:100%"></td>
          </tr>
          <tr id="z3">
            <td align="center">平均预算：</td>
            <td><input name="pjys" type="text" id="pjys" style="height:20; width:100%" size="40"></td>
          </tr>
          <tr>
            <td align="center">上传缩略图：</td>
            <td><table width="553" border="0" cellspacing="1" cellpadding="0">
              <tr>
                <td width="323"><input name="image" type="text" id="image" size="40"></td>
                <td width="270"><iframe src="tongyongsc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td align="center">上传地图：</td>
            <td><table width="560" border="0" cellspacing="1" cellpadding="0">
              <tr>
                <td width="299"><input name="ditu" type="text" id="textfield4" size="40"></td>
                <td width="300"><iframe src="tongyongscc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table></td>
          </tr>
          <tr style="display:none">
            <td align="center">标题颜色：</td>
            <td>
<select name=jcolor size=1 id="jcolor">
<option value="">标题颜色</option>
<option style="background-color:Black;color:Black" value=Black>黑 色</option>
<option style="background-color:Red;color:Red" value=Red>红 色</option>
<option style="background-color:Yellow;color:Yellow" value=Yellow>黄 色</option>
<option style="background-color:Green;color:Green" value=Green>绿 色</option>
<option style="background-color:Orange;color:Orange" value=Orange>橙 色</option>
<option style="background-color:Purple;color:Purple" value=Purple>紫 色</option>
<option style="background-color:Blue;color:Blue" value=Blue>蓝 色</option>
<option style="background-color:Brown;color:Brown" value=Brown>褐 色</option>
<option style="background-color:Teal;color:Teal" value=Teal>墨 绿</option>
<option style="background-color:Navy;color:Navy" value=Navy>深 蓝</option>
<option style="background-color:Maroon;color:Maroon" value=Maroon>赭 石</option>
<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">粉 绿</option>
<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">淡 绿</option>
<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">黄 灰</option>
<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">翠 绿</option>
<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">综 红</option>
<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">砖 红</option>
<option style="background-color:#6495ED;color: #6495ED" value="#6495ED">淡 蓝</option>
<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">暗 红</option>
<option style="background-color:#FF1493;color: #FF1493" value="#FF1493">玫瑰红</option>
<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">紫 红</option>
<option style="background-color:#FFD700;color: #FFD700" value="#FFD700">桔 黄</option>
<option style="background-color:#DAA520;color: #DAA520" value="#DAA520">军 黄</option>
<option style="background-color:#808080;color: #808080" value="#808080">烟 灰</option>
<option style="background-color:#778899;color: #778899" value="#778899">深 灰</option>
<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">灰 蓝</option>
</select></td>
          </tr>
          <tr>
            <td align="center">说明：</td>
            <td>
			
			
			
			
			
			 <textarea name="content" style="display:none"></textarea>
			 
			 
	   <iframe id="eWebEditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="250"></iframe>			</td>
          </tr>
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="添加" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit2" value="重置" style="width:80; height:30; cursor:hand">
              <input name="sjid" type="hidden" id="sjid" value="<%=request.querystring("sjid")%>"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <%
if trim(request.querystring("action"))="edit" then
if trim(request.form("submit"))="修改" then
	id=trim(request.querystring("id"))
	dim rs,sql
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,3
	rs("jtype")=request.form("jtype")
	rs("diqu")=request.form("diqu")
	rs("jgname")=request.form("jgname")
	rs("jgdz")=request.form("jgdz")
	rs("sjid")=request.form("sjid")
	rs("jgyb")=request.form("jgyb")
	rs("jgtel")=request.form("jgtel")
	rs("yysj")=request.form("yysj")
	rs("zjdcz")=request.form("zjdcz")
	rs("zbxx")=request.form("zbxx")
	rs("pjys")=request.form("pjys")
	rs("jcolor")=request.form("jcolor")	
	rs("jcontent")=request.form("content")	
	
	rs("sjpic")=request.form("image")	
	rs("sjditu")=request.form("ditu")	
	
	
	
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('修改成功!');location='?action=list&sjid="&sjid&"'</script>"
	response.end
end if
id=trim(request.querystring("id"))
set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,1
%>







      
      <script language="JavaScript">
<!--
function show(){
if (form1.jtype.value=="医药企业名录"){
 document.getElementById("z1").style.display='none';
 document.getElementById("z2").style.display='none';
 document.getElementById("z3").style.display='none';
 document.getElementById("z4").style.display='none';
}
else{
 document.getElementById("z1").style.display='';
 document.getElementById("z2").style.display='';
 document.getElementById("z3").style.display='';
 document.getElementById("z4").style.display='';
}

}

//-->
</script>
      










      <script language="javascript" type="text/javascript">
// 验证用户名和留言
function check_edit(){
	var notnull;
	notnull=true;
	if (document.form1.title.value==""){
		alert("标题不能为空！");
		document.form1.title.focus();
		notnull=false;
		}
	return notnull;
	}
      </script>
      <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
        <tr>
          <td bgcolor="#DFEFFF" style="padding-left:20px"><strong><font color="#215dc6">修改--<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>--信息</font></strong></td>
        </tr>
      </table>
      <br>
      <form name="form1" method="post" action="?action=edit&id=<%= trim(request.querystring("id")) %>" onSubmit="return check_edit()">
        <table width="100%" height="125" border="0" cellpadding="3" cellspacing="0" bordercolor="#cccccc">
          <tr>
            <td width="78" align="right">商家地区：</td>
            <td><% 	  
sqle="select * from sh_sort order by anclassidorder asc"
set res=conn.execute(sqle)  
 %>
              <select name="diqu" id="diqu" style="width:100%">
                <% do while not res.eof %>
                <option value="<%= res("anclass") %>"<%if res("anclass")=rs("diqu") then response.Write("selected")%>><%= res("anclass") %></option>
                <%
		 res.movenext
		loop
		res.close
		set res=nothing
		 %>
                </select>            </td>
          </tr>
          <tr>
            <td width="78" align="right">经营类型：</td>
            <td>
			
            
            
              <input name="jtype" type="text" id="textfield2" size="50" value="<%=rs("jtype")%>">
              
              <% 	  
sqle="select * from sjsort where sjid="&sjid&" order by anclassidorder asc"
set res=conn.execute(sqle)  
 %>
              <select name="11diqu" id="1diqu" onChange="(document.form1.jtype.value+=this.options[this.selectedIndex].value+',')">
              
                  <option>请选择经营类型</option>
              
                  <% do while not res.eof %>
                  <option value="<%= res("anclass") %>"><%= res("anclass") %></option>
                  <%
		 res.movenext
		loop
		res.close
		set res=nothing

		 %>
              </select>            </td>
          </tr>
          <tr>
            <td width="78" align="right">商家名称：</td>
            <td><input type="text" name="jgname" id="textfield" style="width:100%" value="<%=rs("jgname")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right">商家地址：</td>
            <td><input type="text" name="jgdz" id="textfield" style="width:60%" value="<%=rs("jgdz")%>"> 邮编:<input type="text" name="jgyb" id="textfield" style="width:30%" value="<%=rs("jgyb")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right">联系电话：</td>
            <td><input type="text" name="jgtel" id="textfield" style="width:100%" value="<%=rs("jgtel")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right">营业时间：</td>
            <td><input name="yysj" type="text" id="yysj" size="40" style="height:20; width:100%" value="<%=rs("yysj")%>"></td>
          </tr>
          <tr>
            <td width="78" align="right"><p>周边信息：</p></td>
            <td><input name="zbxx" type="text" id="yywz3" size="40" style="height:20; width:100%" value="<%=rs("zbxx")%>"></td>
          </tr>
          
          
         <tr id="z3">
            <td width="78" align="right">最近的车站：</td>
            <td><input name="zjdcz" type="text" id="zjdcz" size="40" style="height:20; width:100%"  value="<%=rs("zjdcz")%>"></td>
          </tr>
          
            <td width="78" align="right">平均预算：</td>
            <td><input name="pjys" type="text" id="pjys" size="40" style="height:20; width:100%" value="<%=rs("pjys")%>"></td>
          </tr>  <tr>
              <td align="center">上传缩略图：</td>
              <td><table width="553" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="323"><input name="image" type="text" id="image2" size="40" value="<%=rs("sjpic")%>"></td>
                    <td width="270"><iframe src="tongyongsc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                  </tr>
              </table></td>
            </tr>
            <tr>
              <td align="center">上传地图：</td>
              <td><table width="560" border="0" cellspacing="1" cellpadding="0">
                  <tr>
                    <td width="299"><input name="ditu" type="text" id="textfield3" size="40" value="<%=rs("sjditu")%>"></td>
                    <td width="300"><iframe src="tongyongscc.asp" width="270" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
                  </tr>
              </table></td>
            </tr>
          
          <tr>
            <td width="78" align="right">标题颜色：</td>
            <td>
			
			
<select name=jcolor size=1 id="jcolor">
<option value="">选择颜色</option>
<option style="background-color:Black;color:Black" value=Black <%if rs("jcolor")="Black" then response.write "selected" %>>黑 色</option>
<option style="background-color:Red;color:Red" value=Red <%if rs("jcolor")="Red" then response.write "selected" %>>红 色</option>
<option style="background-color:Yellow;color:Yellow" value=Yellow <%if rs("jcolor")="Yellow" then response.write "selected" %>>黄 色</option>
<option style="background-color:Green;color:Green" value=Green <%if rs("jcolor")="Green" then response.write "selected" %>>绿 色</option>
<option style="background-color:Orange;color:Orange" value=Orange <%if rs("jcolor")="Orange" then response.write "selected" %>>橙 色</option>
<option style="background-color:Purple;color:Purple" value=Purple <%if rs("jcolor")="Purple" then response.write "selected" %>>紫 色</option>
<option style="background-color:Blue;color:Blue" value=Blue <%if rs("jcolor")="Blue" then response.write "selected" %>>蓝 色</option>
<option style="background-color:Brown;color:Brown" value=Brown <%if rs("jcolor")="Brown" then response.write "selected" %>>褐 色</option>
<option style="background-color:Teal;color:Teal" value=Teal <%if rs("jcolor")="Teal" then response.write "selected" %>>墨 绿</option>
<option style="background-color:Navy;color:Navy" value=Navy <%if rs("jcolor")="Navy" then response.write "selected" %>>深 蓝</option>
<option style="background-color:Maroon;color:Maroon" value=Maroon <%if rs("jcolor")="Maroon" then response.write "selected" %>>赭 石</option>
<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF" <%if rs("jcolor")="#00FFFF" then response.write "selected" %>>粉 绿</option>
<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4" <%if rs("jcolor")="#7FFFD4" then response.write "selected" %>>淡 绿</option>
<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4" <%if rs("jcolor")="#FFE4C4" then response.write "selected" %>>黄 灰</option>
<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00" <%if rs("jcolor")="#7FFF00" then response.write "selected" %>>翠 绿</option>
<option style="background-color:#D2691E;color: #D2691E" value="#D2691E" <%if rs("jcolor")="#D2691E" then response.write "selected" %>>综 红</option>
<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50" <%if rs("jcolor")="#FF7F50" then response.write "selected" %>>砖 红</option>
<option style="background-color:#6495ED;color: #6495ED" value="#6495ED" <%if rs("jcolor")="#6495ED" then response.write "selected" %>>淡 蓝</option>
<option style="background-color:#DC143C;color: #DC143C" value="#DC143C" <%if rs("jcolor")="#DC143C" then response.write "selected" %>>暗 红</option>
<option style="background-color:#FF1493;color: #FF1493" value="#FF1493" <%if rs("jcolor")="#FF1493" then response.write "selected" %>>玫瑰红</option>
<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF" <%if rs("jcolor")="#FF00FF" then response.write "selected" %>>紫 红</option>
<option style="background-color:#FFD700;color: #FFD700" value="#FFD700" <%if rs("jcolor")="#FFD700" then response.write "selected" %>>桔 黄</option>
<option style="background-color:#DAA520;color: #DAA520" value="#DAA520" <%if rs("jcolor")="#DAA520" then response.write "selected" %>>军 黄</option>
<option style="background-color:#808080;color: #808080" value="#808080" <%if rs("jcolor")="#808080" then response.write "selected" %>>烟 灰</option>
<option style="background-color:#778899;color: #778899" value="#778899" <%if rs("jcolor")="#778899" then response.write "selected" %>>深 灰</option>
<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE" <%if rs("jcolor")="#B0C4DE" then response.write "selected" %>>灰 蓝</option>
</select>			</td>
          </tr>
          <tr>
            <td width="78" align="right">说明：</td>
            <td><textarea name="content" style="display:none"><%=rs("jcontent")%></textarea>
                <iframe id="ewebeditor1" src="<%=webed%>" frameborder="0" scrolling="no" width="100%" height="250"></iframe></td>
          </tr>
          <tr>
            <td colspan="2" style="padding-left:100px"><input type="submit" name="submit" value="修改" style="width:80; height:30; cursor:hand">
              &nbsp;
              <input type="reset" name="submit22" value="重置" style="width:80; height:30; cursor:hand">
              <input name="sjid" type="hidden" id="sjid" value="<%=sjid%>"></td>
          </tr>
        </table>
      </form>
      <% end if %>
      <% 
if trim(request.querystring("action"))="del" then
	id=trim(request.querystring("id"))
	sjid=request.QueryString("sjid")
	id=replacebadchar(id)
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,3
	rs.delete
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('修改成功!');location='?action=list&sjid="&sjid&"'</script>"
end if
 %>
 
 
      <% 
if trim(request.querystring("zhiding"))="zdyes" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,1,3
	rs("tuijian")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('已取消置顶!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>


      <% 
if trim(request.querystring("zhiding"))="zdno" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_paijiexi where id="&id
	rs.open sql,conn,3,2
	rs("tuijian")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('置顶成功!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>
 
 
 
      <% 
if trim(request.querystring("toutiao"))="ttyes" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
	page=request.QueryString("page")
	cid=request.QueryString("cid")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_News where id="&id
	rs.open sql,conn,3,2
	rs("toutiao")=0
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('头条信息已取消!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>
 
 
      <% 
if trim(request.querystring("toutiao"))="ttno" then
	id=trim(request.querystring("jid"))
	sjid=trim(request.querystring("sjid"))
	page=request.QueryString("page")
	keywords=request.QueryString("keywords")
	user_ename=request.QueryString("user_ename")
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_News where id="&id
	rs.open sql,conn,3,2
	rs("toutiao")=1
	rs.update
	rs.requery
	rs.close
	set rs=nothing
	response.write "<script>alert('头条设置成功!');location='?Action=list&page="&page&"&keywords="&keywords&"&user_ename="&user_ename&"&cid="&cid&"&sjid="&sjid&"'</script>"
	response.end
end if
 %>
 
 
 
 
 
 
 
 
 
 
<%
if trim(request.querystring("action"))="list" then
 %>
 
 
 
 
 
  
 
 
 
 
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:7px">
  <form name="form" method="post" action="?Action=list&sjid=<%=sjid%>">
    <tr> 
      <td align="left">　<%
	if request("sjid")=1 then
	 response.write"<font color=#ff0000>美食</font>"
	  elseif request("sjid")=2 then 
	   response.write"<font color=#ff0000>休闲</font>" 
	   elseif request("sjid")=3 then 
	   response.Write"<font color=#ff0000>客房</font>" 
	   end if%>--信息搜索：

        <input name="keywords" type="text" class="input" id="keywords" style="width:150px;height:21px; padding-left:5px" onFocus='this.select()' onBlur="if (this.value ==''){this.value=this.defaultValue}" onClick="if(this.value=='输入信息关键词')this.value=''" value="输入信息关键词">
	  <input name="Submit" type="submit" class="bt" id="Submit" value="搜索">
      </td>
      <td align="right">&nbsp;</td>
    </tr>
  </form>
</table>
 
 
 
 
 
 <table width="100%" border="0" cellspacing="0" cellpadding="6">
   <tr>
     <td width="82%" valign="top"><table width="99%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
       <tr align="center">
         <td width="26%" align="left" bgcolor="#DFEFFF" style="padding-left:10px">机构名称 
           
           	  <% 	  
sql_classid="select * from sh_sort"
set rs_classid=conn.execute(sql_classid)  
 %>	  
	  <select name="classid" id="classid" onchange='javascript:window.open(this.options[this.selectedIndex].value,"_self")'>
	<option value="?action=list&dq=&sjid=<%=sjid%>">选择地区</option>

	  <% do while not rs_classid.eof %>
        <option value="?action=list&dq=<%= rs_classid("anclass") %>&sjid=<%=sjid%>" <% if request.QueryString("dq")=rs_classid("anclass") then response.write"selected"%>><%= rs_classid("anclass") %></option>
		<%
		 rs_classid.movenext
		loop
		 %>
	<option value="?action=list&dq=&sjid=<%=sjid%>">所有地区</option>
    </select>           </td>
         <td width="24%" align="left" bgcolor="#DFEFFF">电话</td>
         <td width="11%" align="center" bgcolor="#DFEFFF">缩略图|地图</td>
         <td width="13%" align="center" bgcolor="#DFEFFF" style="display:none">分店管理</td>
         <td width="10%" align="center" bgcolor="#DFEFFF">优惠券</td>
         <td width="5%" align="center" bgcolor="#DFEFFF">置顶</td>
         <td width="5%" align="center" bgcolor="#DFEFFF">删除</td>
         <td width="6%" align="center" bgcolor="#DFEFFF">编辑</td>
       </tr>
     </table>
       <br>
       <%
	   
	   
	   
	   
	   
	   
	   
	   
	   
	   	newsclass=request.form("newsclass")
	keywords=request.form("keywords")
	cid=request.querystring("cid")
	dq=request.querystring("dq")
	set rs=server.createobject("adodb.recordset")
		
	sql="select * from jiedai_paijiexi where 1=1 "
	if keywords<>"" then
	sql=sql+" and jgname like '%"&keywords&"%' or jcontent  like '%"&keywords&"%' "
	end if
	
	if sjid<>"" then
	sql=sql+" and sjid="&sjid&" "
	end if
	
	
	
	if cid<>"" then
	sql=sql+" and jtype='"&cid&"' "
	end if

	if dq<>"" then
	sql=sql+" and diqu='"&dq&"' "
	end if
	
	
	
	sql=sql+" order by tuijian desc,id desc"
		
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

	   
	   
	   
if rs.bof then
response.write("<center><br><br><br><br><br><br>")
response.write("<font color=red>暂无</font>信息！")
response.write("<br><br><br><br><br><br></center>")
end if

 for i=1 to rs.pagesize
    if rs.eof then exit for 
 %>
       <table width="99%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc">
         <tr>
           <td width="26%" class="line" style="padding-left:10px">·&nbsp;<font color="<%=rs("jcolor")%>"><%= rs("jgname") %></font></td>
           <td width="24%" class="line"><%= rs("jgtel") %></td>
           <td width="11%" align="center" class="line"><%if rs("sjpic")<>"" then%><a href="../uploadfile/<%=rs("sjpic")%>" target="_blank"><img src="images/arrow38.gif" width="18" height="18" border="0" align="absmiddle"></a><%else%><img src="images/arrow38.gif" width="18" height="18" border="0" align="absmiddle" style="filter:progid:DXImageTransform.Microsoft.BasicImage(grayScale=1)"><%end if%>| <%if rs("sjditu")<>"" then%><a href="../uploadfile/<%=rs("sjditu")%>" target="_blank"><img src="images/arrow38.gif" width="18" height="18" align="absmiddle" border="0"></a><%else%><img src="images/arrow38.gif" width="18" height="18" align="absmiddle" border="0" style="filter:progid:DXImageTransform.Microsoft.BasicImage(grayScale=1)"><%end if%></td>
           <td width="13%" align="center" class="line" style="display:none"><a href="jiedai_fendian.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=add">添加</a> | <a href="jiedai_fendian.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=list">管理</a></td>
           <td width="10%" align="center" class="line"><a href="jiedai_huoban.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=add">添加</a> | <a href="jiedai_huoban.asp?sid=<%=rs("id")%>&sjid=<%=sjid%>&action=list">管理</a></td>
           <td width="5%" align="center" class="line"><%if rs("tuijian")=1 then%>
               <a href="?Action=list&zhiding=zdyes&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>&dq=<%=dq%>&sjid=<%=sjid%>"><img src="images/Ok.gif" alt="已置顶" width="16" height="16" border="0" /></a>
               <%else%>
               <a href="?Action=list&zhiding=zdno&jid=<%=rs("id")%>&page=<%=page%>&keywords=<%=keywords%>&user_ename=<%=user_ename%>&cid=<%=cid%>&dq=<%=dq%>&sjid=<%=sjid%>"><img src="images/err.gif" alt="未置顶" width="12" height="11" border="0" /></a>
               <%end if%></td>
           <td width="5%" align="center" class="line"><a href="?action=del&id=<%= rs("id") %>&sjid=<%=sjid%>" onClick="return window.confirm('确定删除吗?');">[删除]</a> </td>
           <td width="6%" align="center" class="line"><a href="?action=edit&id=<%= rs("id") %>&sjid=<%=sjid%>">[编辑]</a> </td>
         </tr>
       </table>
       <table width="100%" border="0" cellspacing="0" cellpadding="0">
         <tr>
           <td height="5"></td>
         </tr>
       </table>
       <% 
rs.movenext
next
else
response.write"<div align=center><br>暂无信息<br><br></div>"
end if
%></td>
   </tr>
 </table> 
 <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0" bordercolor="#cccccc" background="images/bg_title.gif">
  <tr>
    <td align="center">第<%= page %>页&nbsp;
        <% if page<>1 then %>
        <a href="?action=list&page=1&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">首页</a>
        <% else %>
      首页
      <% end if %>
      &nbsp;
      <% if page>1 then %>
      <a href="?action=list&page=<%= page-1 %>&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">上一页</a>
      <% else %>
      上一页
      <% end if %>
      &nbsp;
      <% if page<rs.pagecount then %>
      <a href="?action=list&page=<%= page+1 %>&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">下一页</a>
      <% else %>
      下一页
      <% end if %>
      &nbsp;
      <% if page<rs.recordcount then %>
      <a href="?action=list&page=<%= rs.recordcount %>&cid=<%= cid %>&dq=<%=dq%>&sjid=<%=sjid%>">末页</a>
      <% else %>
      末页
      <% end if %>
      &nbsp;总数<%= rs.recordcount %>条</td>
  </tr>
</table>
<% end if %></td>
  </tr>
</table>
</body>
</html>