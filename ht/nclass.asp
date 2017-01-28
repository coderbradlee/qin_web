<!--#include file="conn.asp"-->
<%dim anclassid,anclass,paixu
anclass=request.QueryString("nclass")
anclassid=request.QueryString("id")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/body.css" rel="stylesheet" type="text/css">
</head>
<body>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th height=25 colspan=7 align="left" class="tableHeaderText" style="padding-left:72px">商品小类管理</th>	
  </tr>
  <tr> 
    <td colspan=4  class="forumRowHighlight"><div align="center">
	<select name="select" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
	<option >选择商品大类</option>
	<%set rs=server.createobject("adodb.recordset")
		rs.Open "select * from sh_sort order by anclassidorder",conn,1,1
		do while not rs.eof %>
	<option value="nclass.asp?id=<%=rs("anclassid")%>&anclass=<%=rs("anclass")%>" <%if rs("anclassid")=cint(request.QueryString("id")) then %>selected<%end if%>><%=trim(rs("anclass"))%></option>
    <%rs.movenext
		loop
		rs.close
		set rs=nothing
		%>
     </select>
    <%if request.QueryString("id")<>"" then
        response.Write "当前分类："&request.QueryString("anclass")
        end if%>
		 <%
        if anclassid="" then
        response.Write "<font color=red></font>"
        else
        set rs=server.CreateObject("adodb.recordset")
        rs.Open "select * from sh_sort2 where anclassid="&anclassid&" order by nclassidorder",conn,1,1
         if rs.EOF and rs.BOF then
		  response.Write "<font color=red>还没有大分类</font>"
		  paixu=0
		  else
         do while not rs.EOF
    %></div>    </td>
  </tr>
  <tr>
    <td width="26%" align="center" class="forumRowHighlight" >所属大类</td> 
   <td width="20%" class="forumRowHighlight" ><div align="center">分类名称</div></td>
   <td width="26%" class="forumRowHighlight"><div align="center">分类排序</div></td>
   <td width="28%" class="forumRowHighlight"><div align="center">确定操作</div></td>
 </tr>
   <form name="form1" method="post" action="savenclass.asp?action=edit&id=<%=rs("nclassid")%>&anclass=<%=request.QueryString("anclass")%>">
   <tr>
     <td width="26%" align="center" class="forumRowHighlight" ><select name="matype" id="matype">
       <option >选择商品分类</option>
       <%set rs2=server.createobject("adodb.recordset")
		rs2.Open "select * from sh_sort order by anclassidorder",conn,1,1
		do while not rs2.eof %>
       <option value="<%=rs2("anclassid")%>" <%if rs2("anclassid")=cint(request.QueryString("id")) then %>selected<%end if%>><%=trim(rs2("anclass"))%></option>
       <%rs2.movenext
		loop
		rs2.close
		set rs2=nothing
		%>
     </select></td> 
     <td width="20%" class="forumRowHighlight" ><div align="center">
	 <input name="nclass" type="text" id="nclass" size="16" value="<%=trim(rs("nclass"))%>">
	 <input name="anclassid" type="hidden" value="<%=request.QueryString("id")%>" id="Hidden1"></div>     </td>
     <td width="26%"  class="forumRowHighlight"><div align="center"><input name="nclassidorder" type="text" id="nclassidorder" size="4" value="<%=int(rs("nclassidorder"))%>"></div></td>
	<td width="28%"  class="forumRowHighlight"><div align="center"><input class=button type="submit" name="Submit" value="修 改">&nbsp;<a href="savenclass.asp?id=<%=int(rs("nclassid"))%>&action=del&anclassid=<%=request.QueryString("id")%>&anclass=<%=request.QueryString("anclass")%>" onClick="return confirm('此删除操作，会删除此分类下的所有商品！您确定进行删除操作吗？')"><font color=red>删除</font></a> </div></td>
  </tr>
 </form>
 <%rs.movenext
        loop
        paixu=rs.RecordCount
        rs.close
        set rs=nothing
        end if
        end if
	%>
</table>
<br>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th height=25 colspan=7 align="left" class="tableHeaderText" style="padding-left:72px">添加商品小分类</th>	
  </tr>
    
  <tr>
    <td width="26%" align="center" class="forumRowHighlight">所属大类</td> 
   <td width="20%" class="forumRowHighlight"><div align="center">分类名称</div></td>
   <td width="26%" class="forumRowHighlight"><div align="center">分类排序</div></td>
   <td width="28%" class="forumRowHighlight"><div align="center">确定操作</div></td>
 </tr>
 <form name="form2" method="post" action="savenclass.asp?action=add&anclass=<%=request.QueryString("anclass")%>">
 <tr>
   <td width="26%" align="center" class="forumRowHighlight"><select name="select2" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
     <option >选择商品分类</option>
     <%set rs2=server.createobject("adodb.recordset")
		rs2.Open "select * from sh_sort order by anclassidorder",conn,1,1
		do while not rs2.eof %>
     <option value="nclass.asp?id=<%=rs2("anclassid")%>&anclass=<%=rs2("anclass")%>" <%if rs2("anclassid")=cint(request.QueryString("id")) then %>selected<%end if%>><%=trim(rs2("anclass"))%></option>
     <%rs2.movenext
		loop
		rs2.close
		set rs2=nothing
		%>
   </select></td> 
   <td width="20%" class="forumRowHighlight"><div align="center"><input name="nclass2" type="text" id="nclass22" size="16"><input name="anclassid" type="hidden" value="<%=request.QueryString("id")%>"></div></td>
   <td width="26%" class="forumRowHighlight"><div align="center"><input name="nclassidorder2" type="text" id="nclassidorder22" size="4" value="<%=paixu+1%>"></div></td>
   <td width="28%" class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit2" value="添 加"></div></td>
 </tr>
</form>
</table>
</body>
</html>
