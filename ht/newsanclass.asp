<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/body.css" rel="stylesheet" type="text/css">
</head>
<body>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th class="tableHeaderText" colspan=8 height=25>区域管理</th>	
  </tr>
  <tr>
    <td width="14%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">中文名称</div></td>
    <td width="12%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">英文名称</td>
    <td width="23%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">标示图</td>
    <td width="31%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">BANNER图</td>
    <td width="7%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">排序</div></td>
    <td width="13%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">确定操作</div></td>
  </tr>
    <%set rs=server.CreateObject("adodb.recordset")
		  rs.Open "select * from jiedai_newsclass order by flag asc ",conn,1,1
		  dim paixu
		  if rs.EOF and rs.BOF then
		  response.Write "<div align=center><font color=red>还没有分类</font></center>"
		  paixu=0
		  else
		  do while not rs.EOF
	%>
     <form name="form1p" method="post" action="savenewsanclass.asp?action=edit&id=<%=int(rs("id"))%>">
    <tr>
      <td class="forumRowHighlight"><div align="center"><input name="anclass" type="text" id="anclass" size="12" value="<%=trim(rs("classname"))%>"></div></td>
	  <td align="center" class="forumRowHighlight"><input name="e_anclass" type="text" id="e_anclass" size="12" value="<%=trim(rs("e_classname"))%>"></td>
	  <td align="center" class="forumRowHighlight">
      
      
      <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><input name="image" type="text" id="tu<% =rs("id") %>" value="<% =rs("tupian") %>" size="40" style="height:20; width:60px;"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp?ffs=form1p&id=tu<% =rs("id") %>" width="200" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table>
      
      
      </td>
	  <td align="center" class="forumRowHighlight">
      
      
         <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>中：
                  <input name="pimg" type="text" id="ig<% =rs("id") %>" value="<% =rs("images") %>" size="40" style="height:20; width:60px;"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp?ffs=form1p&id=ig<% =rs("id") %>" width="200" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table>
            
            
           
         <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>英：
                  <input name="e_pimg" type="text" id="iggs<% =rs("id") %>" value="<% =rs("e_images") %>" size="40" style="height:20; width:60px;"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp?ffs=form1p&id=iggs<% =rs("id") %>" width="200" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table> 
            
            
      
      
      
      </td>
	  <td class="forumRowHighlight"><div align="center"><input name="anclassidorder" type="text" id="anclassidorder" size="4" value="<%=int(rs("flag"))%>"></div></td>
     <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit" value="修  改">&nbsp; 
     
     <% if rs("id")>37 then %>
     <a href="savenewsanclass.asp?id=<%=int(rs("id"))%>&action=del" onClick="return confirm('此操作会删除此大类下包含的小分类和商品！您确定进行删除操作吗？')"><font color=red>删除</font></a> 
     <% end if %>
     
     
     </div></td>
   </tr>
   </form>
   <%rs.MoveNext
          loop
          paixu=rs.RecordCount
          end if%>
</table>
<br>
<table cellpadding="3" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr height=25> 
	<th class="tableHeaderText" colspan=8 height=25>添加区域</th>	
  </tr>
  <tr> 
    <td class="forumRowHighlight" colspan=6 ><div align="center">注意：各项名称不能含有非法字符</div></td>
  </tr>
  <tr>
    <td width="9%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">中文名称</div></td>
    <td width="11%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">英文名称</td>
    <td width="29%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">标示图</td>
    <td width="31%" align="center" bgcolor="#CFDEEB" class="forumRowHighlight">BANNER图</td>
    <td width="9%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">排序</div></td>
   <td width="11%" bgcolor="#CFDEEB" class="forumRowHighlight"><div align="center">确定操作</div></td>
 </tr>
 <form name="form1s" method="post" action="savenewsanclass.asp?action=add">
  <tr>
    <td class="forumRowHighlight"><div align="center"><input name="anclass2" type="text" id="anclass2" size="12"></div></td>
    <td align="center" class="forumRowHighlight"><input name="e_anclass2" type="text" id="e_anclass2" size="12"></td>
    <td align="center" class="forumRowHighlight">
    
    <table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><input name="image" type="text" id="tupp" size="40" style="height:20; width:60px;"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp?ffs=form1s&id=tupp" width="200" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table>
    
    
    </td>
    <td align="center" class="forumRowHighlight">
    
    
    
    
    
    <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>中：
                  <input name="pimg" type="text" id="ighs"  size="40" style="height:20; width:60px;"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp?ffs=form1p&id=ighs" width="200" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table>
            
            
           
         <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>英：
                  <input name="e_pimg" type="text" id="iggskk"  size="40" style="height:20; width:60px;"></td>
                <td style="padding-left:8px"><iframe src="jiedai_up.asp?ffs=form1p&id=iggskk" width="200" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
              </tr>
            </table>
    
    
    
    
    
    
    
    
    
    
    </td>
    <td class="forumRowHighlight"><div align="center"><input name="anclassidorder2" type="text" id="anclassidorder2" size="4" value="<%=paixu+1%>"></div></td>
    <td class="forumRowHighlight"><div align="center"><input class="button" type="submit" name="Submit3" value="添 加"></div></td>
  </tr>
  </form>
</table>
</body>
</html>
