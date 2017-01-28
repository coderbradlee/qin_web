<!--#include file="conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<TITLE><%=title%></TITLE>
<meta name="keywords" content="<%=keywords_content%>" />
<meta name="description" content="<%=description_content%>" />
<link href="css.css" rel="stylesheet" type="text/css" />
</head>

<body>


<!--#include file="top.asp" -->


<table width="940" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="940" valign="top"><table width="940" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="163" height="450" align="center" valign="top" background="images/leftbian.jpg" bgcolor="#ECE9D8"><!--#include file="left1.asp" --></td>
        <td width="22" align="center" valign="top" background="images/dian.jpg" bgcolor="#FFFFFF"></td>
        <td valign="top" bgcolor="#F7FBED"><table width="100%" height="13" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr>
            <td></td>
          </tr>
        </table>
          <table width="755" height="28" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="755" align="center" style="background-image:url(images/200842641825173.jpg); background-repeat:no-repeat; background-position:center" class="dzt"><%
		N=request.QueryString("Nclass")
		set rh=server.CreateObject("adodb.recordset")
		seh="select * from jiedai_newsclass where id="&N&""
		rh.open seh,conn,1,1
		%><%=rh("classname")%><%rh.close:set rh=nothing%></td>
            </tr>
          </table>
          
          
          
          
          
          
          
          <table width="92%" height="102" border="0" align="center" cellpadding="15" cellspacing="0" class="zw">
  <tr>
    <td><table width="98%" height="321" border="0" align="center" cellpadding="0" cellspacing="0" class="zw">
      <tbody>
        <tr>
          <td valign="top"><% 
dim rs,sql
set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_News where classid="&N&" order by tuijian desc,id desc"
	rs.open sql,conn,1,1
	rs.pagesize=15
	
	if not rs.eof then
			if trim(request.querystring("ynpage")<>"") then 
			if isnumeric(trim(request.querystring("ynpage")))=false then
			page=1
			else
			page=cint(trim(request.querystring("ynpage")))
			end if
			else
			page=1
			end if
		
			if page<1 then
				page=1
			elseif page>rs.pagecount then
				page=rs.pagecount
			end if
			rs.absolutepage=page
 %>
              <%
			for news=1 to rs.pagesize
			if rs.eof then exit for
			 %>
              <table width="98%" border="0" align="center" cellpadding="3" cellspacing="0" bgcolor="#efefef" class="zw">
                <tr>
                  <td width="24" align="center" bgcolor="#FFFFFF" class="xiahuaxian"><table width="50%" border="0" cellpadding="0" cellspacing="0" class="zw">
                      <tr>
                        <td align="center" style="font-size:12px;">·</td>
                      </tr>
                  </table></td>
                  <td bgcolor="#FFFFFF" class="xiahuaxian"><a href="News_Show.asp?id=<%= rs("id") %>&amp;Nclass=<%=rs("classid")%>&amp;N=N" target="_blank" class="zw"><font color="<%=rs("titlecolor")%>"><%= rs("title") %></font></a></td>
                  <td width="150" align="center" bgcolor="#FFFFFF" class="xiahuaxian"><font color="#888888"><%= rs("addtime") %>&nbsp;</font></td>
                </tr>
              </table>
              <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="line">
                <tr>
                  <td height="1" background="images/dot.jpg"></td>
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
response.write"<div align=center><br>暂无信息<br></div>"
end if
 %>
              <br />
              <table width="100%" height="26" border="0" align="center" cellpadding="0" cellspacing="0" class="zw">
                <tr>
                  <td width="78%" align="center">第<%= page %>页&nbsp;
                      <% if page<>1 then %>
                      <a href="?ynpage=1&amp;jid=<%=jid%>" class="zw">首页</a>
                      <% else %>
                    首页
                    <% end if %>
                    &nbsp;
                    <% if page>1 then %>
                    <a href="?ynpage=<%= page-1 %>&amp;jid=<%=jid%>" class="zw">上一页</a>
                    <% else %>
                    上一页
                    <% end if %>
                    &nbsp;
                    <% if page<rs.pagecount then %>
                    <a href="?ynpage=<%= page+1 %>&amp;jid=<%=jid%>" class="zw">下一页</a>
                    <% else %>
                    下一页
                    <% end if %>
                    &nbsp;
                    <% if page<rs.pagecount then %>
                    <a href="?ynpage=<%=rs.pagecount%>&amp;jid=<%=jid%>" class="zw">末页</a>
                    <% else %>
                    末页
                    <% end if %>
                    &nbsp;总数<%= rs.recordcount %>条</td>
                  <td width="22%" align="center">转到第
                    <select name="select" onchange='javascript:window.open(this.options[this.selectedIndex].value,&quot;_top&quot;)'>
                        <%for m = 1 to rs.pagecount%>
                        <option value="?ynpage=<%=m%>&amp;jid=<%=jid%>"<%if page=m then response.write"selected"%>><%=m%></option>
                        <% next %>
                      </select>
                    页</td>
                </tr>
            </table></td>
        </tr>
      </tbody>
    </table></td>
  </tr>
</table>

          
          
          
          
          
          
          
          
          
          
          </td>
        </tr>
    </table></td>
  </tr>
</table>
<!--#include file="foot.asp" -->
</body>
</html>
