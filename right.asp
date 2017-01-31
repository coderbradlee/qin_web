<div class="right_list">
        	<h1>新闻资讯</h1>
            <ul>
         
           <%
					set rs=server.createobject("adodb.recordset")
					sql="select top 8 * from zhibenhui_news  order by id desc"
					rs.open sql,conn,1,1	
					do while not rs.eof

					%> 
		<li><a href="news_show.asp?id=<% =rs("id") %>" ><% =rs("title")%></a></li>
				
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
         
         
         
            </ul>
        </div>