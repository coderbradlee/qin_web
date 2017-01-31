<!--#include file="conn.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<TITLE><%=title%></TITLE>
<meta name="keywords" content="<%=keywords_content%>" />
<meta name="description" content="<%=description_content%>" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<script type="text/javascript" src="../js/jquery.min.js"></script>
<script type="text/javascript" src="../js/main.js" ></script>
<link href="css/index.css" rel="stylesheet" type="text/css" />
</head>

<body>

<!--#include file="top.asp" -->

<div class="banner">
<object id="bcastr4" data="/bcastr4.swf?xml=img.xml" type="application/x-shockwave-flash" width="961" height="373"><param name="movie" value="bcastr4.swf?xml=img.xml" /><param name="wmode" value="transparent" /></object>


</div>
<div class="ititle">
	
    <div class="inews fl">
    	<ul class="inewsti fl">
        	<li><a href="news_list.asp?a=37" >news</a></li>
            <li><a href="news_list.asp?a=25" >reprot</a></li>
            <li><a href="news_list.asp?a=38" >merge dynamic</a></li>
        </ul>
        <ul class="imore fr"><a href="news.asp" class="a100" >More</a></ul>
    </div>
    
    <div class="iabout fl">
    	<ul class="iaboutti">
        	<h1><a href="about.asp" class="a100" >About Us</a></h1>
            <ul class="imore fr"><a href="about.asp" class="a100" >More</a></ul>
        </ul>
    </div>
    
    <div class="ishehui fl">
    	<h1 class="ishehuiti"><a href="honor.asp" class="a100" >Corporate Social Responsibility</a></h1>
    </div>
    
    

</div>
<div class="imain">
	
<div class="inewsmain fl">
    	<ul class="anewsm">
        	  <%
					set rs=server.createobject("adodb.recordset")
					sql="select top 6 * from zhibenhui_news  where e_title<>'' and classid=37 order by id desc"
					rs.open sql,conn,1,1	
					do while not rs.eof

					%> 
		<li><a href="news_show.asp?id=<% =rs("id") %>" ><% =got(rs("e_title"),30) %></a><span>[<% =FormatDate(rs("addtime"),4) %>]</span></li>
				
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
   
        </ul>
        
        	<ul class="bnewsm">
        <%
					set rs=server.createobject("adodb.recordset")
					sql="select top 6 * from zhibenhui_news  where  e_title<>'' and classid=25 order by id desc"
					rs.open sql,conn,1,1	
					do while not rs.eof

					%> 
		<li><a href="news_show.asp?id=<% =rs("id") %>" ><% =got(rs("e_title"),30) %></a><span>[<% =FormatDate(rs("addtime"),4) %>]</span></li>
				
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
        </ul>
         <ul class="dnewsm">
        <%
					set rs=server.createobject("adodb.recordset")
					sql="select top 6 * from zhibenhui_news  where classid=38 order by id desc"
					rs.open sql,conn,1,1	
					do while not rs.eof

					%> 
		<li><a href="news_show.asp?id=<% =rs("id") %>" ><% =got(rs("title"),30) %></a><span>[<% =FormatDate(rs("addtime"),4) %>]</span></li>
				
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
        </ul>
    </div>
    
    
  <div class="iabout fl mline"><ul class="iaboutmain">   
    <%
	set ris=server.createobject("adodb.recordset")
	sqli="select * from jiedai_qita where id=2"
	 ris.open sqli,conn,1,1
	if not ris.eof then								  
	i_body=ris("e_body")
	end if	  
	 ris.close
	 set ris=nothing
	if i_body="" then
	response.Write"NO INFO"
	else
	response.Write i_body
	end if		
			%>
    </ul></div>
    
    
    <div class="ishehui fl"><center>
      <a href="honor.asp" target="_blank"><img src="images/dot002.png" width="213" height="175" border="0" /></a>
    </center></div>
    
    
    

</div>



<!--#include file="foot.asp" -->


</body>
</html>
