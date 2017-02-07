<!--#include file="conn.asp" -->
<!--#include file="Function_Page.asp" -->
<%
	 mf="academy"	
	 a_title="智本汇商学院"					  
									  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<TITLE><%=a_title%>_<%=title%></TITLE>
<meta name="keywords" content="<%=keywords_content%>" />
<meta name="description" content="<%=description_content%>" />
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/main.js" ></script>
<link href="css/main.css" rel="stylesheet" type="text/css" />
</head>

<body>


<!--#include file="top.asp" -->


<Div class="nmain">

<!-- .nleft -->
	<div class="nleft">
    	<ul class="left_list">
        	
            
                    	  <%
					
	                if aid<>"" then aid=int(aid)
					set res=server.createobject("adodb.recordset")
					sql="select * from academy_newsclass order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 

 <li><a href="academy_list.asp?a=<%=res("id")%>"  ><%=res("classname")%></a></li>
    
    
        
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
         
            
            
        </ul>

    </div>
 <!-- END .nleft-->   
 
 <!-- .ncenter -->
 <div class="ncenter">
	<ul class="nc_title">首页 > 智本汇商学院  <span> > <% =a_title %></span></ul>
   	<%call banner(200)%>
    <ul class="nbody">
  
  
  
    	  <%
					set res=server.createobject("adodb.recordset")
					sql="select * from academy_newsclass order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 
  
  
	  <div class="news_home">
 			 <ul class="news_hti"><h1><% =res("classname") %></h1> <a href="academy_list.asp?a=<% =res("id") %>" >more</a></ul>
          <ul class="news_main">
           	  <div class="fl"><img src="uploadfile/<% =res("tupian") %>"  /></div>
              <div class="fr"><ul class="honor"> 
     
     
     
  <%
					set rs=server.createobject("adodb.recordset")
					sql="select top 6 * from academy_news  where classid="&res("id")&" order by id desc"
					rs.open sql,conn,1,1	
					do while not rs.eof

					%> 
		<li><a href="academy_show.asp?id=<% =rs("id") %>" ><% =got(rs("title"),36) %></a><span>[<% =FormatDate(rs("addtime"),4) %>]</span></li>
				
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
	
     
     </ul></div>
          </ul>
          <div class="clearfix"></div>
  		</div>
  
  
  <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
  
  




    
    </ul>
 </div>
 <!-- END .ncenter -->
    
    <!-- .nright -->
    <div class="nright">
    	
        
        <!--#include file="cright.asp" -->

        
    </div>
    <!-- END .nright -->
    
    
 <div class="clearfix"></div>
</Div>




<!--#include file="foot.asp" -->


</body>
</html>
