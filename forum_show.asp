<!--#include file="conn.asp" -->


<% 	




set ris=server.createobject("adodb.recordset")
    sqli="select * from forum_fuwu where id="&request.QueryString("id")
	 ris.open sqli,conn,1,1
	if not ris.eof then								  
	ai_ti=ris("classid")
	ai_body=ris("body")
	end if
		  
	 ris.close
	 set ris=nothing

				  
	a_title="并购交易社区"

	 
	 mf="forum"						  
									  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<TITLE><%=a_title%>_<%=title%></TITLE>
<meta name="keywords" content="<%=keywords_content%>" />
<meta name="description" content="<%=description_content%>" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
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
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=17 or id= 18 or id=19 or id= 20 order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof
					%> 

  <li><a href="forum.asp?a=<%=res("id")%>"  ><%=res("classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
	<li><A href="forum_list.asp" class="focus" >并购交易社区</A></li>
            
            
            
        </ul>

    </div>
 <!-- END .nleft-->   
 
 <!-- .ncenter -->
 <div class="ncenter">
	<ul class="nc_title">首页 > 并购交易社区  <span> > <% =a_title %></span></ul>
   	<%call banner(204)%>
    <ul class="nbody">
    
		<p style="font-weight:bold; text-align:center"><%=ai_ti%></p>
                            <p><%=ai_body%></p>
                   
                            



    
    </ul>
 </div>
 <!-- END .ncenter -->
    
    <!-- .nright -->
    <div class="nright">
    	
        
        <!--#include file="right.asp" -->

        
    </div>
    <!-- END .nright -->
    
    
 <div class="clearfix"></div>
</Div>




<!--#include file="foot.asp" -->


</body>
</html>
