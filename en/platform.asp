<!--#include file="conn.asp" -->

<% 	set ris=server.createobject("adodb.recordset")
	if request.QueryString("a")="" then
	sqli="select * from jiedai_qita where id=12"
	else
    sqli="select * from jiedai_qita where id="&request.QueryString("a")
	end if
	 ris.open sqli,conn,1,1
									  
	a_title=ris("e_classid")
	a_body=ris("e_body")
	aid=ris("id")
	bimg=ris("e_tupian")
		  
	 ris.close
	 set ris=nothing
	  if a_body="" then a_body="资料整理更新中···"
									  
	 a_body=LoseStyleTag(a_body) '过滤STYLE	
	 
	 mf="platform"						  
									  
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<TITLE><%=a_title%>_<%=title%></TITLE>
<meta name="keywords" content="<%=keywords_content%>" />
<meta name="description" content="<%=description_content%>" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<script type="text/javascript" src="../js/jquery.min.js"></script>
<script type="text/javascript" src="../js/main.js" ></script>
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
					sql="select * from jiedai_qita where id=12 or id= 13 or id=14 or id= 15 or id=16 order by flag asc"
					res.open sql,conn,1,1
					if aid<>"" then
					i=1
					else
					i=0 
					end if					
					do while not res.eof
					%> 

  <li><a href="platform.asp?a=<%=res("id")%>" <%if aid=int(res("id")) or i=0 then  response.Write"class=""focus""" end if%> ><%=res("e_classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  i=i+1
					  loop
					  res.close
					  set res=nothing
					  %>
	<li><A href="platform_list.asp" >more...</A></li>
            
            
            
        </ul>

    </div>
 <!-- END .nleft-->   
 
 <!-- .ncenter -->
 <div class="ncenter">
	<ul class="nc_title">Home > Platform  <span> > <% =a_title %></span></ul>
   	<%call banner2(bimg)%>
    <ul class="nbody">
    
 <% =a_body %>

    
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
