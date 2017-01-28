<!--#include file="conn.asp" -->

<% 
a_title="Site Map"					  
									  
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
					sql="select * from jiedai_yinhang where e_classid<>'' order by flag asc"
					res.open sql,conn,1,1
					do while not res.eof
					%> 

  <li><a href="About.asp?a=<%=res("id")%>"  ><%=res("e_classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
	<li><A href="honor.asp" >Public Welfare and Honour</A></li>
            
            
            
        </ul>

    </div>
 <!-- END .nleft-->   
 
 <!-- .ncenter -->
 <div class="ncenter">
	<ul class="nc_title">Home  <span> > <% =a_title %></span></ul>
   <%call banner(197)%>
   	<ul class="nbody">
    
<style>
._map{ width:100%; padding-top:5px; padding-bottom:30px;}
._map li{ float:left; width:100%; padding-bottom:5px; padding-top:5px; border-bottom:1px solid #DFDFDF;}
._map li.sbg{ background:#EFEFEF}
._map li a.ap{ display:block; width:100%; text-indent:10px; text-decoration:underline;  font-size:14px; height:25px; line-height:25px;  float:left}
._map li ul{ float:left; line-height:32px;}
._map li ul li{ float:left; width:auto; margin-left:10px; display:inline; padding:0; text-align:center; border:0; }
._map li ul li a{ color:#9B9B9B}
._map li ul li a:hover{ color:#F30}
</style>





<div class="_map" >
        
        	<li class="home"><a href="index.asp" class="ap" >Home</a></li>
            <li class="sbg"><a href="about.asp" class="ap" >About Us</a><ul>
            
              <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_yinhang where e_classid<>'' order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof
					%> 

  <li><a href="About.asp?a=<%=res("id")%>"  ><%=res("e_classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
   
	<li style="width:155px; text-align:left" ><A href="honor.asp"  >Public Welfare and Honour</A></li>
            
  
            </ul></li>
            <li><a href="news.asp" class="ap" >News</a><ul>
            
                       	  <%
		
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_newsclass where e_classname<>'' order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 

 <li><a href="news_list.asp?a=<%=res("id")%>"  ><%=res("e_classname")%></a></li>
    
    
        
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
         
    
            
            </ul></li>
            <li class="sbg"><a href="case.asp" class="ap" >Portfolio</a><ul>
            
              <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jd_caseclass where e_classname<>'' order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof

					%>  <li><a href="case.asp?a=<%=res("id")%>"  ><%=res("e_classname")%></a></li>
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
  
            </ul></li>
            <li><a href="wenhua.asp" class="ap" >Corperate Culture</a><ul>
            
            
                          <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=7 or id= 8 order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof
					%> 

  <li><a href="wenhua.asp?a=<%=res("id")%>"  ><%=res("e_classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
	<li ><A href="wenhua_list.asp" >Culture Activities</A></li>
 
            </ul></li>
            <li class="sbg"><a href="contact.asp" class="ap" >Contact Us</a></li>
      </div>






    
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
