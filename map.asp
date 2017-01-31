<!--#include file="conn.asp" -->

<% 
a_title="网站地图"					  
									  
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
					sql="select * from jiedai_yinhang order by flag asc"
					res.open sql,conn,1,1
					do while not res.eof
					%> 

  <li><a href="About.asp?a=<%=res("id")%>"  ><%=res("classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
	<li><A href="honor.asp" >公益与荣誉</A></li>
            
            
            
        </ul>

    </div>
 <!-- END .nleft-->   
 
 <!-- .ncenter -->
 <div class="ncenter">
	<ul class="nc_title">首页  <span> > <% =a_title %></span></ul>
   	<%call banner(198)%>
    <ul class="nbody">
    
<style>
._map{ width:100%; padding-top:5px; padding-bottom:30px;}
._map li{ float:left; width:100%; padding-bottom:5px; padding-top:5px; border-bottom:1px solid #DFDFDF;}
._map li.sbg{ background:#EFEFEF}
._map li a.ap{ display:block; width:70px; text-indent:4px; font-family:"微软雅黑"; font-size:14px; height:30px; line-height:30px;  float:left}
._map li ul{ float:left;height:30px; line-height:32px;}
._map li ul li{ float:left; width:75px; margin-left:3px; display:inline; padding:0; text-align:center; border:0; }
._map li ul li a{ color:#9B9B9B}
._map li ul li a:hover{ color:#F30}
</style>





<div class="_map" >
        
        	<li class="home"><a href="index.asp" class="ap" >首页</a></li>
            <li class="sbg"><a href="about.asp" class="ap" >关于我们</a><ul>
            
              <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_yinhang order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof
					%> 

  <li><a href="About.asp?a=<%=res("id")%>"  ><%=res("classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
	<li><A href="honor.asp" >公益与荣誉</A></li>
            
  
            </ul></li>
            <li><a href="news.asp" class="ap" >新闻资讯</a><ul>
            
                       	  <%
		
					set res=server.createobject("adodb.recordset")
					sql="select * from zhibenhui_newsclass order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 

 <li><a href="news_list.asp?a=<%=res("id")%>"  ><%=res("classname")%></a></li>
    
    
        
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
         
    
            
            </ul></li>

            <li class="sbg"><a href="academy.asp" class="ap" >商学院</a><ul>
            
              <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from academy_newsclass order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof

					%>  <li><a href="academy_list.asp?a=<%=res("id")%>"  ><%=res("classname")%></a></li>
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
  
            </ul></li>
            <li><a href="platform.asp" class="ap" >公众平台</a><ul>
            
                       	  <%
		
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=12 or id= 13 or id=14 or id= 15 or id=16 order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 

 <li><a href="platform.asp?a=<%=res("id")%>"  ><%=res("classid")%></a></li>
    
    
        
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
         
    
            
            </ul></li>

            <li class="sbg"><a href="forum.asp" class="ap" >交易社区</a><ul>
            
              <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=17 or id= 18 or id=19 or id= 20 order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof

					%>  <li><a href="forum.asp?a=<%=res("id")%>"  ><%=res("classid")%></a></li>
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
  
            </ul></li>
            <li class="sbg"><a href="partner.asp" class="ap" >合伙人</a><ul>
            
              <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=9 or id= 10 or id=11 order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof

					%>  <li><a href="partner.asp?a=<%=res("id")%>"  ><%=res("classid")%></a></li>
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
  
            </ul></li>
            <li class="sbg"><a href="contact.asp" class="ap" >联系我们</a></li>
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
