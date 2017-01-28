<!--#include file="conn.asp" -->
<!--#include file="Function_Page.asp" -->
<%
		N=request.QueryString("a")
		set rh=server.CreateObject("adodb.recordset")
		if n<>"" then
		seh="select * from jd_caseclass where id="&N&""
		else
		seh="select * from jd_caseclass where e_classname<>'' order by id asc"
		end if
		rh.open seh,conn,1,1
		if not rh.eof then
		N=rh("id")
		a_title=rh("e_classname")
		aid=rh("id")
		end if
		rh.close:set rh=nothing
		if a_title="" then a_title="Portfolio"
		
		mf="case"
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
					
	                if aid<>"" then aid=int(aid)
					set res=server.createobject("adodb.recordset")
					sql="select * from jd_caseclass where e_classname<>'' order by flag asc"
					res.open sql,conn,1,1
					
					if aid<>"" then
					i=1
					else
					i=0 
					end if					
					do while not res.eof

					%> 

 <li><a href="case.asp?a=<%=res("id")%>"   <%if aid=int(res("id")) or i=0 then  response.Write"class=""focus""" end if%> ><%=res("e_classname")%></a></li>
    
    
        
<%
					  res.movenext
					  i=i+1
					  loop
					  res.close
					  set res=nothing
					  %>
         
            
            
        </ul>

    </div>
 <!-- END .nleft-->   
 
 <!-- .ncenter -->
 <div class="ncenter">
	<ul class="nc_title">Home > Portfolio  <span> > <% =a_title %></span></ul>
   	<%call banner(201)%>
    <ul class="nbody" style="width:100%;">
  
  <style>
  	.news_hti{ background:#D4D4D4}
	.news_hti h1{ font-weight:normal; color:#000; padding-left:20px;}
	.news_main .fl img{ width:133px; height:58px; border:0}
  </style>
  
    	  <%
					set res=server.createobject("adodb.recordset")
					sql="select * from jd_case where e_title<>'' order by id desc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 
  
  
	  <div class="news_home">
 			 <ul class="news_hti"><h1><% =res("e_title") %></h1> <a href="http://<% =res("wblink") %>" target="_blank" ><% =res("wblink") %></a></ul>
          <ul class="news_main">
           	  <div class="fl"><img src="../uploadfile/<% =res("tupian") %>"  /></div>
              <div class="fr"><% =res("e_content") %></div>
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
