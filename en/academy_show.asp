<!--#include file="conn.asp" -->
<%

		
Jid=checkStr(request.QueryString("id"))
if not isnumeric(jid) then
response.write"参数有误"
response.end
end if


set ra=server.createobject("adodb.recordset")
sql="select * from academy_news where id="&jid
ra.open sql,conn,1,3
ra("click")=ra("click")+1
N=ra("classid")
ra.update

	set rh=server.CreateObject("adodb.recordset")
		if N<>"" then
		seh="select * from academy_newsclass where id="&N&""
		else
		seh="select * from academy_newsclass order by id asc"
		end if
		rh.open seh,conn,1,1	
		a_title=rh("classname")
		aid=rh("id")
		N=aid
		bimg=rh("tupian")
		rh.close:set rh=nothing
		
		

		
		
		%>	


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<TITLE><%=ra("title")%>_<%=a_title%>_<%=title%></TITLE>
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
					
	                if aid<>"" then aid=int(aid)
					set res=server.createobject("adodb.recordset")
					sql="select * from academy_newsclass order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 

 <li><a href="academy_list.asp?a=<%=res("id")%>" <%if aid=int(res("id"))  then  response.Write"class=""focus""" end if%> ><%=res("classname")%></a></li>
    
    
        
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
   	<%call banner2(bimg)%>
    <ul class="nbody">
  		
							<p style="font-weight:bold; text-align:center"><%=ra("title")%></p>
                            <p><%=ra("content")%></p>
                   
                            
                            


    
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
