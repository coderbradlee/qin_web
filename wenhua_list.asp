<!--#include file="conn.asp" -->
<!--#include file="Function_Page.asp" -->

<%					  
	a_title="文化生活"

	 
	 mf="wenhua"						  
									  
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
					sql="select * from jiedai_qita where id=7 or id= 8 order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof
					%> 

  <li><a href="wenhua.asp?a=<%=res("id")%>"  ><%=res("classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
	<li><A href="wenhua_list.asp" class="focus" >文化生活</A></li>
            
            
            
        </ul>

    </div>
 <!-- END .nleft-->   
 
 <!-- .ncenter -->
 <div class="ncenter">
	<ul class="nc_title">首页 > 企业文化  <span> > <% =a_title %></span></ul>
   	<%call banner(204)%>
    <ul class="nbody">
    


<ul class="news_list" id="newslist" > 
     
     
     
      <%
Set mypage=new xdownpage
mypage.getconn=conn
mypage.getsql="select * from jiedai_fuwu where  classid<>'' order by  flag asc"
mypage.pagesize=15
set rs=mypage.getrs()
for i=1 to mypage.pagesize
if not rs.eof then 
%>
			
		<li><a href="wenhua_show.asp?id=<% =rs("id") %>" ><% =rs("classid") %></a><span>[<% =FormatDate(rs("addtime"),4) %>]</span></li>
				
<%
rs.movenext
else
exit for	 
end if
next
%>
	
     
     </ul>


<div class="clearfix n8"></div>
<div class="quotes">
    <%=mypage.showpage()%>

</div>

<%
rs.close
set rs=nothing
'end if
%> 
        
   <script language="javascript">showtable('newslist','li','#eaeef1')</script>  




    
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
