<!--#include file="conn.asp" -->
<!--#include file="Function_Page.asp" -->
<%
		N=request.QueryString("a")
		set rh=server.CreateObject("adodb.recordset")
		if n<>"" then
		seh="select * from zhibenhui_newsclass where id="&N&""
		else
		seh="select * from zhibenhui_newsclass order by id asc"
		end if
		rh.open seh,conn,1,1
		if not rh.eof then
		N=rh("id")
		a_title=rh("classname")
		aid=rh("id")
		bimg=rh("images")
		end if
		rh.close:set rh=nothing
		if a_title="" then a_title="新闻中心"
		
		
		
		keys=Trim(Request.Form("keyword"))
		
		if keys<>"" then
		a_title="信息搜索"
		aid=999999999999
		end if
		
		
		mf="news"
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
					
	                if aid<>"" then aid=int(aid)
					set res=server.createobject("adodb.recordset")
					sql="select * from zhibenhui_newsclass order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof

					%> 

 <li><a href="news_list.asp?a=<%=res("id")%>" <%if aid=int(res("id"))  then  response.Write"class=""focus""" end if%> ><%=res("classname")%></a></li>
    
    
        
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
	<ul class="nc_title">首页 > 新闻资讯  <span> > <% =a_title %></span></ul>
   	<%call banner2(bimg)%>
    <ul class="nbody">
  
  
    <ul class="news_list" id="newslist" > 
     
     
     
      <%
Set mypage=new xdownpage
mypage.getconn=conn

if keys<>"" then
mypage.getsql="select * from zhibenhui_News where  title like '%"&keys&"%' order by tuijian desc,id desc"
else
mypage.getsql="select * from zhibenhui_News where classid="&N&"  and title<>'' order by tuijian desc,id desc"
end if



mypage.pagesize=15
set rs=mypage.getrs()
for i=1 to mypage.pagesize
if not rs.eof then
ntis= rs("title")
ntis=replace(ntis,keys,"<font style='color:red'>"&keys&"</font>")
%>
			
		<li><a href="news_show.asp?id=<% =rs("id") %>" ><% =got(rs("title"),60) %></a><span>[<% =FormatDate(rs("addtime"),4) %>]</span></li>
				
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
    	
        
        <!--#include file="cright.asp" -->

        
    </div>
    <!-- END .nright -->
    
    
 <div class="clearfix"></div>
</Div>




<!--#include file="foot.asp" -->


</body>
</html>
