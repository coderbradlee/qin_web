<!--#include file="conn.asp" -->

<% a_title="Jobs"
	 
	 mf="contact"						  
									  
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
	<ul class="nc_title">Home  <span> > <% =a_title %></span></ul><ul></ul>
   <%call banner(205)%>
    <ul class="nbody">
    
	  <p>    At Haide Capital, you can find an innovative team full of vitality, ideals and passion. Besides the satisfactory remuneration and benefits, what will make you love to work here are opportunities for designing your career path, a stage for fair competition, a friendly climate, co-operative and trustworthy team fellows... It is at Haide Capital that you feel like coming back home enjoying the interesting and challenging cause with sisters and brothers. Do come and join us.</p>
	  <p>&nbsp;</p>
	  <script>
$(document).ready(function(){
$(".job_list li ul:not(:first)").hide(); 
$(".job_list li:first h1").addClass("jian");
// $("dd:not(:last)").hide();  //试试$("dd:not(:last)").hide();
$(".job_list li h1").click(function(){
var pcl=$(this).attr("class");
if(pcl!="jian"){
$(".job_list li ul").slideUp("slow");
$(".job_list li h1").removeClass("jian");
$(this).addClass("jian");
$(this).parent().find('ul:eq(0)').slideDown("slow");
};	
return false;
});
 });
 </script>
      <style>
      	.job_list{ width:100%; padding-top:8px; padding-bottom:15px;}
		.job_list li{ width:100%; float:left;}
		.job_list li h1{ width:100%; background:url(images/jia.gif) no-repeat 3px center; height:29px; border-bottom:1px dotted #B7B7B7; font-size:12px; line-height:29px; text-indent:20px; cursor:pointer}
		.job_list li h1.jian{ background:url(images/shou.gif) no-repeat 3px center;}
		.job_list li ul{ width:99%; margin:0 auto; padding-top:10px; padding-bottom:15px;}
      </style>
      
      <div class="job_list">
          
                        <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jd_job where e_classid<>'' order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof
					%> 

    
     	<li>
        	<h1><%=res("e_classid")%></h1>
            <ul><% =res("e_body") %></ul>
        </li> 
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
        
      </div>
      
      
      
      
<TABLE width=100% border=0 style="margin-top:30px;" >
<TBODY>
<TR>
<TD width="100%" >
<IMG border=0 src="/uploadfile/20110621104922212.jpg">
</TD>
</TR></TBODY></TABLE>
<p style="font-weight:bold">&nbsp;</p>
<p style="font-weight:bold">At the same time, the invested enterprises also have opened the door for you to join. Click the “<a href="case.asp" target="_blank">Portfolio</a>” to find the most suitable starting point for your career. </p>
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
