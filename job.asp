<!--#include file="conn.asp" -->

<% a_title="招贤纳士"
	 
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
   	<%call banner(206)%>
    <ul class="nbody">
    
	  <p>    海德邦和是一个富有理想和激情的团队，充满着追求创新的进取精神和蓬勃向上的朝气，公司提供给员工的是公平竞争的环境，是可持续发展的机会和空间，是丰厚的薪金和良好的福利，在这里工作您的个性能够得到充分发挥。这里有充满乐趣的工作、互相信任的同事，海德邦和欢迎您的加盟！</p>
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
					sql="select * from jd_job order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof
					%> 

    
     	<li>
        	<h1><%=res("classid")%></h1>
            <ul><% =res("body") %></ul>
        </li> 
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
        
      </div>
      
      
      
 
      
<TABLE width=100% border=0 cellPadding=3 cellSpacing=1 borderColor=#d3d3d9 bgColor="#DFDFDF" style="margin-top:30px;" >
<TBODY>
<TR>
<TD width="100%" height=30 bgcolor="#FFFFFF" style="border-bottom:2px solid #BCBCBC; padding-left:15px;"><STRONG><FONT style="COLOR: #003366">&nbsp;联系我们</FONT></STRONG></TD></TR>
<TR>
<TD bgcolor="#FFFFFF" style="LINE-HEIGHT: 35px; PADDING-LEFT: 40px; PADDING-TOP: 13px"><STRONG>上海市长宁区宣化路300号，华宁国际广场南塔16层</STRONG> <BR>电 话： +86 21 6167 7313<BR>传 真： +86 21 6167 7286<BR>haidecapital@haidecapital.com</TD></TR></TBODY></TABLE>
<p style="font-weight:bold">&nbsp;</p>
<p style="font-weight:bold">同时，我们投资的被投企业也为您的加入敞开大门。 点击 <a href="case.asp" target="_blank">智本汇商学院</a> 为您的职业生涯寻找最适合的起点。</p>
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
