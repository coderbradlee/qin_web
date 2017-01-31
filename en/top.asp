
<div class="header">
	<div class="top huise">
   	  <ul>
          <form id="form1" name="form1" method="post" action="news_list.asp">
          	<li>
          	  <input type="text" name="keyword" id="keyword" class="skey" />
          	</li>
            <li style="padding-top:2px;">
              <input type="image" name="imageField" id="imageField" src="images/sgo.gif" />
            </li>
          </form>
      </ul>
   	  <ul>
   	    <a href="job.asp">Jobs</a> | <a href="map.asp">Site Map</a> &nbsp;&nbsp;<a href="../">¼òÌåÖÐÎÄ</a> | <a href="index.asp">English</a>
      </ul>
       
    </div>
    
    <div class="nav_logo">
   	  <h1 class="logo"><a href="index.asp" class="a100" >Capital</a></h1>
        <div class="menus">
        
        	<li class="home"><a href="index.asp" class="ap<% if mf="" or mf="index"  then response.Write"2"%>" >Home</a></li>
            <li><a href="about.asp" class="ap<% if  mf="about"  then response.Write"2"%>" >About Us</a><ul style="width:170px;">
            
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
            
            
            
            
            <div></div>
            </ul></li>
            <li><a href="news.asp" class="ap<% if  mf="news"  then response.Write"2"%>" >News</a><ul>
            
                       	  <%
		
					set res=server.createobject("adodb.recordset")
					sql="select * from zhibenhui_newsclass where e_classname<>'' order by flag asc"
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
         
              <div></div>
            
            </ul></li>
            <li style="width:117px; margin-left:25px;"><a href="platform.asp" class="ap<% if  mf="platform"  then response.Write"2"%>" >platform</a><ul style="margin-left:-10px;">
            
            
                          <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=12 or id= 13 or id=14 or id= 15 or id=16 order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof
					%> 

  <li><a href="platform.asp?a=<%=res("id")%>"  ><%=res("e_classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
            </ul></li>
            <li style="width:117px; margin-left:25px;"><a href="forum.asp" class="ap<% if  mf="forum"  then response.Write"2"%>" >forum</a><ul style="margin-left:-10px;">
            
            
                          <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=17 or id= 18 or id=19 or id= 20 order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof
					%> 

  <li><a href="forum.asp?a=<%=res("id")%>"  ><%=res("e_classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
            </ul></li>
            <li style="margin-left:9px;"><a href="case.asp" class="ap<% if  mf="case"  then response.Write"2"%>" >Portfolio</a><ul>
            
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
              <div></div>
            </ul></li>
            <li style="width:117px; margin-left:25px;"><a href="partner.asp" class="ap<% if  mf="partner"  then response.Write"2"%>" >partner</a><ul style="margin-left:-10px;">
            
            
                          <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jiedai_qita where id=9 or id= 10 or id= 11 order by flag asc"
					res.open sql,conn,1,1	
					do while not res.eof
					%> 

  <li><a href="partner.asp?a=<%=res("id")%>"  ><%=res("e_classid")%></a></li>
    
    
        
          <%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
             <div></div>
            </ul></li>
            <li style="margin-left:35px;"><a href="contact.asp" class="ap<% if  mf="contact"  then response.Write"2"%>" >Contact Us</a></li>
        </div>
    </div>
    
</div>


 <%
sub banner(b_id)
if b_id="" then
b_img=""
else
set rgs=conn.execute("select * from jiedai_img where id="&b_id&"")
if not rgs.eof then
bimg=rgs("tupian")
bimg="../uploadfile/"&bimg
end if
rgs.close
set rgs=nothing
end if
		%>
        <ul class="nbanner"><img src="<%=bimg%>" width="520" height="105" /></ul>
        <%end sub%>



<% sub banner2(bimg)
bimg="../uploadfile/"&bimg
 %>

        <ul class="nbanner"><img src="<%=bimg%>" width="520" height="105" /></ul>
<% end sub %>
