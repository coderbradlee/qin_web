
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
   	    <a href="job.asp">������ʿ</a> | <a href="map.asp">��վ��ͼ</a> &nbsp;&nbsp;<a href="index.asp">��������</a> | <a href="en/">English</a>
      </ul>
       
    </div>
    
    <div class="nav_logo">
   	  <h1 class="logo"><a href="index.asp" class="a100" >�Ǳ���</a></h1>
        <div class="menus">
        
        	<li class="home"><a href="index.asp" class="ap<% if mf="" or mf="index"  then response.Write"2"%>" >��ҳ</a></li>
            <li><a href="about.asp" class="ap<% if  mf="about"  then response.Write"2"%>" >��������</a><ul>
            
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
	<li><A href="honor.asp" >����������</A></li>
            
            
            
            
            <div></div>
            </ul></li>
            <li><a href="news.asp" class="ap<% if  mf="news"  then response.Write"2"%>" >������Ѷ</a><ul>
            
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
         
              <div></div>
            
            </ul></li>
            <li><a href="case.asp" class="ap<% if  mf="case"  then response.Write"2"%>" >�Ǳ�����ѧԺ</a><ul>
            
              <%
					
	              
					set res=server.createobject("adodb.recordset")
					sql="select * from jd_caseclass order by flag asc"
					res.open sql,conn,1,1				
					do while not res.eof

					%>  <li><a href="case.asp?a=<%=res("id")%>"  ><%=res("classname")%></a></li>
<%
					  res.movenext
					  loop
					  res.close
					  set res=nothing
					  %>
              <div></div>
            </ul></li>
            <li><a href="wenhua.asp" class="ap<% if  mf="wenhua"  then response.Write"2"%>" >��ҵ�Ļ�</a><ul>
            
            
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
	<li><A href="wenhua_list.asp" >�Ļ�����</A></li>
             <div></div>
            </ul></li>
            <li><a href="contact.asp" class="ap<% if  mf="contact"  then response.Write"2"%>" >��ϵ����</a></li>
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
bimg="uploadfile/"&bimg
end if
rgs.close
set rgs=nothing
end if
		%>
        <ul class="nbanner"><img src="<%=bimg%>" width="520" height="105" /></ul>
        <%end sub%>



<% sub banner2(bimg)
bimg="uploadfile/"&bimg
 %>

        <ul class="nbanner"><img src="<%=bimg%>" width="520" height="105" /></ul>
<% end sub %>
