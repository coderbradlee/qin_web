<div class="right_list">
        	<h1>��ϵ����</h1>
            <ul>
         	<dd style="padding-left:16px; padding-right:16px; float:left">
 <%
	set ris=server.createobject("adodb.recordset")
	sqli="select * from jiedai_qita where id=5"
	 ris.open sqli,conn,1,1
	if not ris.eof then								  
	i_body=ris("body")
	end if	  
	 ris.close
	 set ris=nothing
	if i_body="" then
	response.Write"������Ϣ"
	else
	response.Write i_body
	end if		
			%>
         </dd>
         
            </ul>
        </div>