<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<!--#include file="functions.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<LINK href="images/style.css" type=text/css rel=stylesheet>
<title>�ޱ����ĵ�</title></head>

<SCRIPT language=javascript>
var curObj= null;
function document_onclick() {
if(window.event.srcElement.tagName=='A'||window.event.srcElement.tagName=='FONT'){

if(curObj!=null)
//curObj.style.color="";
curObj.style.fontWeight="normal";
curObj=window.event.srcElement;
//curObj.style.color="#EF1802";
curObj.style.fontWeight="bold";
}
}
</SCRIPT>

<SCRIPT language=JavaScript src="js/alt.js"></SCRIPT>
<BODY onclick=document_onclick();>

<TABLE height="100%" cellSpacing=0 cellPadding=0 width=230 border=0>
  <TBODY>
  <TR>
    <TD 
    style="BACKGROUND-IMAGE: url(images/admin-leftbg.gif); BACKGROUND-REPEAT: repeat-y" 
    vAlign=top align=middle>
      <DIV 
      style="MARGIN: auto; OVERFLOW: auto; WIDTH: 190px; HEIGHT: 100%; align: center">
      <TABLE cellSpacing=0 cellPadding=0 width=160 border=0>
        <TBODY>
        <TR>
          <TD><TABLE 
            style="BORDER-RIGHT: #ddedf6 1px solid; BORDER-TOP: #ddedf6 1px solid; BORDER-LEFT: #ddedf6 1px solid; BORDER-BOTTOM: #ddedf6 1px solid; BACKGROUND-COLOR: #ffffff" 
            height=25 cellSpacing=5 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD><B><IMG src="images/cms-ico2.gif" 
                  align=absMiddle>&nbsp;&nbsp;&nbsp;</B><strong><font color="#ff3300">������Ϣ</font></strong></TD>
              </TR>
              </TBODY></TABLE>
            <TABLE style="MARGIN-TOP: 10px; MARGIN-BOTTOM: 10px" cellSpacing=0 
            cellPadding=4 width="100%" border=0>
              <TR>
                <TD><img src="images/plus-.gif" align=absMiddle> <a 
                  href="config.asp" 
                  target=mainFrame>��������</a></TD>
              </TR>
              
              




<!--	<TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="jiedai_banner.asp?action=list" 
                  target=mainFrame>��ҳ��ͼ����</A> | <A 
                  href="jiedai_banner.asp?action=add" 
                  target=mainFrame>���</A></TD>
                      </TR>
					   

-->



             	<TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="jiedai_img.asp?action=list" 
                  target=mainFrame>��ͼ����</A></TD>
                      </TR>    
					   
              
              
              
           
             					  
   <% 
	set rs=server.createobject("adodb.recordset")
	sql="select * from jiedai_qita   order by flag asc"
	rs.open sql,conn,1,1
	do while not rs.eof 
	%>
                <TR>
                  <TD><img src="images/plus-.gif" align=absMiddle> <a 
                  href="jiedai_qita.asp?action=edit&id=<%=rs("id")%><% if rs("id")=5 or rs("id")=2 then %>&noimg=yes<% end if %>" 
                  target=mainFrame><%=rs("classid")%></a>   <%if rs("e_body")<>"" then 
		   response.Write"" 
		   else
		   response.Write"<img src=""images/noen.jpg"" />" 
		   end if
		   %></TD>
                </TR>
                <%rs.movenext
		  loop
		%>              
              
              
              
             
             
             </TABLE>
		   
		   
		   
               
                  
                  
                  
                  
                  
                  
                  
                  
                                <TABLE 
            style="BORDER-RIGHT: #ddedf6 1px solid; BORDER-TOP: #ddedf6 1px solid; BORDER-LEFT: #ddedf6 1px solid; BORDER-BOTTOM: #ddedf6 1px solid; BACKGROUND-COLOR: #ffffff" 
            height=25 cellSpacing=5 cellPadding=0 width="100%" border=0>
                    <TBODY>
                      <TR>
                        <TD><B><IMG src="images/cms-ico1.gif" 
                  align=absMiddle>&nbsp;&nbsp;&nbsp;</B><FONT 
                  color=#ff3300><B>��������</B></FONT></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE style="MARGIN-TOP: 10px; MARGIN-BOTTOM: 10px" cellSpacing=0 
            cellPadding=4 width="100%" border=0>
                    <TBODY>
             			  
					  
					  
                      	<TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="jiedai_yinhang.asp?action=list" 
                  target=mainFrame>��˾��Ϣ</A> | <A 
                  href="jiedai_yinhang.asp?action=add" 
                  target=mainFrame>���</A></TD>
                      </TR>  
                      
                      
                      
					  
					<TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="jiedai_honor.asp?action=list" 
                  target=mainFrame>��ҵ����</A> | <A 
                  href="jiedai_honor.asp?action=add" 
                  target=mainFrame>���</A></TD>
                      </TR>  
					  
					  
					  
  	<TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="jd_job.asp?action=list" 
                  target=mainFrame>������ʿ</A> | <A 
                  href="jd_job.asp?action=add" 
                  target=mainFrame>���</A></TD>
                      </TR>  
                      
              

					  
					  
					  
                    </TBODY>
                  </TABLE>
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  
                  <TABLE 
            style="BORDER-RIGHT: #ddedf6 1px solid; BORDER-TOP: #ddedf6 1px solid; BORDER-LEFT: #ddedf6 1px solid; BORDER-BOTTOM: #ddedf6 1px solid; BACKGROUND-COLOR: #ffffff" 
            height=25 cellSpacing=5 cellPadding=0 width="100%" border=0>
                    <TBODY>
                      <TR>
                        <TD><B><IMG src="images/cms-ico1.gif" 
                  align=absMiddle>&nbsp;&nbsp;&nbsp;</B><FONT 
                  color=#ff3300><B>���Ź���</B></FONT></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE style="MARGIN-TOP: 10px; MARGIN-BOTTOM: 10px" cellSpacing=0 
            cellPadding=4 width="100%" border=0>
                    <TBODY>
                      
					  
					  
		     <!--<TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A href="newsanclass.asp" 
                  target=mainFrame>������</A></TD>
                      </TR>
					  
					  
                      <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A href="jiedai_news.asp?action=add" 
                  target=mainFrame>��Ѷ�����ţ���� ��</A></TD>
                      </TR>
                      <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="jiedai_news.asp?action=list" 
                  target=mainFrame>��Ѷ�����ţ�����</A></TD>
                      </TR>-->
                        <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A href="zhibenhui_news.asp?action=add" 
                  target=mainFrame>�Ǳ������� ��ӡ�</A></TD>
                      </TR>
                      <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="zhibenhui_news.asp?action=list" 
                  target=mainFrame>�Ǳ������� ����</A></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  
                  
                  
                  
                  
                  
                  
                  
                   
                  <TABLE 
            style="BORDER-RIGHT: #ddedf6 1px solid; BORDER-TOP: #ddedf6 1px solid; BORDER-LEFT: #ddedf6 1px solid; BORDER-BOTTOM: #ddedf6 1px solid; BACKGROUND-COLOR: #ffffff" 
            height=25 cellSpacing=5 cellPadding=0 width="100%" border=0>
                    <TBODY>
                      <TR>
                        <TD><B><IMG src="images/cms-ico1.gif" 
                  align=absMiddle>&nbsp;&nbsp;&nbsp;</B><FONT 
                  color=#ff3300><B>�Ǳ�����ѧԺ</B></FONT></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE style="MARGIN-TOP: 10px; MARGIN-BOTTOM: 10px" cellSpacing=0 
            cellPadding=4 width="100%" border=0>
                    <TBODY>
                      
					  
					  
		     <!--<TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A href="caseanclass.asp" 
                  target=mainFrame>������</A></TD>
                      </TR>-->
					  
					  
                      <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A href="jd_case.asp?action=add" 
                  target=mainFrame>��Ϣ��� ��</A></TD>
                      </TR>
                      <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="jd_case.asp?action=list" 
                  target=mainFrame>��Ϣ����</A></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  
                  
                 
                  <TABLE 
            style="BORDER-RIGHT: #ddedf6 1px solid; BORDER-TOP: #ddedf6 1px solid; BORDER-LEFT: #ddedf6 1px solid; BORDER-BOTTOM: #ddedf6 1px solid; BACKGROUND-COLOR: #ffffff" 
            height=25 cellSpacing=5 cellPadding=0 width="100%" border=0>
                    <TBODY>
                      <TR>
                        <TD><B><IMG src="images/cms-ico1.gif" 
                  align=absMiddle>&nbsp;&nbsp;&nbsp;</B><FONT 
                  color=#ff3300><B>��ҵ�Ļ�</B></FONT></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <TABLE style="MARGIN-TOP: 10px; MARGIN-BOTTOM: 10px" cellSpacing=0 
            cellPadding=4 width="100%" border=0>
                    <TBODY>
                      
					  
		
					  
					  
                      <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A href="zhibenhui_fuwu.asp?action=add" 
                  target=mainFrame>��Ϣ��� ��</A></TD>
                      </TR>
                      <TR>
                        <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="zhibenhui_fuwu.asp?action=list" 
                  target=mainFrame>��Ϣ����</A></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  
                     
                  
                  
                  
             
                
                  <TABLE 
            style="BORDER-RIGHT: #ddedf6 1px solid; BORDER-TOP: #ddedf6 1px solid; BORDER-LEFT: #ddedf6 1px solid; BORDER-BOTTOM: #ddedf6 1px solid; BACKGROUND-COLOR: #ffffff" 
            height=25 cellSpacing=5 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD><B><IMG src="images/cms-ico3.gif" 
                  align=absMiddle>&nbsp;&nbsp;&nbsp;</B><FONT 
                  color=#ff3300><B>ϵͳ����</B></FONT></TD></TR></TBODY></TABLE>
            <TABLE style="MARGIN-TOP: 10px; MARGIN-BOTTOM: 10px" cellSpacing=0 
            cellPadding=4 width="100%" border=0>
              <TBODY>
              <TR>
                <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="admin_Manage.asp" 
                  target=mainFrame>����Ա�޸�</A></TD></TR>
              <TR>
                <TD><IMG src="images/plus-.gif" align=absMiddle> <A 
                  href="add_admin_Manage.asp" 
                  target=mainFrame>����Ա���</A></TD></TR>
</TBODY></TABLE></TD></TR></TBODY></TABLE>
      </DIV></TD></TR></TBODY></TABLE>




</body>
</html>
