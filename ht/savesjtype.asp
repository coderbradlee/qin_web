<!--#include file="conn.asp"-->
<%dim action,anclassid
anclassid=request.QueryString("id")
action=request.querystring("action")
select case action
'//���������
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sjsort",conn,1,3
rs.AddNew
rs("anclass")=trim(request("anclass2"))
rs("anclassidorder")=int(request("anclassidorder2"))
rs("fudongjia")=int(request("fudongjia2"))
rs("bid")=request("classid")
rs("sjid")=request("sjid")
rs.Update
rs.Close
set rs=nothing
response.Redirect "shangjiatype_add.asp?sjid="&request("sjid")&""
'//�޸�����
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sjsort where anclassid="&anclassid,conn,1,3
rs("anclass")=trim(request("anclass"))
rs("anclassidorder")=int(request("anclassidorder"))
rs("fudongjia")=int(request("fudongjia"))
rs("bid")=request("classid")
'response.write request("sjid")
'response.end
rs("sjid")=request("sjid")

rs.Update
rs.Close
set rs=nothing
response.Redirect "shangjiatype_add.asp?sjid="&request("sjid")&""
'//ɾ������
case "del"
conn.execute ("delete from sjsort where anclassid="&anclassid)
'conn.execute ("delete from sh_sort2 where anclassid="&anclassid)
'conn.execute ("delete from product where anclassid="&anclassid)
response.Redirect "shangjiatype_add.asp?sjid="&request("sjid")&""
end select
%>