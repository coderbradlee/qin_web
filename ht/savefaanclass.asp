<!--#include file="conn.asp"-->
<%dim action,anclassid
anclassid=request.QueryString("id")
action=request.querystring("action")
select case action
'//添加新数据
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from jiedai_faclass",conn,1,3
rs.AddNew
rs("classname")=trim(request("anclass2"))
rs("flag")=int(request("anclassidorder2"))
'rs("fudongjia")=int(request("fudongjia2"))
'rs("bid")=request("classid")
rs.Update
rs.Close
set rs=nothing
response.Redirect "faanclass.asp"
'//修改数据
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from jiedai_faclass where id="&anclassid,conn,1,3
rs("classname")=trim(request("anclass"))
rs("flag")=int(request("anclassidorder"))
'rs("fudongjia")=int(request("fudongjia"))
'rs("bid")=request("classid")

rs.Update
rs.Close
set rs=nothing
response.Redirect "faanclass.asp"
'//删除数据
case "del"
conn.execute ("delete from jiedai_faclass where id="&anclassid)
'conn.execute ("delete from sh_sort2 where anclassid="&anclassid)
conn.execute ("delete from jiedai_fangan where classid="&anclassid)
response.Redirect "faanclass.asp"
end select
%>