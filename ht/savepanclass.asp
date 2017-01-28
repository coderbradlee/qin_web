<!--#include file="conn.asp"-->
<%dim action,anclassid
anclassid=request.QueryString("id")
action=request.querystring("action")
select case action
'//添加新数据
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sh_sort",conn,1,3
rs.AddNew
rs("anclass")=trim(request("anclass2"))
rs("e_anclass")=trim(request("e_anclass2"))
rs("anclassidorder")=int(request("anclassidorder2"))
rs("fudongjia")=int(request("fudongjia2"))
'rs("tupian")=request("tupian")
rs("bid")=request("classid")
rs.Update
rs.Close
set rs=nothing
response.Redirect "anclass.asp"
'//修改数据
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from sh_sort where anclassid="&anclassid,conn,1,3
rs("anclass")=trim(request("anclass"))
rs("e_anclass")=trim(request("e_anclass"))
rs("anclassidorder")=int(request("anclassidorder"))
rs("fudongjia")=int(request("fudongjia"))
rs("tupian")=request("tupian")
rs("bid")=request("classid")

rs.Update
rs.Close
set rs=nothing
response.Redirect "anclass.asp"
'//删除数据
case "del"
conn.execute ("delete from sh_sort where anclassid="&anclassid)
conn.execute ("delete from sh_sort2 where anclassid="&anclassid)
conn.execute ("delete from product where anclassid="&anclassid)
response.Redirect "anclass.asp"
end select
%>