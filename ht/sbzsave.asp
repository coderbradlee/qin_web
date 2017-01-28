<!--#include file="conn.asp"-->
<%dim action,anclassid
anclassid=request.QueryString("id")
action=request.querystring("action")
select case action
'//添加新数据
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from xbz",conn,1,3
rs.AddNew
rs("anclass")=trim(request("anclass2"))
rs("anclassidorder")=int(request("anclassidorder2"))
rs("fudongjia")=int(request("fudongjia2"))
rs("bid")=request("classid")
rs.Update
rs.Close
set rs=nothing
response.Redirect "xbz.asp"
'//修改数据
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from xbz where anclassid="&anclassid,conn,1,3
rs("anclass")=trim(request("anclass"))
rs("anclassidorder")=int(request("anclassidorder"))
rs("fudongjia")=int(request("fudongjia"))
rs("bid")=request("classid")

rs.Update
rs.Close
set rs=nothing
response.Redirect "xbz.asp"
'//删除数据
case "del"
conn.execute ("delete from xbz where anclassid="&anclassid)
conn.execute ("delete from jiedai_xyk where anclassid="&anclassid)
response.Redirect "xbz.asp"
end select
%>