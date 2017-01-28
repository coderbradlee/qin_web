<!--#include file="conn.asp"-->
<%dim action,anclassid
anclassid=request.QueryString("id")
action=request.querystring("action")
select case action
'//添加新数据
case "add" 
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from jd_caseclass",conn,1,3
rs.AddNew
rs("classname")=trim(request("anclass2"))
rs("e_classname")=trim(request("e_anclass2"))
rs("flag")=int(request("anclassidorder2"))
'rs("fudongjia")=int(request("fudongjia2"))
'rs("bid")=request("classid")
rs("tupian")=trim(request("image"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "caseanclass.asp"
'//修改数据
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from jd_caseclass where id="&anclassid,conn,1,3
rs("classname")=trim(request("anclass"))
rs("e_classname")=trim(request("e_anclass"))
rs("flag")=int(request("anclassidorder"))
'rs("fudongjia")=int(request("fudongjia"))
'rs("bid")=request("classid")
rs("tupian")=trim(request("image"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "caseanclass.asp"
'//删除数据
case "del"
conn.execute ("delete from jd_caseclass where id="&anclassid)
'conn.execute ("delete from sh_sort2 where anclassid="&anclassid)
conn.execute ("delete from jd_case where classid="&anclassid)
response.Redirect "caseanclass.asp"
end select
%>