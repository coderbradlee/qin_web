<!--#include file="conn.asp" -->
<!--#include file="session.asp" -->
<%dim action,url,i,abc,anclassid,anclass
anclassid=request("anclassid")
anclass=request.querystring("anclass")
url="http://" & request.servervariables("http_host") & finddir(request.servervariables("url"))
action=request.querystring("action")
'//添加新数据


if action="add" and anclass="" then


set rs=server.createobject("adodb.recordset")
rs.open "select * from sh_sort",conn,1,3
rs.addnew


rs("anclass")=trim(request("nclass2"))

'rs("changyong")=int(request("changyong"))
rs.update
rs.close
set rs=nothing
response.redirect url&"nclass.asp"
end if




select case action
case "add"
set rs=server.createobject("adodb.recordset")
rs.open "select * from sh_sort2",conn,1,3
rs.addnew
rs("nclass")=trim(request("nclass2"))
rs("nclassidorder")=int(request("nclassidorder2"))
rs("anclassid")=int(request("anclassid"))
'rs("changyong")=int(request("changyong"))
rs.update
rs.close
set rs=nothing
response.redirect url&"nclass.asp?id="&anclassid&"&anclass="&anclass
'//修改数据

case "edit"
set rs=server.createobject("adodb.recordset")
rs.open "select * from sh_sort2 where nclassid="&request.querystring("id"),conn,1,3
rs("nclass")=trim(request("nclass"))
rs("nclassidorder")=int(request("nclassidorder"))
rs("anclassid")=request("matype")
'rs("changyong")=int(request("changyong"))
rs.update
rs.close
set rs=nothing
response.redirect url&"nclass.asp?id="&anclassid&"&anclass="&anclass
'//删除数据
case "del"
anclassid=request.querystring("anclassid")
conn.execute ("delete from sh_sort2 where nclassid="&request.querystring("id"))
conn.execute ("delete from product where nclassid="&request.querystring("id"))
response.redirect url&"nclass.asp?id="&anclassid&"&anclass="&anclass
end select
%>
<%
function finddir(filepath)
	finddir=""
	for i=1 to len(filepath)
	if left(right(filepath,i),1)="/" or left(right(filepath,i),1)="\" then
	  abc=i
	  exit for
	end if
	next
	if abc <> 1 then
	finddir=left(filepath,len(filepath)-abc+1)
	end if
end function
%>
                                                                                                                          