<!--#include file="session.asp" -->

<link href="images/body.css" rel="stylesheet" type="text/css">
<!--#include FILE="clsUp.asp"-->
<%
dim upfile,formPath,ServerPath,FSPath,formName,FileName,oFile,upfilecount
upfilecount=0
FileRoot=server.MapPath("../uploadfile")&"\"
set upfile=new clsUp ''建立上传对象
upfile.NoAllowExt="tmp;asa;inc;asp;php;js;vb;vbs;html;htm;aspx;sql;txt;cer;png;"
upfile.GetData (102400000)   '取得上传数据,限制最大上传10M
if upfile.isErr then  '如果出错
    select case upfile.isErr
	case 1
	Response.Write "<a href=javascript:history.go(-1)>你没有上传数据呀???是不是搞错了??</a>"
	case 2
	Response.Write "<a href=javascript:history.go(-1)>你上传的文件超出我们的限制,最大10M</a>"
	end select
	else
	FSPath=GetFilePath(Server.mappath("savetofile.asp"),"\")'取得当前文件在服务器路径
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'取得在网站上的位置
	for each formName in upfile.file '列出所有上传了的文件
	   set oFile=upfile.file(formname)
	   FileName=upfile.form(formName)'取得文本域的值
		if oFile.filename="" then
			response.write"<a href=javascript:history.go(-1)>请选择要上传文件</a>!</a>"
			response.end
		end if
	   if not FileName>"" then  FileName=oFile.filename'如果没有输入新的文件名,就用原来的文件名
	'upfile.SaveToFile formname,FSPath&"upfiles/"&FileName   ''保存文件 也可以使用AutoSave来保存,参数一样,但是会自动建立新的文件名
	SFileName=upfile.AutoSave(formname,FileRoot&FileName)
	'将上传信息写入数据库
	'dim conn,db,uppath,sql,FileExt
	'Set conn=Server.CreateObject("ADODB.Connection")
	'db="admin/DBsf4f34jwekks234ds/DBdff344657dfg5tggdjjh.mdb"
	'conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(""&db&"") 
	'conn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath(""&db&"")
	'uppath="upfile/"&FileName
	'sql = "insert into upload(upfile,filename,filetype,filesize) values ('"& uppath &"','"&oFile.FileName&"','"& oFile.FileExt &"',"&oFile.filesize&")"
	'conn.execute sql

	if upfile.iserr then 
		Response.Write upfile.errmessage
	else
		upfilecount=upfilecount+1
		session("file")=SFileName
		'Response.Write "上传成功：<a href=./"&session("file")&" target=_blank>"&session("file")&"</a>"
		
		response.write"<script>parent.myform.image.value='"& session("file") &"';</script>"
		response.write "<script>alert('上传成功');</script>"
		
	end if
	set oFile=nothing
	next
	end if
	set upfile=nothing  '删除此对象
	function GetFilePath(FullPath,str)
	  If FullPath <> "" Then
		GetFilePath = left(FullPath,InStrRev(FullPath, str))
		Else
		GetFilePath = ""
	  End If
	End function

%>