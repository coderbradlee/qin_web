<!--#include file="session.asp" -->

<link href="images/body.css" rel="stylesheet" type="text/css">
<!--#include FILE="clsUp.asp"-->
<%
dim upfile,formPath,ServerPath,FSPath,formName,FileName,oFile,upfilecount
upfilecount=0
FileRoot=server.MapPath("../uploadfile")&"\"
set upfile=new clsUp ''�����ϴ�����
upfile.NoAllowExt="tmp;asa;inc;asp;php;js;vb;vbs;html;htm;aspx;sql;txt;cer;png;"
upfile.GetData (102400000)   'ȡ���ϴ�����,��������ϴ�10M
if upfile.isErr then  '�������
    select case upfile.isErr
	case 1
	Response.Write "<a href=javascript:history.go(-1)>��û���ϴ�����ѽ???�ǲ��Ǹ����??</a>"
	case 2
	Response.Write "<a href=javascript:history.go(-1)>���ϴ����ļ��������ǵ�����,���10M</a>"
	end select
	else
	FSPath=GetFilePath(Server.mappath("savetofile.asp"),"\")'ȡ�õ�ǰ�ļ��ڷ�����·��
	ServerPath=GetFilePath(Request.ServerVariables("HTTP_REFERER"),"/")'ȡ������վ�ϵ�λ��
	for each formName in upfile.file '�г������ϴ��˵��ļ�
	   set oFile=upfile.file(formname)
	   FileName=upfile.form(formName)'ȡ���ı����ֵ
		if oFile.filename="" then
			response.write"<a href=javascript:history.go(-1)>��ѡ��Ҫ�ϴ��ļ�</a>!</a>"
			response.end
		end if
	   if not FileName>"" then  FileName=oFile.filename'���û�������µ��ļ���,����ԭ�����ļ���
	'upfile.SaveToFile formname,FSPath&"upfiles/"&FileName   ''�����ļ� Ҳ����ʹ��AutoSave������,����һ��,���ǻ��Զ������µ��ļ���
	SFileName=upfile.AutoSave(formname,FileRoot&FileName)
	'���ϴ���Ϣд�����ݿ�
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
		'Response.Write "�ϴ��ɹ���<a href=./"&session("file")&" target=_blank>"&session("file")&"</a>"
		
		response.write"<script>parent.myform.image.value='"& session("file") &"';</script>"
		response.write "<script>alert('�ϴ��ɹ�');</script>"
		
	end if
	set oFile=nothing
	next
	end if
	set upfile=nothing  'ɾ���˶���
	function GetFilePath(FullPath,str)
	  If FullPath <> "" Then
		GetFilePath = left(FullPath,InStrRev(FullPath, str))
		Else
		GetFilePath = ""
	  End If
	End function

%>