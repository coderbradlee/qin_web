<%@LANGUAGE="VBSCRIPT" CODEPAGE="936" %> 
<%Session.CodePage=936%>

<%

'on error resume next
dim provider,path,pass,dsn,conn
provider="provider=microsoft.jet.oledb.4.0;"
path="data source=" & server.mappath("../jdshuju/#jiedai.mdb")
pass=";jet oledb:database password="
dsn=provider&path&pass
set conn=server.createobject("adodb.connection")
conn.open dsn

function checkStr(str)
str=replace(str,"'","")
Str=Replace(Str,chr(39),"") '����SQLע��
Str=Replace(Str,chr(91),"") '����SQLע��[
Str=Replace(Str,chr(93),"") '����SQLע��]
Str=Replace(Str,chr(37),"") '����SQLע��%
Str=Replace(Str,chr(59),"") '����SQLע��
Str=Replace(Str,chr(43),"") '����SQLע��;
Str=Replace(Str,chr(45),"") '����SQLע��+
Str=Replace(Str,chr(123),"") '����SQLע��{
Str=Replace(Str,chr(125),"") '����SQLע��}

checkStr=Str '���ؾ��������ַ��滻���Str
if isnull(str) then
checkStr = ""
exit function 
end if
end function

%>
                                                                                                                          
<%


'�����ݿ��ж�����
  
 	set rsxml=server.createobject("adodb.recordset")
	sql="select * from jiedai_img where tuijian=1 and classid<>1 order by id desc"
	rsxml.open sql,conn,1,3

xmlfile=server.mappath("../img.xml")
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile(xmlfile,True)

'------------------����xml

'Set f = Server.CreateObject("Scripting.FileSystemObject")
'Set myfile = f.Createtextfile(server.mappath(file),true)

   
myfile.writeline "<?xml version=""1.0"" encoding=""utf-8""?>"
myfile.writeline "<data>"
myfile.writeline "<channel>"



do while not rsxml.eof

myfile.writeline "<item>"
'myfile.writeline "<link>"&rsxml("wblink")&"</link>"
myfile.writeline "<image>uploadfile/"&rsxml("tupian")&"</image>"
'myfile.writeline "<title>"&rsxml("title")&"</title>"
myfile.writeline "</item>"

rsxml.movenext
loop
rsxml.close
set rsxml=nothing

myfile.writeline "</channel>"

myfile.writeline "<config>"
myfile.writeline "<autoPlayTime>5</autoPlayTime>"          'ͼƬ�л�ʱ�䣬Ĭ��ֵ��8����λ��

'     myfile.writeline "<isHeightQuality>false</isHeightQuality>"  'ͼƬ��С�Ƿ���ø������ķ�����Ĭ��ֵfalse

'     myfile.writeline "<windowOpen>_self</windowOpen>"      'ͼƬ���ӵĴ򿪷�ʽ��Ĭ��ֵ��_blank��,���´��ڴ򿪣�Ҳ����ʹ�á�_self��,ʹ�ñ����ڴ�

myfile.writeline "<btnSetMargin>auto 10 10 auto</btnSetMargin>"
'��ť��λ�ã����ֵ�λ�ã�����css��margin���Ĭ��ֵ��auto 5 5 auto�����ĸ���ֵ���� �� �� �� �� ����ڲ������ľ��룬�ĸ���ֵ�ÿո�ֿ������������ֵ�á�auto����д ���������϶��벢����10���صľ������д ��10 auto auto 10��, ���½Ƕ����ǡ�auto 10 10 auto��

myfile.writeline "<scaleMode>exactFil</scaleMode>"

'ͼƬ����ģʽ: Ĭ��ֵ�ǡ�noBorder��
'��showAll��: ���Կ���ȫ��ͼƬ,���ֱ���,�������»�������
'��exactFil��: ����ͼƬ����̨�ĳߴ�,���ܱ���ʧ��
'��noScale��: ͼƬ��ԭʼ�ߴ�,�޷���
'��noBorder��: ͼƬ����������,���ֱ���,���ܻᱻ�ü�

myfile.writeline "<changImageMode>click</changImageMode>"

'�л�ͼƬ�ķ�����Ĭ��ֵ��click��,����л�ͼƬ��������ʹ�á�hover��,�����ͣ���л�ͼƬ


myfile.writeline "<isShowBtn>false</isShowBtn>"

'�Ƿ���ʾ��ť��Ĭ��ֵ��true��


myfile.writeline "<transform>breatheBlur</transform>"

'ͼƬ����ģʽ: Ĭ��ֵ�ǡ�alpha��
'��alpha��: ͸���ȵ��뵭��
'��blur��: ģ�����뵭��
'��left��: ��ͼƬ����
'��right��: �ҷ�ͼƬ����
'��top��: �Ϸ�ͼƬ����
'��bottom��: �·�ͼƬ����
'��breathe��: ��һ���ط����ĵ��뵭��
'��breatheBlur��: ��һ���ط�����ģ�����뵭��


'myfile.writeline "<btnDistance>20</btnDistance>"
'ÿ����ť�ľ��룬Ĭ��ֵ20

'myfile.writeline "<titleBgColor>0xff6600</titleBgColor>"
'���ⱳ������ɫ��Ĭ��0xff6600

'myfile.writeline "<titleTextColor>0xffffff</titleTextColor>"

'�������ֵ���ɫ��Ĭ��0xffffff

myfile.writeline "<titleBgAlpha>0.75</titleBgAlpha>"

'���ⱳ����͸���ȣ�Ĭ��0.75

'myfile.writeline "<titleFont>΢���ź�</titleFont>"

'�������ֵ����壬Ĭ��ֵ��΢���źڡ�

'myfile.writeline "<titleMoveDuration>1</titleMoveDuration>"

'���ⱳ��������ʱ�䣬Ĭ��ֵ1����λ��

'myfile.writeline "<btnAlpha>0.7</btnAlpha>"

'��ť��͸���ȣ�Ĭ��ֵ0.7


'myfile.writeline "<btnTextColor>0xffffff</btnTextColor>"

'��ť���ֵ���ɫ��Ĭ��ֵ0xffffff


'myfile.writeline "<btnDefaultColor>0��1B3433</btnDefaultColor>"
'��ť��Ĭ����ɫ��Ĭ��ֵ0��1B3433

'myfile.writeline "<btnHoverColor>0xff9900</btnHoverColor>"

'��ť��Ĭ����ɫ��Ĭ��ֵ0xff9900

'myfile.writeline "<btnFocusColor>0xff6600</btnFocusColor>"

'��ť��ǰ��ɫ��Ĭ��ֵ0xff6600



myfile.writeline "<isShowTitle>false</isShowTitle>"        ' �Ƿ���ʾ���⣬Ĭ��ֵ��true��

'myfile.writeline "<roundCorner>0</roundCorner>"           '����Բ��

myfile.writeline "<isShowAbout>false</isShowAbout>"    '�Ƿ���ʾ������Ϣ��Ĭ��ֵ��true��
myfile.writeline "</config>"

myfile.writeline "</data>"


myfile.close











'�����ݿ��ж�����
  
 	set rsxml=server.createobject("adodb.recordset")
	sql="select * from jiedai_img where tuijian=1 and classid=1 order by id desc"
	rsxml.open sql,conn,1,3

xmlfile=server.mappath("../en/img.xml")
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile(xmlfile,True)

'------------------����xml

'Set f = Server.CreateObject("Scripting.FileSystemObject")
'Set myfile = f.Createtextfile(server.mappath(file),true)

   
myfile.writeline "<?xml version=""1.0"" encoding=""utf-8""?>"
myfile.writeline "<data>"
myfile.writeline "<channel>"



do while not rsxml.eof

myfile.writeline "<item>"
'myfile.writeline "<link>"&rsxml("wblink")&"</link>"
myfile.writeline "<image>../uploadfile/"&rsxml("tupian")&"</image>"
'myfile.writeline "<title>"&rsxml("title")&"</title>"
myfile.writeline "</item>"

rsxml.movenext
loop

myfile.writeline "</channel>"

myfile.writeline "<config>"
myfile.writeline "<autoPlayTime>5</autoPlayTime>"          'ͼƬ�л�ʱ�䣬Ĭ��ֵ��8����λ��

'     myfile.writeline "<isHeightQuality>false</isHeightQuality>"  'ͼƬ��С�Ƿ���ø������ķ�����Ĭ��ֵfalse

'     myfile.writeline "<windowOpen>_self</windowOpen>"      'ͼƬ���ӵĴ򿪷�ʽ��Ĭ��ֵ"_blank",���´��ڴ򿪣�Ҳ����ʹ��"_self",ʹ�ñ����ڴ�

myfile.writeline "<btnSetMargin>auto 10 10 auto</btnSetMargin>"
'��ť��λ�ã����ֵ�λ�ã�����css��margin���Ĭ��ֵ"auto 5 5 auto"���ĸ���ֵ���� �� �� �� �� ����ڲ������ľ��룬�ĸ���ֵ�ÿո�ֿ������������ֵ��"auto"��д ���������϶��벢����10���صľ������д "10 auto auto 10��, ���½Ƕ�����"auto 10 10 auto"

myfile.writeline "<scaleMode>exactFil</scaleMode>"

'ͼƬ����ģʽ: Ĭ��ֵ��"noBorder"
'"showAll": ���Կ���ȫ��ͼƬ,���ֱ���,�������»�������
'"exactFil": ����ͼƬ����̨�ĳߴ�,���ܱ���ʧ��
'"noScale": ͼƬ��ԭʼ�ߴ�,�޷���
'"noBorder": ͼƬ����������,���ֱ���,���ܻᱻ�ü�

myfile.writeline "<changImageMode>click</changImageMode>"

'�л�ͼƬ�ķ�����Ĭ��ֵ"click",����л�ͼƬ��������ʹ��"hover",�����ͣ���л�ͼƬ


myfile.writeline "<isShowBtn>false</isShowBtn>"

'�Ƿ���ʾ��ť��Ĭ��ֵ"true"


myfile.writeline "<transform>breatheBlur</transform>"

'ͼƬ����ģʽ: Ĭ��ֵ��"alpha"
'"alpha": ͸���ȵ��뵭��
'"blur": ģ�����뵭��
'"left": ��ͼƬ����
'"right": �ҷ�ͼƬ����
'"top": �Ϸ�ͼƬ����
'"bottom": �·�ͼƬ����
'"breathe": ��һ���ط����ĵ��뵭��
'"breatheBlur": ��һ���ط�����ģ�����뵭��


'myfile.writeline "<btnDistance>20</btnDistance>"
'ÿ����ť�ľ��룬Ĭ��ֵ20

'myfile.writeline "<titleBgColor>0xff6600</titleBgColor>"
'���ⱳ������ɫ��Ĭ��0xff6600

'myfile.writeline "<titleTextColor>0xffffff</titleTextColor>"

'�������ֵ���ɫ��Ĭ��0xffffff

myfile.writeline "<titleBgAlpha>0.75</titleBgAlpha>"

myfile.writeline "<isShowTitle>false</isShowTitle>"        ' �Ƿ���ʾ���⣬Ĭ��ֵ"true"

myfile.writeline "<isShowAbout>false</isShowAbout>"    '�Ƿ���ʾ������Ϣ��Ĭ��ֵ"true"
myfile.writeline "</config>"

myfile.writeline "</data>"


myfile.close





%>
<meta http-equiv="refresh" content="0;URL=jiedai_img.asp?action=list" />

