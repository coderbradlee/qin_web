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
Str=Replace(Str,chr(39),"") '过滤SQL注入
Str=Replace(Str,chr(91),"") '过滤SQL注入[
Str=Replace(Str,chr(93),"") '过滤SQL注入]
Str=Replace(Str,chr(37),"") '过滤SQL注入%
Str=Replace(Str,chr(59),"") '过滤SQL注入
Str=Replace(Str,chr(43),"") '过滤SQL注入;
Str=Replace(Str,chr(45),"") '过滤SQL注入+
Str=Replace(Str,chr(123),"") '过滤SQL注入{
Str=Replace(Str,chr(125),"") '过滤SQL注入}

checkStr=Str '返回经过上面字符替换后的Str
if isnull(str) then
checkStr = ""
exit function 
end if
end function

%>
                                                                                                                          
<%


'从数据库中读数据
  
 	set rsxml=server.createobject("adodb.recordset")
	sql="select * from jiedai_img where tuijian=1 and classid<>1 order by id desc"
	rsxml.open sql,conn,1,3

xmlfile=server.mappath("../img.xml")
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile(xmlfile,True)

'------------------生成xml

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
myfile.writeline "<autoPlayTime>5</autoPlayTime>"          '图片切换时间，默认值是8，单位秒

'     myfile.writeline "<isHeightQuality>false</isHeightQuality>"  '图片缩小是否采用高质量的方法，默认值false

'     myfile.writeline "<windowOpen>_self</windowOpen>"      '图片连接的打开方式，默认值”_blank”,在新窗口打开，也可以使用”_self”,使用本窗口打开

myfile.writeline "<btnSetMargin>auto 10 10 auto</btnSetMargin>"
'按钮的位置，文字的位置，用了css的margin概念，默认值”auto 5 5 auto”，四个数值代表 上 右 下 左 相对于播放器的距离，四个数值用空格分开，不需具体数值用”auto”填写 ，比如左上对齐并都有10像素的距离可以写 “10 auto auto 10″, 右下角对齐是”auto 10 10 auto”

myfile.writeline "<scaleMode>exactFil</scaleMode>"

'图片放缩模式: 默认值是”noBorder”
'“showAll”: 可以看到全部图片,保持比例,可能上下或者左右
'“exactFil”: 放缩图片到舞台的尺寸,可能比例失调
'“noScale”: 图片的原始尺寸,无放缩
'“noBorder”: 图片充满播放器,保持比例,可能会被裁剪

myfile.writeline "<changImageMode>click</changImageMode>"

'切换图片的方法，默认值”click”,点击切换图片，还可以使用”hover”,鼠标悬停就切换图片


myfile.writeline "<isShowBtn>false</isShowBtn>"

'是否显示按钮，默认值”true”


myfile.writeline "<transform>breatheBlur</transform>"

'图片放缩模式: 默认值是”alpha”
'“alpha”: 透明度淡入淡出
'“blur”: 模糊淡入淡出
'“left”: 左方图片滚动
'“right”: 右方图片滚动
'“top”: 上方图片滚动
'“bottom”: 下方图片滚动
'“breathe”: 有一点点地放缩的淡入淡出
'“breatheBlur”: 有一点点地放缩的模糊淡入淡出


'myfile.writeline "<btnDistance>20</btnDistance>"
'每个按钮的距离，默认值20

'myfile.writeline "<titleBgColor>0xff6600</titleBgColor>"
'标题背景的颜色，默认0xff6600

'myfile.writeline "<titleTextColor>0xffffff</titleTextColor>"

'标题文字的颜色，默认0xffffff

myfile.writeline "<titleBgAlpha>0.75</titleBgAlpha>"

'标题背景的透明度，默认0.75

'myfile.writeline "<titleFont>微软雅黑</titleFont>"

'标题文字的字体，默认值”微软雅黑”

'myfile.writeline "<titleMoveDuration>1</titleMoveDuration>"

'标题背景动画的时间，默认值1，单位秒

'myfile.writeline "<btnAlpha>0.7</btnAlpha>"

'按钮的透明度，默认值0.7


'myfile.writeline "<btnTextColor>0xffffff</btnTextColor>"

'按钮文字的颜色，默认值0xffffff


'myfile.writeline "<btnDefaultColor>0×1B3433</btnDefaultColor>"
'按钮的默认颜色，默认值0×1B3433

'myfile.writeline "<btnHoverColor>0xff9900</btnHoverColor>"

'按钮的默认颜色，默认值0xff9900

'myfile.writeline "<btnFocusColor>0xff6600</btnFocusColor>"

'按钮当前颜色，默认值0xff6600



myfile.writeline "<isShowTitle>false</isShowTitle>"        ' 是否显示标题，默认值”true”

'myfile.writeline "<roundCorner>0</roundCorner>"           '设置圆角

myfile.writeline "<isShowAbout>false</isShowAbout>"    '是否显示关于信息，默认值”true”
myfile.writeline "</config>"

myfile.writeline "</data>"


myfile.close











'从数据库中读数据
  
 	set rsxml=server.createobject("adodb.recordset")
	sql="select * from jiedai_img where tuijian=1 and classid=1 order by id desc"
	rsxml.open sql,conn,1,3

xmlfile=server.mappath("../en/img.xml")
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.CreateTextFile(xmlfile,True)

'------------------生成xml

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
myfile.writeline "<autoPlayTime>5</autoPlayTime>"          '图片切换时间，默认值是8，单位秒

'     myfile.writeline "<isHeightQuality>false</isHeightQuality>"  '图片缩小是否采用高质量的方法，默认值false

'     myfile.writeline "<windowOpen>_self</windowOpen>"      '图片连接的打开方式，默认值"_blank",在新窗口打开，也可以使用"_self",使用本窗口打开

myfile.writeline "<btnSetMargin>auto 10 10 auto</btnSetMargin>"
'按钮的位置，文字的位置，用了css的margin概念，默认值"auto 5 5 auto"，四个数值代表 上 右 下 左 相对于播放器的距离，四个数值用空格分开，不需具体数值用"auto"填写 ，比如左上对齐并都有10像素的距离可以写 "10 auto auto 10″, 右下角对齐是"auto 10 10 auto"

myfile.writeline "<scaleMode>exactFil</scaleMode>"

'图片放缩模式: 默认值是"noBorder"
'"showAll": 可以看到全部图片,保持比例,可能上下或者左右
'"exactFil": 放缩图片到舞台的尺寸,可能比例失调
'"noScale": 图片的原始尺寸,无放缩
'"noBorder": 图片充满播放器,保持比例,可能会被裁剪

myfile.writeline "<changImageMode>click</changImageMode>"

'切换图片的方法，默认值"click",点击切换图片，还可以使用"hover",鼠标悬停就切换图片


myfile.writeline "<isShowBtn>false</isShowBtn>"

'是否显示按钮，默认值"true"


myfile.writeline "<transform>breatheBlur</transform>"

'图片放缩模式: 默认值是"alpha"
'"alpha": 透明度淡入淡出
'"blur": 模糊淡入淡出
'"left": 左方图片滚动
'"right": 右方图片滚动
'"top": 上方图片滚动
'"bottom": 下方图片滚动
'"breathe": 有一点点地放缩的淡入淡出
'"breatheBlur": 有一点点地放缩的模糊淡入淡出


'myfile.writeline "<btnDistance>20</btnDistance>"
'每个按钮的距离，默认值20

'myfile.writeline "<titleBgColor>0xff6600</titleBgColor>"
'标题背景的颜色，默认0xff6600

'myfile.writeline "<titleTextColor>0xffffff</titleTextColor>"

'标题文字的颜色，默认0xffffff

myfile.writeline "<titleBgAlpha>0.75</titleBgAlpha>"

myfile.writeline "<isShowTitle>false</isShowTitle>"        ' 是否显示标题，默认值"true"

myfile.writeline "<isShowAbout>false</isShowAbout>"    '是否显示关于信息，默认值"true"
myfile.writeline "</config>"

myfile.writeline "</data>"


myfile.close





%>
<meta http-equiv="refresh" content="0;URL=jiedai_img.asp?action=list" />

