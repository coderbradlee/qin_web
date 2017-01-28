<%@LANGUAGE="VBSCRIPT" CODEPAGE="936" %> 
<%Session.CodePage=936%>
<%Response.Addheader "Content-Type","text/html; charset=gb2312"%>
<%
'on error resume next
dim provider,path,pass,dsn,conn
provider="provider=microsoft.jet.oledb.4.0;"
path="data source=" & server.mappath("jdshuju/#jiedai.mdb")
pass=";jet oledb:database password="
dsn=provider&path&pass
set conn=server.createobject("adodb.connection")
conn.open dsn






sql="select * from web_config where id=1"
set rgs=conn.execute(sql)

webicp=rgs("webicp")
webimage=rgs("image")
if webimage<>"" then
if left(webimage,7)<>"http://" then  webimage="/uploadfile/"&webimage
end if

title=rgs("webname")
keywords_content=rgs("webkeyword")
description_content=rgs("webdes")

rgs.close
set rgs=nothing
function checkStr(str)
str=replace(str,"'","")
Str=Replace(Str,chr(39),"") 'SQL?
Str=Replace(Str,chr(91),"") 'SQL?[
Str=Replace(Str,chr(93),"") 'SQL?]
Str=Replace(Str,chr(37),"") 'SQL?%
Str=Replace(Str,chr(59),"") 'SQL?
Str=Replace(Str,chr(43),"") 'SQL?;
Str=Replace(Str,chr(45),"") 'SQL?+
Str=Replace(Str,chr(123),"") 'SQL?{
Str=Replace(Str,chr(125),"") 'SQL?}

checkStr=Str '??滻Str
if isnull(str) then
checkStr = ""
exit function 
end if
end function

jd_username=session("jd_username")
jd_userid=session("jd_userid")

%>



<%
'**************************************************
'got
'  ???????
'  str   ----??
'       strlen ----?
'???
'**************************************************
Function got(ByVal str, ByVal strlen)
    If str = "" Then
        got = ""
        Exit Function
    End If
    Dim l, t, c, i, strTemp
    str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    l = Len(str)
    t = 0
    strTemp = str
    strlen = CLng(strlen)
    For i = 1 To l
        c = Abs(Asc(Mid(str, i, 1)))
        If c > 255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t >= strlen Then
            strTemp = Left(str, i)
            Exit For
        End If
    Next
    If strTemp <> str Then
        strTemp = strTemp & "..."
    End If
    got = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
End Function
%>

<% 
Function RemoveHTML(strText) 
Dim RegEx 
Set RegEx = New RegExp 
RegEx.Pattern = "<[^>]*>" 
RegEx.Global = True 
RemoveHTML = RegEx.Replace(strText, "") 
End Function
%>    




<% 
Function imgg(imgurl)
if imgurl="" or imgurl="../uploadfile/" then
imgg="images/nopic.jpg"
else
imgg=imgurl
end if
End Function

 %>






 <%
urrl=request.servervariables("http_url") '???
uu=replace(urrl,"/","")
str=Request.ServerVariables("Query_String")'??
'str=LCase(str)
urll=replace(uu,"?"&str&"","") '???
url=LCase(replace(urll,".asp","")) '?".ASP"? ?С
'Response.Write url
%>


<%



Function LoseStyleTag(ContentStr)  '过滤 style 标记
Dim ClsTempLoseStr,regEx
if ContentStr<>"" then
ClsTempLoseStr = Cstr(ContentStr)
Set regEx = New RegExp
regEx.Pattern = "(<style)+[^<>]*>[^\0]*(<\/style>)+"
regEx.IgnoreCase = True
regEx.Global = True
ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
LoseStyleTag = ClsTempLoseStr
Set regEx = Nothing
end if
End Function



Function LoseFontTag(ContentStr)  '过滤 FONT 标记
Dim ClsTempLoseStr,regEx
ClsTempLoseStr = Cstr(ContentStr)
Set regEx = New RegExp
regEx.Pattern = "<(\/){0,1}font[^<>]*>"
regEx.IgnoreCase = True
regEx.Global = True
ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
LoseFontTag = ClsTempLoseStr
Set regEx = Nothing
End Function


Function LoseIFrameTag(ContentStr)  '过滤 iframe 标记
Dim ClsTempLoseStr,regEx
ClsTempLoseStr = Cstr(ContentStr)
Set regEx = New RegExp
regEx.Pattern = "(<iframe){1,}[^<>]*>[^\0]*(<\/iframe>){1,}"
regEx.IgnoreCase = True
regEx.Global = True
ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
LoseIFrameTag = ClsTempLoseStr
Set regEx = Nothing
End Function




'================================================
   '函数名：FormatDate
   '作 用：格式化日期
   '参 数：DateAndTime   ----原日期和时间
   '        para   ----日期格式
   '返回值：格式化后的日期
   '================================================
  
   Public Function FormatDate(DateAndTime, para)
  
     On Error Resume Next
     Dim y, m, d, h, mi, s, strDateTime
     FormatDate = DateAndTime
     If Not IsNumeric(para) Then Exit Function
     If Not IsDate(DateAndTime) Then Exit Function
     y = CStr(Year(DateAndTime))
     m = CStr(Month(DateAndTime))
     If Len(m) = 1 Then m = "0" & m
     d = CStr(Day(DateAndTime))
     If Len(d) = 1 Then d = "0" & d
     h = CStr(Hour(DateAndTime))
     If Len(h) = 1 Then h = "0" & h
     mi = CStr(Minute(DateAndTime))
     If Len(mi) = 1 Then mi = "0" & mi
     s = CStr(Second(DateAndTime))
     If Len(s) = 1 Then s = "0" & s
     
     Select Case para
  
     Case "1"
    '显示格式：09年07月06日 13:44 
    strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  
     Case "2"
    '显示格式：2009-07-06
    strDateTime = y & "-" & m & "-" & d
  
     Case "3"
    '显示格式：2009/07/06
    strDateTime = y & "/" & m & "/" & d
  
     Case "4"
    '显示格式：2009年07月06日
    strDateTime = y & "." & m & "." & d
  
     Case "5"
    '显示格式：07-06 13:45
    strDateTime = m & "-" & d & " " & h & ":" & mi
  
     Case "6"
    '显示格式：07/06
    strDateTime = m & "/" & d
  
     Case "7"
    '显示格式：07月06日
    strDateTime = m & "月" & d & "日"
  
     Case "8"
    '显示格式：2009年07月
    strDateTime = y & "年" & m & "月"
  
     Case "9"
    '显示格式：2009-07
    strDateTime = y & "-" & m
  
     Case "10"
    '显示格式：2009/07
    strDateTime = y & "/" & m
  
     Case "11"
      '显示格式：09年07月06日 13:45
    strDateTime = right(y,2) & "年" &m & "月" & d & "日 " & h & ":" & mi
  
     Case "12"
    '显示格式：09-07-06
    strDateTime = right(y,2) & "-" &m & "-" & d
  
     Case "13"
    '显示格式：07-06
    strDateTime = m & "-" & d
   
     Case "14"
    '显示格式：13:45
    strDateTime = h & ":" & mi
  
     Case Else
  
    strDateTime = DateAndTime
  
     End Select
  
   FormatDate = strDateTime
  
   End Function








function myRandn(n) '生成随机数字，n为数字的个数
  dim thechr
  thechr = ""
  for i=1 to n
    dim zNum,zNum2
    Randomize
    zNum = cint(9*Rnd)
    zNum = zNum + 48 '这里换成77可以生成字母
    thechr = thechr & chr(zNum)
  next
    MyRandn = thechr
End Function

' 生成订单号
dingdan = Year(Now())&Month(Now())&Day(Now())&Hour(Now())&minute(Now())&second(Now())&myRandn(5) 



'生成的是一个不重复的数组
Function GetRnd(lowerNum,upperNum)
    Dim unit,RndNum,Fun_X
    unit = upperNum - lowerNum
    Redim MyArray(unit)
    For Fun_I=0 To unit
        myArray(Fun_I)= lowerNum + Fun_I
    Next
    For Fun_I=0 To round(unit)
        RndNum = getRndNumber(Fun_I,unit)
        Fun_X = myArray(RndNum)
        myArray(RndNum)=myArray(Fun_I)
        myArray(Fun_I)=Fun_X
    Next
    GetRnd = Join(myArray)
End Function

Function getRndNumber(lowerbound,upperbound)
     Randomize
     getRndNumber=Int((upperbound-lowerbound+1)*Rnd+lowerbound)
End Function 
'Response.Write GetRnd(1,1000)






Function ClearHtml(Content) 
    Content=Zxj_ReplaceHtml("&#[^>]*;", "", Content) 
    Content=Zxj_ReplaceHtml("</?marquee[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?object[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?param[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?embed[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?table[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml(" ","",Content) 
    Content=Zxj_ReplaceHtml("</?tr[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?th[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?p[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?a[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?img[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?tbody[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?li[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?span[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?div[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?th[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?td[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?script[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("(javascript|jscript|vbscript|vbs):", "", Content) 
    Content=Zxj_ReplaceHtml("on(mouse|exit|error|click|key)", "", Content) 
    Content=Zxj_ReplaceHtml("<\\?xml[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("<\/?[a-z]+:[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?font[^>]*>", "", Content) 
    Content=Zxj_ReplaceHtml("</?b[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?u[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?i[^>]*>","",Content) 
    Content=Zxj_ReplaceHtml("</?strong[^>]*>","",Content) 	
	 Content=Zxj_ReplaceHtml("&nbsp;","",Content) 
    ClearHtml=Content 
   End Function


Function Zxj_ReplaceHtml(patrn, strng,content) 
    IF IsNull(content) Then 
    content="" 
    End IF 
    Set regEx = New RegExp ' 建立正则表达式。 
    regEx.Pattern = patrn ' 设置模式。 
    regEx.IgnoreCase = true ' 设置忽略字符大小写。 
    regEx.Global = True ' 设置全局可用性。 
    Zxj_ReplaceHtml=regEx.Replace(content,strng) ' 执行正则匹配 
   End Function 







%>









                                                                                                                          