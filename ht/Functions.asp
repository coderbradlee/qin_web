<%
'**************************************************
'函数名：replacebadchar
'作  用：过滤非法的sql字符
'参  数：strchar-----要过滤的字符
'返回值：过滤后的字符
'**************************************************
function replacebadchar(strchar)
	if strchar="" then
		replacebadchar=""
	else
		replacebadchar=replace(replace(replace(replace(replace(replace(replace(strchar,"'",""),"*",""),"?",""),"(",""),")",""),"<",""),".","")
        replacebadchar=replace(replacebadchar," ","")
        replacebadchar=replace(replacebadchar,";","")
		
		replacebadchar=lcase(replacebadchar)
        replacebadchar=replace(replacebadchar,"or","")
        replacebadchar=replace(replacebadchar,"and","")
        replacebadchar=replace(replacebadchar,"not","")
		
        replacebadchar=replace(replacebadchar,"select","")
        replacebadchar=replace(replacebadchar,"drop","")
        replacebadchar=replace(replacebadchar,"delete","")
        replacebadchar=replace(replacebadchar,"update","")
        replacebadchar=replace(replacebadchar,"insert","")
		
        replacebadchar=replace(replacebadchar,"count","")
        replacebadchar=replace(replacebadchar,"exec","")
        replacebadchar=replace(replacebadchar,"truncate","")
        replacebadchar=replace(replacebadchar,"net","")
		
        replacebadchar=replace(replacebadchar,"asc","")
        replacebadchar=replace(replacebadchar,"char","")
        replacebadchar=replace(replacebadchar,"mid","")
	end if
end function
%>
                   

  <%

function Replace_Text(fString)
if isnull(fString) then
Replace_Text=""
exit function
else
fString=trim(fString)
fString=replace(fString,"'","''")
fString=replace(fString,";","；")
fString=replace(fString,"--","—")
fString=server.htmlencode(fString)
Replace_Text=fString
end if	
end function



%>                                                                                                                   