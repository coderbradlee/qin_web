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
Str=Replace(Str,chr(39),"") 'SQLע
Str=Replace(Str,chr(91),"") 'SQLע[
Str=Replace(Str,chr(93),"") 'SQLע]
Str=Replace(Str,chr(37),"") 'SQLע%
Str=Replace(Str,chr(59),"") 'SQLע
Str=Replace(Str,chr(43),"") 'SQLע;
Str=Replace(Str,chr(45),"") 'SQLע+
Str=Replace(Str,chr(123),"") 'SQLע{
Str=Replace(Str,chr(125),"") 'SQLע}

checkStr=Str 'ؾַ滻Str
if isnull(str) then
checkStr = ""
exit function 
end if
end function

webed="../qi500@lm_webe/qi500@edit.htm?id=content&style=blue"
webeda="../qi500@lm_webe/qi500@edit.htm?id=nr&style=blue"

webeden="../qi500@lm_webe/qi500@edit.htm?id=e_content&style=blue1"
%>
                                                                                                                          