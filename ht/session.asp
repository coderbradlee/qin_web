<% 
session.timeout=30
if session("admin")=empty and session("password")=empty then
	response.write "<script>location.href='exit.asp'</script>"
end if
 %>
                                                                                                                          