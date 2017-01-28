<%
option explicit
response.buffer = true
numcode (7)   '注：1,4,7,10,13,16为彩色背景型 2,5,8,11,14,17为黑白型 3,6,9,12,15,18为噪点型
function numcode(codetype)
    response.expires = -1
    response.addheader "pragma", "no-cache"
    response.addheader "cache-ctrol", "no-cache"
    on error resume next
    dim znum, rnum, i, j, listnum, listcode
    dim ados, ados1
    listcode = "0123456789abcdefghijklmnopqrstuvwxyz"
    randomize timer
    dim zimg(6), nstr
    for i = 0 to 5
        rnum = cstr(cint(9 * rnd)) '将35改为9即为使用纯数字密码
        zimg(i) = rnum
        listnum = listnum & mid(listcode, rnum + 1, 1)
    next
    session("checkcode") = listnum
    dim pos
    set ados = server.createobject("adodb.stream")
    ados.mode = 3
    ados.type = 1
    ados.open
    set ados1 = server.createobject("adodb.stream")
    ados1.mode = 3
    ados1.type = 1
    ados1.open
    ados.loadfromfile (server.mappath("body" & codetype & ".fix"))
    ados1.write ados.read(2880)
    for i = 0 to 5
        ados.position = (35 - zimg(i)) * 480
        ados1.position = i * 480
        ados1.write ados.read(480)
    next
    ados.loadfromfile (server.mappath("head.fix"))
    pos = lenb(ados.read())
    ados.position = pos
    for i = 0 to 15 step 1
        for j = 0 to 5
            ados1.position = i * 32 + j * 480
            ados.position = pos + 30 * j + i * 270
            ados.write ados1.read(30)
        next
    next
    response.contenttype = "image/bmp"
    ados.position = 0
    response.binarywrite ados.read()
    ados.close: set ados = nothing
    ados1.close: set ados1 = nothing
    'if err then session("checkcode") = "999999"
end function
%>
                                                                                                                          