dim window, i

set window = createwindow()
window.document.write "<html><body bgcolor=buttonface>Current loop is: <span id='output'></span> </body> </html>"
window.document.title = "Processing..."
window.resizeto 300, 150
window.moveto 200, 200

do
for i = 0 to 120000
'for i = 0 to 30000
    show i
	
Set ws = CreateObject("wscript.shell")
	WScript.Sleep(30000)
	ws.sendkeys "{SCROLLLOCK}"

next
loop
window.close

function show(value)
    on error resume next
    window.output.innerhtml = value
    if err then wscript.quit
end function

function createwindow()

    dim signature, shellwnd, proc
    on error resume next
    set createwindow = nothing
    signature = left(createobject("Scriptlet.TypeLib").guid, 38)
    set proc = createobject("WScript.Shell").exec("mshta about:""<script>moveTo(-32000,-32000);</script><hta:application id=app border=dialog minimizebutton=no maximizebutton=no scroll=no showintaskbar=yes contextmenu=no selection=no innerborder=no /><object id='shellwindow' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shellwindow.putproperty('" & signature & "',document.parentWindow);</script>""")
    do
        if proc.status > 0 then exit function
        for each shellwnd in createobject("Shell.Application").windows
            set createwindow = shellwnd.getproperty(signature)
            if err.number = 0 then exit function
            err.clear
        next
    loop
end function
