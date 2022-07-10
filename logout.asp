<% DIM content_type: content_type=Request.ServerVariables("HTTP_ACCEPT")
Session.Abandon
IF (INSTR(UCASE(content_type),"JSON")>0) THEN
    Response.ContentType = "application/json"
    Response.CharSet = "ISO-8859-1"
%>{
	"success": true
}<%
ELSE
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=10" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title></title>
    <script type="text/javascript">
        // transfers sessionStorage from one tab to another
        var sessionStorage_transfer = function (event) {
            if (!event) { event = window.event; } // ie suq
            if (!event.newValue) return;          // do nothing if no value to work with
            if (event.key == 'getSessionStorage') {
                // another tab asked for the sessionStorage -> send it
                localStorage.setItem('sessionStorage', JSON.stringify(sessionStorage));
                // the other tab should now have it, so we're done with it.
                localStorage.removeItem('sessionStorage'); // <- could do short timeout as well.
            } else if (event.key == 'sessionStorage' && !sessionStorage.length) {
                // another tab sent data <- get it
                var data = JSON.parse(event.newValue);
                for (var key in data) {
                    sessionStorage.setItem(key, data[key]);
                }
            }
        };

        // listen for changes to localStorage
        if (window.addEventListener) {
            window.addEventListener("storage", sessionStorage_transfer, false);
        } else {
            window.attachEvent("onstorage", sessionStorage_transfer);
        };

        // Ask other tabs for session storage (this is ONLY to trigger event)
        if (!sessionStorage.length) {
            localStorage.setItem('getSessionStorage', 'foobar');
            localStorage.removeItem('getSessionStorage', 'foobar');
        };
    </script>
</head>
<body>
    <div>
        <label>Hasta luego!</label>
    </div>
    <script type="text/javascript">
        if (typeof (sessionStorage) !== "undefined") {
            sessionStorage.setItem("userId",undefined);
        }
    </script>
</body>
</html>
<% END IF %>