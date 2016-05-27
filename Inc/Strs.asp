<%
dim ShowSwfStr,ShowMoveSwf,ShowImgStr,ShowMoveImg,ShowMoveDiv
ShowSwfStr="<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width='°ÁWTH' height='°ÁHGH'><param name='movie' value='°ÁURL'><param name='quality' value='high'><param name='menu' value='false'><embed src='°ÁURL' width='°ÁWTH' height='°ÁHGH' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' menu='false'></embed></object>"
ShowImgStr="<a href='°ÁGOURL' target='_blank' title='°ÁTEST'><img src='°ÁURL' class='°ÁCLASS'></a>"

ShowMoveSwf="<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width='°ÁWTH' height='°ÁHGH'><param name='movie' value='°ÁURL'><param name='quality' value='high'><param name='menu' value='false'><param name='wmode' value='transparent'><embed src='°ÁURL' width='°ÁWTH' height='°ÁHGH' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' menu='false' wmode='transparent'></embed></object>"
ShowMoveImg="<a href='°ÁGOURL' target='_blank' title='°ÁTEST'><img src='°ÁURL' onMouseOver='JavaScript:pause_resume()' onMouseOut='JavaScript:pause_resume()' border='0' width='°ÁWTH' height='°ÁHGH'></a>"

ShowMoveDiv="<div id='img' style='position:absolute;z-index=99;'>°ÁDIVBODY</div>" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "<SCRIPT LANGUAGE='JavaScript'>" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "<!-- Begin" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var xPos = document.body.clientWidth-20;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var yPos = document.body.clientHeight/2;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var step = 1;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var delay = 5;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var height = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var Hoffset = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var Woffset = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var yon = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var xon = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var pause = true;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "var interval;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "img.style.top = yPos;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "function changePos() {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "width = document.body.clientWidth;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "height = document.body.clientHeight;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "Hoffset = img.offsetHeight;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "Woffset = img.offsetWidth;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "img.style.left = xPos + document.body.scrollLeft;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "img.style.top = yPos + document.body.scrollTop;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "if (yon) {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "yPos = yPos + step;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}else {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "yPos = yPos - step;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "if (yPos < 0) {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "yon = 1;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "yPos = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "if (yPos >= (height - Hoffset)) {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "yon = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "yPos = (height - Hoffset);" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "if (xon) {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "xPos = xPos + step;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}else {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "xPos = xPos - step;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "if (xPos < 0) {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "xon = 1;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "xPos = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "if (xPos >= (width - Woffset)) {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "xon = 0;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "xPos = (width - Woffset);" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf

ShowMoveDiv=ShowMoveDiv & "function start() {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "img.visibility = 'visible';" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "interval = setInterval('changePos()', delay);" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "function pause_resume() {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "if(pause) {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "clearInterval(interval);" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "pause = false;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}else {" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "interval = setInterval('changePos()',delay);" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "pause = true;" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "}" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "start();" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "//  End -->" & vbcrlf
ShowMoveDiv=ShowMoveDiv & "</script>" & vbcrlf
%>