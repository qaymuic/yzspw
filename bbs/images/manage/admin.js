
// ÐÞ¸Ä±à¼­À¸¸ß¶È
function admin_Size(num,objname)
{
	var obj=document.getElementById(objname)
	if (parseInt(obj.rows)+num>=3) {
		obj.rows = parseInt(obj.rows) + num;	
	}
	if (num>0)
	{
		obj.width="90%";
	}
}

function helpscript(n){
	txtRun=n;window.open('admin_helpview.asp','admin_help','toolbar=no,menubar=no,scrollbars=no, resizable=1, location=no, status=no,top=0,left=0,width=600,height=300')

}

function runscript(n){
	txtRun=n;window.open("templates_view.asp","templates_view")
}
function rundvscript(n,astr){
	txtRun=n;window.open("http://bbs.dvbbs.net/loadtemplates.asp?"+astr+"","loadtemplates")
}