var qs = new ActiveXObject("QTTabBarLib.Scripting");
var shell = WScript.CreateObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var wnd = qs.ActiveWindow;

var args = WScript.Arguments;
var option = '';
if (args.Count()){
	for (var i=0, iEnd=args.Count();i<iEnd;i++){
		option+='%20'+args.item(i);
	}
}

var sel = wnd.selectedItems;
if (sel.Count) {
	for (var i = 0; i < sel.Count; i++) {
		var what = sel.Item(i);
		var what = fso.GetFileName(sel.item(0));
		what = encodeURI(what);
		shell.Run('http://google.com/search?q=' + what+option, 0);

	}
}