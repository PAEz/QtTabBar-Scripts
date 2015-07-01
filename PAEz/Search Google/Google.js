var qs = new ActiveXObject("QTTabBarLib.Scripting");
var shell = WScript.CreateObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var wnd = qs.ActiveWindow;

var sel = wnd.selectedItems;
if (sel.Count) {
	for (var i = 0; i < sel.Count; i++) {
		var what = sel.Item(i);
		var what = fso.GetFileName(sel.item(0));
		what = encodeURI(what);
		shell.Run('http://google.com/search?q=' + what, 0);
	}
}