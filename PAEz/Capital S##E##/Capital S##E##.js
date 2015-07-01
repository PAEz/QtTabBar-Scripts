var qs = new ActiveXObject("QTTabBarLib.Scripting");

var wnd = qs.FocusedWindow;

var args = WScript.Arguments;
var option = args(0);

if (wnd && option) {
	var sel = wnd.selectedItems;
	if (sel.Count) {
		var fso = new ActiveXObject("Scripting.FileSystemObject");
		for (var i = 0; i < sel.Count; i++) {
			var path = sel.Item(i);
			var name = fso.GetFileName(path);
			var newName = name.replace(/[sS]\d+[eE]\d+/, function(what) {
				return what.toUpperCase()
			});
			if (name != newName) {
				qs.InvokeCommand("Rename", path, newName);
			}

		}
	}
}
