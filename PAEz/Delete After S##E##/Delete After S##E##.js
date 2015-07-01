var qs = new ActiveXObject("QTTabBarLib.Scripting");

var wnd = qs.FocusedWindow;

if (wnd && option) {
	var sel = wnd.selectedItems;
	if (sel.Count) {
		var fso = new ActiveXObject("Scripting.FileSystemObject");
		for (var i = 0; i < sel.Count; i++) {
			var path = sel.Item(i);
			var _name = fso.GetFileName(path);
			var name = _name.substring(0,_name.lastIndexOf('.'));
			var ext = _name.substring(_name.lastIndexOf('.'));
			var newName = name.replace(/([sS]\d+[eE]\d+)(.*)/,function (a,b){
				return b;
			});
			newName+=ext;
			if (name != newName) {
				qs.InvokeCommand("Rename", path, newName);
			}

		}
	}
}
