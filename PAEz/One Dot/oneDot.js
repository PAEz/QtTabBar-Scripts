var qs = new ActiveXObject("QTTabBarLib.Scripting");
var wnd = qs.FocusedWindow;
if (wnd) {

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var sel = wnd.selectedItems;
	var targets = qs.NewCollection();
	var currentDir = wnd.ActiveTab.Path;

	if (sel && sel.Count > 0) {
		for (var i = 0; i < sel.Count; i++) {
			var path = sel.Item(i);
			if (fso.FileExists(path)) targets.Add(path);
		}
	}
	if (targets && targets.Count > 0) {
		for (var i = 0; i < targets.Count; i++) {
			var current = targets.Item(i);
			if (fso.FileExists(current)) {
				current = fso.GetFile(current);
				var name = current.Name;
				name=name.substring(0,name.lastIndexOf('.')).replace(/[.]/g,' ')+name.substring(name.lastIndexOf('.'));
				current.Name=name;
			}
		}
	}
}
