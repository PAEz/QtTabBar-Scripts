if (typeof(Array.prototype.indexOf) === 'undefined') {
	Array.prototype.indexOf = function(item, start) {
		var length = this.length;
		start = typeof(start) !== 'undefined' ? start : 0;
		for (var i = start; i < length; i++) {
			if (this[i] === item) return i
		}
		return -1
	}
}
var qs = new ActiveXObject("QTTabBarLib.Scripting");
var shell = WScript.CreateObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var wnd = qs.ActiveWindow;

var args = WScript.Arguments,
	namespace, path;
var path = args[0];
if (path === undefined) {
	var sel = wnd.selectedItems;
	if (sel.Count) {
		path = sel.item(0);
	} else WScript.Quit();
}
OpenFileLocation(path);

function OpenFileLocation(path) {
	// shell.Popup (path);
	var allowedExt = {
		'.lnk': 1,
		'.url': 1
	};
	if (!allowedExt[path.substring(path.lastIndexOf('.'))]) return;
	var FSO, oShortcut, sFileSpec, sTarget

	// 'Get the shortcut file's location
	sFileSpec = path;
	// 'Instantiate a Shortcut Object
	oShortcut = shell.CreateShortcut(sFileSpec);
	// 'Retrieve the shortcut's target
	sTarget = oShortcut.TargetPath;

	goTo(sTarget);
}

function goTo(target) {
	var path, file;
	if (fso.FileExists(target)) {
		var file = fso.GetFile(target);
		path = file.ParentFolder;
		file = file.Name;
		var tab = wnd.ActiveTab;
		tab.Path = path;
		WScript.Sleep(100);
		var newSel = qs.NewCollection();
		newSel.Push(file);
		tab.SelectedItems = newSel;
	} else if (fso.FolderExists(target)) {
		var tab = wnd.ActiveTab;
		tab.NavigateTo(target);
	} else {
		// 'complain, er, inform if it's missing
		shell.Popup("Could not find: " + target);
	}
}