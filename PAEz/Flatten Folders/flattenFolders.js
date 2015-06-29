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
			if (fso.FolderExists(path)) targets.Add(path);
		}
	}
	if (targets && targets.Count > 0) {
		for (var i = 0; i < targets.Count; i++) {
			var current = targets.Item(i);
			if (fso.FolderExists(current)) {
				var files = fso.GetFolder(current).Files;
				var fc = new Enumerator(files);
				for (; !fc.atEnd(); fc.moveNext()) {
					var path = fc.item().Path;
					// var name = fso.GetBaseName(fc.item().Name);
					// var ext = fso.GetExtensionName(fc.item().Path);
					path = fso.GetParentFolderName(path);
					fso.MoveFile(fc.item().Path, findUniqueName(fc.item(), currentDir));
				}

				var dirs = fso.GetFolder(current).SubFolders;

				fc = new Enumerator(dirs);
				for (; !fc.atEnd(); fc.moveNext()) {
					fso.MoveFolder(fc.item().Path, findUniqueDir(fc.item(), currentDir))
				}

				var dir = fso.GetFolder(current);
				if(dir.Files.Count===0 && dir.SubFolders.Count===0)fso.DeleteFolder(dir);
			}
		}
	}
}

function findUniqueName(file, destPath) {
	var name = fso.GetBaseName(fc.item().Name);
	var ext = fso.GetExtensionName(fc.item().Path);
	var id = 1;
	var newPath = destPath + '\\' + name + ' (' + id + ').' + ext;
	if (!fso.FileExists(destPath + '\\' + name + '.' + ext)) return destPath + '\\' + name + '.' + ext;
	while (fso.FileExists(newPath)) {
		id++;
		newPath = destPath + '\\' + name + ' (' + id + ').' + ext;
	}
	return newPath;
}

function findUniqueDir(dir, destPath) {
	var name = fc.item().Name;
	var id = 1;
	var newPath = destPath + '\\' + name + ' (' + id + ')'
	if (!fso.FolderExists(destPath + '\\' + name)) return destPath + '\\' + name;
	while (fso.FolderExists(newPath)) {
		id++;
		newPath = destPath + '\\' + name + ' (' + id + ')';
	}
	return newPath;
}