if ( typeof(Array.prototype.indexOf) === 'undefined' ) {
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
var valid = {
	SYSTEM: true,
	USER: true,
	PROCESS: true,
	VOLATILE: true
}
var option = '';
if (args.Count()) {
	namespace = args.item(0);
	if (args.Count() >= 2) path = args.item(1);
} else {
	namespace = 'USER';
}

if (!valid[namespace]) WScript.Quit();

if (path === undefined) {
	var sel = wnd.selectedItems;
	if (sel.Count) {
		path = sel.item(0);
	} else WScript.Quit();
}

var env = shell.Environment(namespace);
var systemEnv =shell.Environment('SYSTEM');
var systemPaths=systemEnv('PATH').toUpperCase().split(';');
var systemExists=systemPaths.indexOf(path.toUpperCase())!=-1;

var paths = env('PATH').toUpperCase().split(';');
var exists = paths.indexOf(path.toUpperCase())!=-1;


if (exists) {
	maybeDelete(path);
} else {
	if(systemExists){
		qs.Messagebox(path + "\nWas found in SYSTEM\rBut due to permissions it cant be deleted.");
		WScript.Quit();
	}
	maybeAdd(path);
}

function maybeAdd(path) {
	if (6 == qs.Messagebox(path + "\nWas not found in " + namespace  + '\rWould you like to Add it?', "YesNo")) {
		paths.push(path);
		env("PATH") = paths.join(';');
	}
}

function maybeDelete(path) {
	if (6 == qs.Messagebox(path + "\nWas found in " + namespace  + '\rWould you like to Delete it?', "YesNo")) {
		paths.splice(paths.indexOf(path), 1);
		env("PATH") = paths.join(';');
	}
}
