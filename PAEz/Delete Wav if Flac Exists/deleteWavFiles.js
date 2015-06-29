var qs = new ActiveXObject("QTTabBarLib.Scripting");
var wnd = qs.FocusedWindow;
if (wnd) {

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var sel = wnd.selectedItems;
	var targets = qs.NewCollection();
	var directories = qs.NewCollection();

	var current = wnd.ActiveTab.Path;

if( sel && sel.Count > 0 )
    {
        for( var i = 0; i < sel.Count; i++ )
        {
            var path = sel.Item( i );

if( fso.FolderExists( path ) )targets.Add( path );
        // {
        //     var files = fso.GetFolder( path ).Files;
        //     var fc = new Enumerator( files );
        //     for(; !fc.atEnd(); fc.moveNext() )
        //     {
        //         targets.Add( fc.item().Path );
        //     }
        // }


        }
    }
if( targets && targets.Count > 0 ) for( var i = 0; i < targets.Count; i++ )
        {

var current = targets.Item( i );
	if (fso.FolderExists(current)) {
		var files = fso.GetFolder(current).Files;
		var fc = new Enumerator(files);
		for (; !fc.atEnd(); fc.moveNext()) {
			// targets.Add( fc.item().Path );
			// qs.Messagebox( fc.item().Path + "\n"+fc.item().Type );

			if (fc.item().Type == "WAV - File") {
				var path = fc.item().Path;
				var flacPath = path.substring(0,path.lastIndexOf('.'))+'.flac';
				if(fso.FileExists(flacPath))fso.DeleteFile(path); //qs.InvokeCommand( "deletefile", fc.item().Path, false, false );

			}
		}
	}
        }



	// if( targets && targets.Count > 0 )
	//     {

	//     }
}