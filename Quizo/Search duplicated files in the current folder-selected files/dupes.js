var qs = new ActiveXObject( "QTTabBarLib.Scripting" );
var wnd = qs.FocusedWindow;
if( wnd )
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var sel = wnd.selectedItems;
    var targets = qs.NewCollection();
    if( sel && sel.Count > 0 )
    {
        for( var i = 0; i < sel.Count; i++ )
        {
            var path = sel.Item( i );
            if( fso.FileExists( path ) )
            {
                targets.Add( path );
            }
        }
    }
    else
    {
        var current = wnd.ActiveTab.Path;
        if( fso.FolderExists( current ) )
        {
            var files = fso.GetFolder( current ).Files;
            var fc = new Enumerator( files );
            for(; !fc.atEnd(); fc.moveNext() )
            {
                targets.Add( fc.item().Path );
            }
        }
    }
 
    if( targets && targets.Count > 0 )
    {
        var dic = new Object();
        for( var i = 0; i < targets.Count; i++ )
        {
            var path = targets.Item( i );
            var hash = wnd.InvokeCommand( "ComputeFileHash", path );
            if( hash )
            {
                if( !dic[hash] )
                {
                    dic[hash] = qs.NewCollection();
                }
                dic[hash].Add( path );
            }
        }
 
        var str = "";
        for( var s in dic )
        {
            var col = dic[s];
            if( col.Count > 1 )
            {
                str += "These are identical:\n";
                for( var i = 0; i < col.Count; i++ )
                {
                    str += col.Item( i ) + "\n";
                }
                str += "\n";
            }
        }
        if( str )
        {
            if( 6 == qs.Messagebox( str + "\nDo you want to send to Clipboard?", "YesNo" ) )
            {
                qs.InvokeCommand( "SetClipboard", str );
            }
        }
        else
        {
            qs.alert( "No duplication." );
        }
    }
    else
    {
        qs.alert( "No file exists." );
    }
}