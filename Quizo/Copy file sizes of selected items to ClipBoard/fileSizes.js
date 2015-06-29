var qs = new ActiveXObject( "QTTabBarLib.Scripting" );
var wnd = qs.ActiveWindow;
var tab = wnd.ActiveTab;
var sel = tab.selectedItems;
if( sel.Count )
{
    var single = sel.Count == 1;
    var str = "";
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    for( var i = 0; i < sel.Count; i++ )
    {
        var path = sel.Item( i );
        if( fso.FileExists( path ) )
        {
            if( single )
            {
                str += fso.GetFile( path ).Size;
            }
            else
            {
                str += path + ", "+ fso.GetFile( path ).Size + "\r\n";
            }
        } 
        else if( fso.FolderExists( path ) )
        {
            if( single )
            {
                str += fso.GetFolder( path ).Size;
            }
            else
            {
                str += path + ", "+ fso.GetFolder( path ).Size + "\r\n";
            }
        }
    }
    wnd.InvokeCommand( "SetClipboard", str );
}