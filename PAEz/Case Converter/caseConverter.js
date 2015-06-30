var qs = new ActiveXObject("QTTabBarLib.Scripting");

var wnd = qs.FocusedWindow;

var args = WScript.Arguments;
var option = args(0);

if( wnd && option)
{
    var sel = wnd.selectedItems;
    if( sel.Count )
    {
        var fso = new ActiveXObject( "Scripting.FileSystemObject" );
        for( var i = 0; i < sel.Count; i++ )
        {
            var path = sel.Item( i );
            var name = fso.GetFileName( path );
 
            if( option == 'upperFirst' )
            {
            	name=name.substring(0,name.lastIndexOf('.')).replace(/[.]/g,' ')+name.substring(name.lastIndexOf('.'));
                qs.InvokeCommand( "Rename", path, UpperAtFirstChar( name ) );
            }
            else if( option == 'lower' )
            {
                qs.InvokeCommand( "Rename", path, name.toLowerCase() );
            }
            else if( option == 'upper' )
            {
                qs.InvokeCommand( "Rename", path, name.substring(0,name.lastIndexOf('.')).toUpperCase()+name.substring(name.lastIndexOf('.')) );
            }
        }
    }
}
 
 
function UpperAtFirstChar( str )
{
    var arr = str.split( " " );
    for( var i = 0; i < arr.length; i++ )
    {
        if( arr[i] )
        {
            if( arr[i].length > 1 )
            {
                arr[i] = arr[i].charAt( 0 ).toUpperCase() + arr[i].substring( 1 );
            }
            else
            {
                arr[i] = arr[i].charAt( 0 ).toUpperCase();
            }
        }
    }
    return arr.join( " " );
}