var qs = new ActiveXObject( "QTTabBarLib.Scripting" );
var wnd = qs.FocusedWindow;
if( wnd )
{
    var sel = wnd.selectedItems;
    if( sel.Count )
    {
        var fso = new ActiveXObject( "Scripting.FileSystemObject" );
        for( var i = 0; i < sel.Count; i++ )
        {
            var path = sel.Item( i );
            var name = fso.GetFileName( path );
 
            if( UpperAtFirstChar( name ) != name )
            {
                qs.InvokeCommand( "Rename", path, UpperAtFirstChar( name ) );
            }
            else if( UpperCased( name ) )
            {
                qs.InvokeCommand( "Rename", path, name.toLowerCase() );
            }
            else
            {
                qs.InvokeCommand( "Rename", path, name.toUpperCase() );
            }
        }
    }
}
 
function UpperCased( str )
{
    for( var i = 0; i < str.length; i++ )
    {
        var ch = str.charAt( i );
        if( ( "a" <= ch && ch <= "z" ) || ( "A" <= ch && ch <= "Z" ) )
        {
            if( ch.toUpperCase() != ch )
            {
                return false;
            }
        }
    }
    return true;
}
 
function LowerCased( str )
{
    for( var i = 0; i < str.length; i++ )
    {
        var ch = str.charAt( i );
        if( ( "a" <= ch && ch <= "z" ) || ( "A" <= ch && ch <= "Z" ) )
        {
            if( ch.toLowerCase() != ch )
            {
                return false;
            }
        }
    }
    return true;
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