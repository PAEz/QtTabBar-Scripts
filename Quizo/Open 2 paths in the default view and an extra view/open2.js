var pathDefault = "C:\\";
var pathExtra = "C:\\Windows";
var extraView = "Left";    // this can be "Bottom"
 
var qs = new ActiveXObject( "QTTabBarLib.Scripting" );
var wnd = qs.activewindow;
if( wnd )
{
    var tab1 = wnd.invokeCommand( "NewTab", pathDefault );
    wnd.invokeCommand( "CloseAllButOne", tab1 );
 
    var tab2 = wnd.invokeCommand( "NewTab", pathExtra, null, extraView );
    wnd.invokeCommand( "CloseAllButOne", tab2 );
 
    if( wnd.invokecommand( "WindowState" ) != 2 )
        wnd.invokecommand( "Restore");
}
else
{
    wnd = qs.invokeCommand( "NewWindow", pathDefault );
 
    qs.Sleep( 300 );
    var tab = wnd.invokeCommand( "NewTab", pathExtra, null, extraView );
    wnd.invokeCommand( "CloseAllButOne", tab );
}