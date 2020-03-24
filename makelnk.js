var args = WScript.Arguments
if ( args.length < 2 ){
    WScript.Echo("two few arguments")
    WScript.Quit()
}
var wsh = new ActiveXObject("WScript.Shell")

var shortcut = wsh.CreateShortCut(args(0))
shortcut.TargetPath = args(1)
if ( args.length >= 2 ){
    shortcut.WorkingDirectory = args(2)
}
shortcut.Save()
