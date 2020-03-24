var wsh = new ActiveXObject("WScript.Shell")

for(var i=0 ; i < WScript.Arguments.length ; i++  ){
    var fname = WScript.Arguments(i)
    var shortcut = wsh.CreateShortCut(fname)
    WScript.Echo( shortcut.TargetPath )
    WScript.Echo( shortcut.WorkingDirectory )
}
