package shortcut

import (
	"path/filepath"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// WShell is the OLE Object for "WScript.Shell"
type WShell struct {
	agent    *ole.IUnknown
	dispatch *ole.IDispatch
}

// NewWShell creates OLE Object for "WScript.Shell".
func NewWShell() (*WShell, error) {
	agent, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		return nil, err
	}
	dispatch, err := agent.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		agent.Release()
		return nil, err
	}
	return &WShell{agent: agent, dispatch: dispatch}, nil
}

// Close releases the OLE Object for "WScript.Shell".
func (wsh *WShell) Close() {
	wsh.dispatch.Release()
	wsh.agent.Release()
}

// Read reads the data of shortcut file. `path` must be absolute path.
func (wsh *WShell) Read(path string) (target string, workingdir string, err error) {
	shortcut, err := oleutil.CallMethod(wsh.dispatch, "CreateShortCut", path)
	if err != nil {
		return "", "", err
	}
	shortcutDis := shortcut.ToIDispatch()
	defer shortcutDis.Release()
	targetPath, err := oleutil.GetProperty(shortcutDis, "TargetPath")
	if err != nil {
		return "", "", err
	}
	workingDir, err := oleutil.GetProperty(shortcutDis, "WorkingDirectory")
	if err != nil {
		return "", "", err
	}
	return targetPath.ToString(), workingDir.ToString(), err
}

// Read reads the data of shortcut file. `path` can be relative path.
func Read(path string) (targetPath string, workingDir string, err error) {
	path, err = filepath.Abs(path)
	if err != nil {
		return "", "", err
	}
	wsh, err := NewWShell()
	if err != nil {
		return "", "", err
	}
	defer wsh.Close()

	return wsh.Read(path)
}

// Make makes a shortcut file.`from`,`to` must be absolute path.
func (wsh *WShell) Make(from, to, dir string) error {
	shortcut, err := oleutil.CallMethod(wsh.dispatch, "CreateShortCut", to)
	if err != nil {
		return err
	}
	shortcutDis := shortcut.ToIDispatch()
	defer shortcutDis.Release()
	_, err = oleutil.PutProperty(shortcutDis, "TargetPath", from)
	if err != nil {
		return err
	}
	_, err = oleutil.PutProperty(shortcutDis, "WorkingDirectory", dir)
	if err != nil {
		return err
	}
	_, err = oleutil.CallMethod(shortcutDis, "Save")
	return err
}

// Make makes a shortcut file. `from`,`to` can be relative paths
func Make(from, to, dir string) error {
	from, err := filepath.Abs(from)
	if err != nil {
		return err
	}
	to, err = filepath.Abs(to)
	if err != nil {
		return err
	}
	wsh, err := NewWShell()
	if err != nil {
		return err
	}
	defer wsh.Close()

	return wsh.Make(from, to, dir)
}
