package shortcut

// Read reads the data of shortcut file. `path` can be relative path.
func Read(path string) (targetPath string, workingDir string, err error) {
	return _read(path)
}

// Make makes a shortcut file. `from`,`to` can be relative paths
func Make(from, to, dir string) error {
	return _make(from, to, dir)
}
