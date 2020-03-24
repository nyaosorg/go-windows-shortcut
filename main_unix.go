// +build !windows

package shortcut

import (
	"errors"
)

func _read(path string) (targetPath string, workingDir string, err error) {
	return "", "", errors.New("not supported OS")
}

func _make(from, to, dir string) error {
	return errors.New("not supported OS")
}
