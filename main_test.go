package shortcut

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"github.com/go-ole/go-ole"
)

func testMake() error {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	wd, err := os.Getwd()
	if err != nil {
		return fmt.Errorf("os.Getwd: %w", err)
	}

	lnkPath := filepath.Join(os.TempDir(), "testShortCut.lnk")
	srcPath := filepath.Join(wd, "main.go")

	err = Make(srcPath, lnkPath, wd)
	if err != nil {
		return fmt.Errorf("Make: %w", err)
	}
	defer os.Remove(lnkPath)

	output, err := exec.Command("cscript", "/nologo", "readlnk.js", lnkPath).Output()
	if err != nil {
		return fmt.Errorf("Command('cscript'): %w", err)
	}
	outputStr := strings.Split(strings.TrimSpace(string(output)), "\n")
	if len(outputStr) < 2 {
		return fmt.Errorf("readlnk.js: two few lines: '%s'", outputStr)
	}
	srcPath2 := strings.TrimSpace(outputStr[0])
	if srcPath2 != srcPath {
		return fmt.Errorf("Make: TargetPath differ: '%s' != '%s'",
			srcPath2, srcPath)
	}
	// println(srcPath2)
	// println(srcPath)
	wd2 := strings.TrimSpace(outputStr[1])
	if wd2 != wd {
		return fmt.Errorf("Make: WorkingDirectory differ: '%s' != '%s'",
			wd2, wd)
	}
	// println(wd2)
	// println(wd)
	return nil
}

func TestMake(t *testing.T) {
	if err := testMake(); err != nil {
		t.Fatal(err.Error())
	}
}

func testRead() error {
	wd, err := os.Getwd()
	if err != nil {
		return fmt.Errorf("os.Getwd: %w", err)
	}

	lnkPath := filepath.Join(os.TempDir(), "testShortCut.lnk")
	srcPath := filepath.Join(wd, "main.go")

	err = exec.Command("cscript", "/nologo", "makelnk.js", lnkPath, srcPath, wd).Run()
	if err != nil {
		return fmt.Errorf("makelnk.js: %w", err)
	}
	defer os.Remove(lnkPath)

	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	targetPath, wd2, err := Read(lnkPath)
	if err != nil {
		return fmt.Errorf("Read: %w", err)
	}
	if targetPath != srcPath {
		return fmt.Errorf("Read: TargetPath: '%s' != '%s'", targetPath, srcPath)
	}
	// println(targetPath)
	// println(srcPath)
	if wd2 != wd {
		return fmt.Errorf("Read: WorkingDirectory: '%s' != '%s'", wd2, wd)
	}
	// println(wd2)
	// println(wd)
	return nil
}

func TestRead(t *testing.T) {
	if err := testRead(); err != nil {
		t.Fatal(err.Error())
	}
}
