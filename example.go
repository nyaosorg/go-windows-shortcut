// +build run

package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/zetamatta/go-windows-shortcut"
)

func main1() error {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	if len(os.Args) < 3 {
		return fmt.Errorf("Usage: go run example.go SRC DST")
	}
	src1, err := filepath.Abs(os.Args[1])
	if err != nil {
		return err
	}
	dst1, err := filepath.Abs(os.Args[2])
	if err != nil {
		return err
	}
	if strings.HasSuffix(strings.ToUpper(dst1), ".lnk") {
		dst1 += ".lnk"
	}
	fmt.Printf("%s --> %s\n", src1, dst1)
	if err := shortcut.Make(src1, dst1, ""); err != nil {
		return err
	}
	src2, dir, err := shortcut.Read(dst1)
	if err != nil {
		return err
	}
	fmt.Printf("%s <-- %s @ %s\n", dst1, src2, dir)
	return nil
}

func main() {
	if err := main1(); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
}
