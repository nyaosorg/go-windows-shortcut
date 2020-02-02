// +build run

package main

import (
	"fmt"
	"os"

	"github.com/go-ole/go-ole"
	"github.com/zetamatta/go-windows-shortcut"
)

func main1() error {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	if len(os.Args) < 3 {
		return fmt.Errorf("Usage: go run example.go SRC DST")
	}
	src1 := os.Args[1]
	dst1 := os.Args[2]

	fmt.Printf("make shortcut: %s --> %s\n", src1, dst1)
	if err := shortcut.Make(src1, dst1, ""); err != nil {
		return err
	}

	src2, _, err := shortcut.Read(dst1)
	if err != nil {
		return err
	}
	fmt.Printf("read shortcut: %s <-- %s\n", dst1, src2)
	return nil
}

func main() {
	if err := main1(); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
}
