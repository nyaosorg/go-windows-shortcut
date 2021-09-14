package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	_ "github.com/mattn/getwild"
	"github.com/nyaosorg/go-windows-shortcut"
)

func mains(args []string) error {
	if len(args) <= 0 {
		return fmt.Errorf("Usage: %s SOURCE-PATH TARGET-PATH(.lnk)", os.Args[0])
	}
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	if len(args) == 1 {
		target, _, err := shortcut.Read(args[0])
		if err != nil {
			return err
		}
		fmt.Println(target)
		return nil
	}
	to := args[len(args)-1]
	toAbsPath, err := filepath.Abs(to)
	toStat, toStatErr := os.Stat(toAbsPath)
	if len(args) == 2 {
		if toStatErr == nil && !toStat.IsDir() {
			return fmt.Errorf("%s: file already exists", to)
		}
	} else {
		if toStatErr != nil {
			return toStatErr
		}
		if !toStat.IsDir() {
			return fmt.Errorf("%s: file already exists", to)
		}
	}

	wd, err := os.Getwd()
	if err != nil {
		return err
	}
	for _, fromPath := range args[:len(args)-1] {
		fromAbsPath, err := filepath.Abs(fromPath)
		if err != nil {
			return err
		}
		toPath := toAbsPath
		if toStatErr == nil {
			toPath = filepath.Join(toAbsPath, filepath.Base(fromAbsPath))
		}
		if !strings.EqualFold(filepath.Ext(toPath), ".lnk") {
			toPath += ".lnk"
		}
		if err := shortcut.Make(fromAbsPath, toPath, wd); err != nil {
			return err
		}
		fmt.Printf("%s -> %s\n", fromAbsPath, toPath)
	}
	return nil
}

func main() {
	if err := mains(os.Args[1:]); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
}
