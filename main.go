package main

import (
	"flag"
	"fmt"
	"os"
	"xd/differ"

	"github.com/xuri/excelize/v2"
)

func main() {
	var (
		fileNameA string
		fileNameB string
		help      bool
	)

	// go run main.go -f1 data/InputA.xlsx -f2 data/InputB.xlsx
	flag.StringVar(&fileNameA, "f1", "", "File 1")
	flag.StringVar(&fileNameB, "f2", "", "File 2")
	flag.BoolVar(&help, "h", false, "Print help")
	flag.Parse()

	if help || fileNameA == "" || fileNameB == "" {
		flag.PrintDefaults()
		return
	}

	fileA := openOrExit(fileNameA)
	fileB := openOrExit(fileNameB)

	excelDiffer := differ.New(fileA, fileB)
	excelDiffer.PrintDiff()

	defer func() {
		if err := fileA.Close(); err != nil {
			fmt.Println(err)
		}
		if err := fileB.Close(); err != nil {
			fmt.Println(err)
		}
	}()
}

func openOrExit(fileName string) *excelize.File {
	file, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}

	return file
}
