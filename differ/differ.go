package differ

import (
	"fmt"
	"strings"
	"sync"

	"github.com/hashicorp/go-set/v2"
	"github.com/xuri/excelize/v2"
)

type Differ struct {
	fileA            *excelize.File
	fileB            *excelize.File
	fileASheetNames  *set.Set[string]
	fileBSheetNames  *set.Set[string]
	comparableSheets *set.Set[string]
}

func New(fileA, fileB *excelize.File) *Differ {
	fileASheetNames := newFileSheetSet(fileA.GetSheetList())
	fileBSheetNames := newFileSheetSet(fileB.GetSheetList())

	comparableSheets := fileASheetNames.Intersect(fileBSheetNames).(*set.Set[string])

	return &Differ{
		fileA:            fileA,
		fileB:            fileB,
		fileASheetNames:  fileASheetNames,
		fileBSheetNames:  fileBSheetNames,
		comparableSheets: comparableSheets,
	}
}

func newFileSheetSet(sheetList []string) *set.Set[string] {
	s := set.New[string](len(sheetList))
	s.InsertSlice(sheetList)
	return s
}

func (d *Differ) findDifferences() *sync.Map {
	differences := sync.Map{}

	var wg sync.WaitGroup
	d.comparableSheets.ForEach(func(sheetName string) bool {
		wg.Add(1)

		// Launch a goroutine for each Excel sheet
		go func(sheetName string, differences *sync.Map, wg *sync.WaitGroup) {
			defer wg.Done()

			maxRows, maxColumns := getMaxSheetBounds(d.fileA, d.fileB, sheetName)
			for i := 1; i <= maxRows; i++ {
				for j := 1; j <= maxColumns; j++ {
					cellName, _ := excelize.CoordinatesToCellName(j, i)
					valA, _ := d.fileA.GetCellValue(sheetName, cellName)
					valB, _ := d.fileB.GetCellValue(sheetName, cellName)

					if valA != valB {
						difference := fmt.Sprintf("%s: %s <> %s", cellName, valA, valB)
						if sheetDifferences, found := differences.Load(sheetName); found {
							differences.Store(sheetName, append(sheetDifferences.([]string), difference))
						} else {
							differences.Store(sheetName, []string{difference})
						}
					}
				}
			}
		}(sheetName, &differences, &wg)

		return true
	})

	wg.Wait()

	return &differences
}

func getMaxSheetBounds(fileA, fileB *excelize.File, sheetName string) (int, int) {
	maxRowsA, maxColumnsA := getSheetBounds(fileA, sheetName)
	maxRowsB, maxColumnsB := getSheetBounds(fileB, sheetName)

	maxRows := max(maxRowsA, maxRowsB)
	maxColumns := max(maxColumnsA, maxColumnsB)

	return maxRows, maxColumns
}

func getSheetBounds(file *excelize.File, sheetName string) (int, int) {
	dimension, _ := file.GetSheetDimension(sheetName)
	dimensionSplit := strings.Split(dimension, ":")

	var (
		maxColumns,
		maxRows int
	)

	if len(dimensionSplit) == 2 {
		maxColumns, maxRows, _ = excelize.CellNameToCoordinates(dimensionSplit[1])
	} else {
		maxColumns = 1
		maxRows = 1
	}

	return maxRows, maxColumns
}

func (d *Differ) printSheetMismatches() {
	printMissingSheets := func(fileName string, missingSheets set.Collection[string]) {
		missingSheets.ForEach(func(missing string) bool {
			fmt.Printf("CAUTION: File %s has no sheet called \"%s\"\n", fileName, missing)
			return true
		})
	}

	printMissingSheets(d.fileA.Path, d.fileASheetNames.Difference(d.fileBSheetNames))
	printMissingSheets(d.fileB.Path, d.fileBSheetNames.Difference(d.fileASheetNames))
}

func (d *Differ) PrintDiff() {
	d.printSheetMismatches()

	d.findDifferences().Range(func(sheetName any, sheetDifferences any) bool {
		fmt.Println(sheetName.(string))

		for _, difference := range sheetDifferences.([]string) {
			fmt.Printf("- %s\n", difference)
		}

		return true
	})
}
