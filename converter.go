package exceliser

import (
	"encoding/csv"
	"os"
	"strconv"
	
	"github.com/xuri/excelize/v2"
)

func CSVtoExcel(path string,newPath string,returnFile bool)(file *excelize.File,err error){
	csvfile,err := os.Open(path)
	defer csvfile.Close()

	if err != nil {
		return nil,err
	}

	reader := csv.NewReader(csvfile)
	fileStructures,err := reader.ReadAll()

	if err != nil {
		return nil,err
	}

	newFile := excelize.NewFile()

	for i,fileStructure := range fileStructures {
		for j,cell := range fileStructure {
			newFile.SetCellStr("Sheet1",getColumnName(j+1)+strconv.Itoa(i+1), cell)
		}
	}

	if err := newFile.SaveAs(newPath); err != nil {
		return nil,err
	}

	if returnFile {
		newCreatedFile,err := excelize.OpenFile(newPath); return newCreatedFile,err
	}

	return nil,nil
}

func getColumnName(col int) string {
    name := make([]byte, 0, 3) // max 16,384 columns (2022)
    const aLen = 'Z' - 'A' + 1 // alphabet length
    for ; col > 0; col /= aLen + 1 {
        name = append(name, byte('A'+(col-1)%aLen))
    }
    for i, j := 0, len(name)-1; i < j; i, j = i+1, j-1 {
        name[i], name[j] = name[j], name[i]
    }
    return string(name)
}