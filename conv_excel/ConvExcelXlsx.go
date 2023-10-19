package conv_excel

import (
	"encoding/csv"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func ConvExcelXlsx(excelFilePath string, excelPassword string, commaStr string) error {

	// CSVファイル出力先
	outputDir := filepath.Dir(excelFilePath)                      // 絶対パス
	fileName := filepath.Base(excelFilePath)                      // ファイル名
	extension := filepath.Ext(fileName)                           // 拡張子
	fileNameWithoutExt := strings.TrimSuffix(fileName, extension) // ファイル名（拡張子なし）
	tmpPath := filepath.Join(outputDir, fileNameWithoutExt+".tmp")
	csvPath := filepath.Join(outputDir, fileNameWithoutExt+".csv")

	// Excelファイル（.xlsx）を開く
	var xlFile *excelize.File
	var xlErr error
	if excelPassword == "" {
		xlFile, xlErr = excelize.OpenFile(excelFilePath)
	} else {
		xlFile, xlErr = excelize.OpenFile(excelFilePath, excelize.Options{Password: excelPassword})
	}
	if xlErr != nil {
		log.Println("[Error]", xlErr)
		return xlErr
	}

	// 本関数「func ConvExcelXlsx()」の処理を抜けるタイミングで実行
	defer func() {
		if err := xlFile.Close(); err != nil {
			log.Println("[Error]", err)
		}
	}()

	// 先頭のシート名を取得
	var sheetName string = xlFile.GetSheetName(0)

	// 指定したシートのすべてのセルを取得
	rows, err := xlFile.GetRows(sheetName)
	if err != nil {
		log.Println("[Error]", err)
		return err
	}

	// 出力用にファイルをオープン
	csvfile, err := os.Create(tmpPath)
	if err != nil {
		log.Panic("[Panic]", err)
	}
	defer csvfile.Close()

	// CSV 形式でデータを書き込む
	csvwriter := csv.NewWriter(csvfile)
	if commaStr == "t" {
		csvwriter.Comma = '\t'
	}
	csvwriter.WriteAll(rows)
	csvfile.Close()

	// 拡張子を「.tmp」から「.csv」へ変更
	if err := os.Rename(tmpPath, csvPath); err != nil {
		log.Println("[Error]", err)
		return err
	}

	return nil
}
