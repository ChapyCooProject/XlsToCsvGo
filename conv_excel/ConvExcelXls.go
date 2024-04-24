package conv_excel

//ライブラリImport
import (
	"encoding/csv"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/extrame/xls"
)

func ConvExcelXls(excelFilePath string, commaStr string) error {

	// CSVファイル出力先
	outputDir := filepath.Dir(excelFilePath)                      // 絶対パス
	fileName := filepath.Base(excelFilePath)                      // ファイル名
	extension := filepath.Ext(fileName)                           // 拡張子
	fileNameWithoutExt := strings.TrimSuffix(fileName, extension) // ファイル名（拡張子なし）
	tmpPath := filepath.Join(outputDir, fileNameWithoutExt+".tmp")
	csvPath := filepath.Join(outputDir, fileNameWithoutExt+".csv")
	// winmsg.ShowMsg(csvPath)

	// Excelファイル（.xls）を開く
	xlFile, err := xls.Open(excelFilePath, "utf-8")
	if err != nil {
		log.Println("[Error]", err)
		return err
	}

	xlSheet := xlFile.GetSheet(0)
	if xlSheet != nil {

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
		for row := 0; row < int(xlSheet.MaxRow); row++ {
			rowData := xlSheet.Row(row)
			var recData []string
			for col := rowData.FirstCol(); col < rowData.LastCol(); col++ {
				cellValue := rowData.Col(col)
				recData = append(recData, cellValue)
			}
			csvwriter.Write(recData)
		}

		csvfile.Close()

		// 拡張子を「.tmp」から「.csv」へ変更
		if err := os.Rename(tmpPath, csvPath); err != nil {
			log.Println("[Error]", err)
			return err
		}

	}

	// log.Fatal(err) // deferが処理されない os.Exit(1)
	// log.Panic(err) // deferが処理される os.Exit(2)

	return nil
}
