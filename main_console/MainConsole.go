package main

//ライブラリImport
import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"MyProjects.com/XlsToCsv/conv_excel"
)

// main
func main() {

	var fDelimiters *string = flag.String("d", "comma", "区切り文字を指定。comma:コンマ区切り、tab:タブ区切り")
	var fPassword *string = flag.String("p", "", "パスワードを指定。")
	var fExcelFile *string = flag.String("f", "", "変換するExcelファイルを指定。（ダブルクォーテーションで囲むこと。）")
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, "ExcelファイルをCSVファイルに変換します。\n\n%s -d=[commma or tab] -p=[パスワード] -f=[Excelファイル]\n\n", filepath.Base(os.Args[0]))
		flag.PrintDefaults()
	}
	flag.Parse()

	// 区切り文字の指定
	commaStr := "c"
	if *fDelimiters == "comma" {
		commaStr = "c"
	}
	if *fDelimiters == "tab" {
		commaStr = "t"
	}

	// 引数確認用
	// fmt.Printf("param -d : %s\n", *fDelimiters)
	// fmt.Printf("param -p : %s\n", *fPassword)
	// fmt.Printf("param -f : %s\n", *fExcelFile)
	// for i := 0; i < len(flag.Args()); i++ {
	// 	fmt.Println("other("+strconv.Itoa(i)+"): ", flag.Arg(i))
	// }

	excelFilePath := *fExcelFile
	if _, err := os.Stat(excelFilePath); err != nil {
		fmt.Println("Excelファイルが見つかりません。")
		return
		// os.Exit(2)
	}

	//取得したパスから拡張子を取得
	fileName := filepath.Base(excelFilePath)
	extension := filepath.Ext(fileName)
	extension = strings.ToLower(extension) // 小文字へ変換

	errFlag := true
	if extension == ".xlsx" {
		if err := conv_excel.ConvExcelXlsx(excelFilePath, *fPassword, commaStr); err == nil {
			errFlag = false
		}
	} else {
		fmt.Println("指定したファイルは、Excelファイルではありません。")
		return
	}
	// if extension == ".xls" {
	// 	if err := conv_excel.ConvExcelXls(excelFilePath, commaStr); err == nil {
	// 		errFlag = false
	// 	}
	// }

	if errFlag == false {
		fmt.Println("Excelファイルを正常に変換しました。")
	} else {
		fmt.Println("Error:Excelファイルの変換に失敗しました。")
	}

}
