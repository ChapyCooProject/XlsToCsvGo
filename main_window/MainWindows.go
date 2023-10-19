package main

import (
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	"MyProjects.com/XlsToCsv/conv_excel"

	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
	"github.com/lxn/win"
)

var OnSizeChangedFlag bool = true
var OnBoundsChangedFlag bool = true

// main
func main() {

	// MyMainWindow構造体をポインタ変数mwとして定義
	mw := &MyMainWindow{}

	// フォームのサイズ
	windowWidth := 320
	windowHeight := 200

	// モニタの縦と横のサイズを取得
	scrWidth := int(win.GetSystemMetrics(win.SM_CXSCREEN))
	scrHeight := int(win.GetSystemMetrics(win.SM_CYSCREEN))

	if _, err := (MainWindow{
		AssignTo: &mw.MainWindow,
		Title:    "Convert Excel To CSV",
		Size:     Size{Width: windowWidth, Height: windowHeight},
		OnSizeChanged: func() {
			if OnSizeChangedFlag {
				mw.MainWindow.SetBounds(walk.Rectangle{
					X:      int((scrWidth - windowWidth) / 2),
					Y:      int((scrHeight - windowHeight) / 2),
					Width:  windowWidth,
					Height: windowHeight,
				})
				OnSizeChangedFlag = false
			}
		},
		OnBoundsChanged: func() {
			if OnBoundsChangedFlag {
				mw.rbComma.SetChecked(true)
				OnBoundsChangedFlag = false
			}
		},
		Layout: VBox{},
		Children: []Widget{
			Label{
				Font:     Font{PointSize: 10},
				AssignTo: &mw.lblComment,
				Text:     "変換したいExcelファイルを指定してください。",
			},
			PushButton{
				Text:      "Excelファイルを選択...",
				Font:      Font{PointSize: 10},
				AssignTo:  &mw.btnConvert,
				OnClicked: mw.clicked_openFileDialog,
			},
			RadioButtonGroupBox{
				Layout: HBox{},
				Buttons: []RadioButton{
					{
						AssignTo: &mw.rbComma,
						Text:     "コンマ区切り",
						Value:    "c",
					},
					{
						AssignTo: &mw.rbTab,
						Text:     "タブ区切り",
						Value:    "t",
					},
				},
			},
			Composite{
				Layout: HBox{},
				Children: []Widget{
					Label{
						Font: Font{PointSize: 10},
						Text: "解除パスワード（必要時）:",
					},
					LineEdit{
						AssignTo:     &mw.editPassword,
						PasswordMode: true,
					},
				},
			},
		},
	}.Run()); err != nil {
		log.Fatal(err)
	}

}

// MyMainWindow構造体
type MyMainWindow struct {
	*walk.MainWindow
	lblComment   *walk.Label
	btnConvert   *walk.PushButton
	rbComma      *walk.RadioButton
	rbTab        *walk.RadioButton
	editPassword *walk.LineEdit
}

func InitLog(outputPath string) {

	// ログ出力 初期設定
	logPath := filepath.Dir(outputPath)
	logPath = filepath.Join(logPath, "log")
	_, err := os.Stat(logPath)
	if os.IsNotExist(err) {
		os.Mkdir(logPath, 0777)
	}
	logPath = filepath.Join(logPath, "log"+time.Now().Format("20060102")+".txt")
	logFile, _ := os.OpenFile(logPath, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0666)
	logwriter := io.MultiWriter(logFile, os.Stdout)
	log.SetFlags(log.Ldate | log.Ltime | log.Llongfile)
	log.SetOutput(logwriter)

}

func (mw *MyMainWindow) clicked_openFileDialog() {

	// カレントディレクトリのフルパスを取得
	exePath, err := os.Executable()
	if err != nil {
		log.Panic(err)
	}

	// ログ出力 初期設定
	InitLog(exePath)

	// 区切り文字の指定
	commaStr := "c"
	if mw.rbComma.Checked() {
		commaStr = "c"
	}
	if mw.rbTab.Checked() {
		commaStr = "t"
	}

	// ファイルダイアログ表示
	dlg := new(walk.FileDialog)
	dlg.InitialDirPath = exePath
	dlg.FilePath = exePath
	dlg.Title = "変換元のExcelファイルを選択"
	dlg.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

	if ok, err := dlg.ShowOpen(mw); err != nil {
		walk.MsgBox(mw, "", "Error : File Open", walk.MsgBoxOK+walk.MsgBoxIconError)
		return
	} else if !ok {
		walk.MsgBox(mw, "", "Cancel", walk.MsgBoxOK+walk.MsgBoxIconError)
		return
	}

	excelFilePath := dlg.FilePath

	//取得したパスから拡張子を取得
	fileName := filepath.Base(excelFilePath)
	extension := filepath.Ext(fileName)
	extension = strings.ToLower(extension) // 小文字へ変換

	errFlag := false
	errText := ""
	if extension == ".xlsx" {
		if err := conv_excel.ConvExcelXlsx(excelFilePath, mw.editPassword.Text(), commaStr); err != nil {
			errFlag = true
			errText = err.Error()
		}
	} else {
		walk.MsgBox(mw, "ファイル形式エラー", "指定したファイルは、Excelファイルではありません。", walk.MsgBoxOK+walk.MsgBoxIconError)
		return
	}

	mw.editPassword.SetText("")

	if errFlag == true {
		walk.MsgBox(mw, "", "ExcelファイルからCSVファイルへの変換に失敗しました。\n\n"+errText, walk.MsgBoxOK+walk.MsgBoxIconError)
	} else {
		walk.MsgBox(mw, "", "ExcelファイルからCSVファイルへ変換しました。", walk.MsgBoxOK+walk.MsgBoxIconInformation)
	}

}
