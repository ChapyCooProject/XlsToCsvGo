package winmsg

import (
	"syscall"

	"github.com/lxn/win"
)

func UTF16PtrFromString(s string) *uint16 {
	result, _ := syscall.UTF16PtrFromString(s)
	return result
}

func ShowMsg(msg string, title string) {
	win.MessageBox(win.HWND(0), UTF16PtrFromString(msg), UTF16PtrFromString(title), win.MB_ICONINFORMATION+win.MB_OK)
}
