package main

import (
	"math/rand"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

// init 种子数，保证生产随机字符串尽量不重复
func init() {
	rand.Seed(time.Now().UnixNano())
}

// letterRunes 生产随机密码的字符集
var letterRunes = []rune("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")

// RandStringRunes 生产随机字符串的函数
func RandStringRunes(n int) string {
	b := make([]rune, n)
	for i := range b {
		b[i] = letterRunes[rand.Intn(len(letterRunes))]
	}

	return string(b)
}

func main() {

	fRead, err := excelize.OpenFile("xxx.xlsx") // 读入的excel
	if err != nil {
		println(err.Error())
		return
	}

	f := excelize.NewFile()
	rows, err := fRead.GetRows("Sheet1")
	_ = f.SetCellValue("Sheet1", "A1", "账号")
	_ = f.SetCellValue("Sheet1", "B1", "密码")
	for i, row := range rows {
		if i <= 1 { // 此处i值可以改变，应该按照第一个学号开始的地方开始写入
			continue
		}

		areaA := "A" + strconv.Itoa(i)
		areaB := "B" + strconv.Itoa(i)
		_ = f.SetCellValue("Sheet1", areaA, row[1])
		_ = f.SetCellValue("Sheet1", areaB, RandStringRunes(10)) // n为密码长度
	}
	if err := f.SaveAs("2.xlsx"); err != nil { // 生产的excel
		println(err.Error())
	}

}
