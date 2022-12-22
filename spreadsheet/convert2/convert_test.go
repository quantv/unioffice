package convert2

import (
	"fmt"
	"testing"

	"github.com/unidoc/unioffice/common/logger"
	"github.com/unidoc/unioffice/spreadsheet"
)

func TestPdf(t *testing.T) {
	logger.SetLogger(logger.NewConsoleLogger(logger.LogLevelDebug))
	RegisterFontFromFile("Times New Roman", FontStyle_Regular, "/home/quantv/Downloads/tmp/font-times-new-roman/SVN-Times New Roman 2.ttf")
	RegisterFontFromFile("Times New Roman", FontStyle_Bold, "/home/quantv/Downloads/tmp/font-times-new-roman/SVN-Times New Roman 2 bold.ttf")
	RegisterFontFromFile("Times New Roman", FontStyle_BoldItalic, "/home/quantv/Downloads/tmp/font-times-new-roman/SVN-Times New Roman 2 bold italic.ttf")
	RegisterFontFromFile("Times New Roman", FontStyle_Italic, "/home/quantv/Downloads/tmp/font-times-new-roman/SVN-Times New Roman 2 italic.ttf")
	wb, _ := spreadsheet.Open("/home/quantv/Downloads/tmp/pdf.xlsx")
	sh := wb.Sheets()[0]

	fmt.Println(sh.Cell("A22").GetString())

	c := ConvertToPdf(&sh)
	c.WriteToFile("/home/quantv/Downloads/tmp/file3.pdf")
}
