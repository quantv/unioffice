package convert2

import (
	"os"
	"regexp"
	"strings"

	"github.com/unidoc/unioffice/common/logger"
	"github.com/unidoc/unipdf/v3/model"
	"github.com/unidoc/unitype"
)

func registerFontsFromDirectory(dirName string) error {
	dir, err := os.Open(dirName)
	if err != nil {
		return err
	}
	defer dir.Close()
	files, err := dir.Readdirnames(0)
	if err != nil {
		return err
	}

	r, _ := regexp.Compile("^[a-z]*.ttf")

	for _, file := range files {
		if r.MatchString(file) {
			_eedg := dirName + "/" + file
			_aaee := registerFont(_eedg)
			if _aaee != nil {
				logger.Log.Debug("unable to process and register font from TTF file %s", _aaee)
				continue
			}
		}
	}
	return nil
}

var styleMap = map[string]FontStyle{
	"Regular":     FontStyle_Regular,
	"Bold":        FontStyle_Bold,
	"Italic":      FontStyle_Italic,
	"Bold Italic": FontStyle_BoldItalic,
}

func registerFont(file string) error {
	_acdg, _cfba := unitype.ParseFile(file)
	if _cfba != nil {
		logger.Log.Debug("Cannot parse TTF file %s", _cfba)
		return _cfba
	}
	font, err := model.NewCompositePdfFontFromTTFFile(file)
	if err != nil {
		return _cfba
	}
	_cec := _acdg.GetNameRecords()
	for _, _cae := range _cec {
		name := _cae[1]
		if name == "" {
			continue
		}
		_abfg := make([]byte, 0)
		for _eeda := 0; _eeda < len(name); _eeda++ {
			if name[_eeda] == 39 || name[_eeda] == 92 {
				continue
			}
			_dfd := 4
			if _eeda+_dfd < len(name) {
				if name[_eeda:_eeda+_dfd] == "\u0000" {
					_eeda = _eeda + _dfd + 1
					continue
				}
			}
			_abfg = append(_abfg, name[_eeda])
		}
		name = strings.Replace(string(_abfg), "x00", "", -1)
		style := _cae[2]
		if style == "" {
			continue
		}
		_abfg = make([]byte, 0)
		for _dbc := 0; _dbc < len(style); _dbc++ {
			if style[_dbc] == 39 || style[_dbc] == 92 {
				continue
			}
			_dae := 4
			if _dbc+_dae < len(style) {
				if style[_dbc:_dbc+_dae] == "\u0000" {
					_dbc = _dbc + _dae + 1
					continue
				}
			}
			_abfg = append(_abfg, style[_dbc])
		}
		style = strings.Replace(string(_abfg), "x00", "", -1)
		RegisterFont(name, styleMap[style], font)
	}
	return nil
}
