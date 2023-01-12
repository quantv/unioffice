package convert2

import (
	_a "image"
	"math"
	"strconv"

	_ce "github.com/unidoc/unioffice/common/logger"
	_ba "github.com/unidoc/unioffice/common/tempstorage"
	_gb "github.com/unidoc/unioffice/internal/convertutils"
	_ae "github.com/unidoc/unioffice/measurement"
	_be "github.com/unidoc/unioffice/schema/soo/dml"
	_cf "github.com/unidoc/unioffice/schema/soo/dml/chart"
	_bb "github.com/unidoc/unioffice/schema/soo/ofc/sharedTypes"
	"github.com/unidoc/unioffice/schema/soo/sml"
	_d "github.com/unidoc/unioffice/spreadsheet"
	_f "github.com/unidoc/unioffice/spreadsheet/reference"
	"github.com/unidoc/unipdf/v3/creator"
	"github.com/unidoc/unipdf/v3/model"
)

func (_bdba *convertContext) determineMaxIndexes() {
	var _deg, _bga int
	_deg = int(_bdba.sheet.MaxColumnIdx())
	_adb := _bdba.sheet.Rows()
	if len(_adb) > 0 {
		_bga = int(_adb[len(_adb)-1].RowNumber())
	}
	for _, _gad := range _bdba._abba {
		if _gad._cfdc >= _bga {
			_bga = _gad._cfdc + 1
		}
		if _gad._eacd >= _deg {
			_deg = _gad._eacd + 1
		}
	}
	_bdba._dbc = _bga
	_bdba._face = _deg
}

// ConvertToPdf converts a sheet to a PDF file. This package is beta, breaking changes can take place.
func ConvertToPdf(sheet *_d.Sheet) *creator.Creator {
	sheetX := sheet.X()
	if sheetX == nil {
		return nil
	}
	var pagesize creator.PageSize
	var _acc bool
	if _cb := sheetX.PageSetup; _cb != nil {
		if _bac := _cb.PaperSizeAttr; _bac != nil {
			pagesize = _egac[*_bac]
		}
		_acc = _cb.OrientationAttr == sml.ST_OrientationLandscape
	}
	if (pagesize == creator.PageSize{}) {
		pagesize = _egac[1]
	}
	if _acc {
		pagesize[0], pagesize[1] = pagesize[1], pagesize[0]
	}
	pdfCreator := creator.New()
	pdfCreator.SetPageSize(pagesize)
	var topAttr, bottomAttr, leftAttr, rightAttr float64
	if _gab := sheetX.PageMargins; _gab != nil {
		leftAttr = _gab.LeftAttr
		rightAttr = _gab.RightAttr
		topAttr = _gab.TopAttr
		bottomAttr = _gab.BottomAttr
	}
	if leftAttr < 0.25 {
		leftAttr = 0.25
	}
	if rightAttr < 0.25 {
		rightAttr = 0.25
	}
	topAttr *= _ae.Inch
	bottomAttr *= _ae.Inch
	leftAttr *= _ae.Inch
	rightAttr *= _ae.Inch
	pdfCreator.SetPageMargins(leftAttr, rightAttr, topAttr, bottomAttr)
	_fb := sheet.Workbook()
	var _ad *_be.Theme
	if len(_fb.Themes()) > 0 {
		_ad = _fb.Themes()[0]
	}
	var scale uint32 = 100
	if sheetX.PageSetup != nil && sheetX.PageSetup.ScaleAttr != nil {
		scale = *sheetX.PageSetup.ScaleAttr
	}
	_ef := &convertContext{
		creator:    pdfCreator,
		sheet:      sheet,
		workbook:   sheet.Workbook(),
		_ggce:      _ad,
		_aeg:       &sheet.Workbook().StyleSheet,
		_bcag:      topAttr,
		_ebae:      leftAttr,
		pageHeight: pagesize[1] - bottomAttr - topAttr,
		pageWidth:  pagesize[0] - rightAttr - leftAttr,
		scale:      scale,
	}
	_ef.makeAnchors()
	_ef.determineMaxIndexes()
	if _ef._dbc == 0 && _ef._face == 0 {
		pdfCreator.NewPage()
		return pdfCreator
	}
	_ef.makeCols()
	_ef.makeRows()
	_ef.makeMergedCells()
	_ef.makeCells()
	_ef.makePagespans()
	_ef.makeRowspans()
	_ef.makePages()
	_ef.fillPages()
	_ef.distributeAnchors()
	_ef.drawSheet()
	return pdfCreator
}

type style struct {
	_gege    *string
	fontSize *float64
	fontName *string
	bold     *bool
	_gdgae   *bool
	_bcd     *bool
	_ffab    *bool
	_eac     *bool
	border1  *border
	border2  *border
	border3  *border
	border4  *border
	_gea     bool
	_egd     sml.ST_VerticalAlignment
	_cbef    sml.ST_HorizontalAlignment
}

func (_afeb *convertContext) imageFromAnchor(_acf *anchor, _gca, _bdcb float64) _a.Image {
	if _acf._cgaf != nil {
		return _acf._cgaf
	}
	if _acf._fagc != nil {
		_adfd, _ggcc := _gb.MakeImageFromChartSpace(_acf._fagc, _gca, _bdcb, _afeb._ggce)
		if _ggcc != nil {
			_ce.Log.Debug("C\u0061\u006e\u006e\u006f\u0074\u0020\u006d\u0061\u006b\u0065\u0020\u0061\u006e\u0020\u0069\u006d\u0061\u0067e\u0020\u0066\u0072\u006f\u006d\u0020\u0063\u0068\u0061\u0072tS\u0070\u0061\u0063e\u003a \u0025\u0073", _ggcc)
			return nil
		}
		return _adfd
	}
	return nil
}
func _bgfc(_symbol *symbol) {
	_creator := creator.New()
	_styleParagraph := _creator.NewStyledParagraph()
	_styleParagraph.SetMargins(0, 0, 0, 0)
	_dee := _styleParagraph.Append(_symbol.value)
	if _symbol._fae != nil {
		_dee.Style = *_symbol._fae
	}
	_symbol._fgfag = _styleParagraph.Height()
	if _symbol._fceca == 0 {
		_symbol._fceca = _styleParagraph.Width()
	}
}
func (_dda *convertContext) fillPages() {
	for _fag, _abb := range _dda._gbg {
		_cec := _dda.rowInfo[_abb._aebb:_abb._fbee]
		for _cbg, _efb := range _cec {
			_cgb := 0
			_cbec := 0.0
			_bbe := []*cell{}
			if _efb._bgc {
				for _, _bcaa := range _efb.cells {
					_bed := _dda.pages_span[_cgb]
					_dda._dfdc = _bed._eed[_fag]
					_dda._dfdc._ebbf = true
					_ccd := _bcaa._bcf
					if _cbec+_ccd > _bed._ebed {
						_dda.addRowToPage(_bbe, _cbg)
						_bbe = []*cell{_bcaa}
						_cbec = _ccd
						_cgb++
					} else {
						_bcaa._bffg = _cbec
						_bbe = append(_bbe, _bcaa)
						_cbec += _ccd
					}
				}
				if len(_bbe) > 0 {
					_egb := _dda.pages_span[_cgb]
					_dda._dfdc = _egb._eed[_fag]
					_dda._dfdc._ebbf = true
					_dda.addRowToPage(_bbe, _cbg)
				}
			}
		}
	}
}
func (_cgg *convertContext) getContentFromCell(cell _d.Cell, _eged *style, width float64, _gdeg bool) ([]*line, sml.ST_CellType) {
	_gdd := cell.X()
	var cellSymbols []*symbol
	switch _gdd.TAttr {
	case sml.ST_CellTypeS:
		_ggf := _gdd.V
		if _ggf != nil {
			_ega, _aceg := strconv.Atoi(*_ggf)
			if _aceg == nil {
				_cba := _cgg.workbook.SharedStrings.X().Si[_ega]
				if _cba.T != nil {
					cellSymbols = _cgg.getSymbolsFromString(*_cba.T, _eged)
				} else if _cba.R != nil {
					cellSymbols = _cgg.getSymbolsFromR(_cba.R, _eged)
				}
			}
		}
	case sml.ST_CellTypeB:
		_gfag := _gdd.V
		if _gfag != nil {
			if *_gfag == "\u0030" {
				cellSymbols = _cgg.getSymbolsFromString("\u0046\u0041\u004cS\u0045", _eged)
			} else {
				cellSymbols = _cgg.getSymbolsFromString("\u0054\u0052\u0055\u0045", _eged)
			}
		}
	case sml.ST_CellTypeStr:
		cellSymbols = _cgg.getSymbolsFromString(cell.GetFormattedValue(), _eged)
	default:
		cellSymbols = _cgg.getSymbolsFromString(cell.GetFormattedValue(), _eged)
	}
	_bff := 0.0
	start := 0.0
	var lines []*line
	var _gdaf bool
	if _eged != nil {
		if _eged._ffab != nil {
			if *_eged._ffab {
				_gdaf = true
			}
		}
		if _eged._eac != nil {
			if *_eged._eac {
				_gdaf = true
			}
		}
	}
	if _gdeg {
		lines = []*line{}
		_acb := width - 2*_eb
		symbols := []*symbol{}
		for _, symbol1 := range cellSymbols {
			_bgfc(symbol1)
			if symbol1.value == "\n" {
				_dbea := findMax(symbols)
				//control line spacing.
				_dbea -= 2
				lines = append(lines, &line{lineSpace: start, symbols: symbols, fontSize: _dbea})
				symbols = []*symbol{symbol1}
				_bff = symbol1._fceca
				start += _dbea
			} else if _bff+symbol1._fceca >= _acb {
				_dbea := findMax(symbols)
				if _gdaf {
					_dbea /= _ga
				}
				lines = append(lines, &line{lineSpace: start, symbols: symbols, fontSize: _dbea})
				symbols = []*symbol{symbol1}
				_bff = symbol1._fceca
				start += _dbea
				//need to increase row height if need.
			} else {
				symbol1.left = _bff
				_bff += symbol1._fceca
				symbols = append(symbols, symbol1)
			}
		}
		_cbgf := findMax(symbols)
		if _gdaf {
			_cbgf /= _ga
		}
		if len(symbols) > 0 {
			lines = append(lines, &line{lineSpace: start, symbols: symbols, fontSize: _cbgf})
		}
	} else {
		for _, _agd := range cellSymbols {
			_bgfc(_agd)
			_agd.left = _bff
			_bff += _agd._fceca
		}
		if len(cellSymbols) > 0 {
			lines = []*line{&line{symbols: cellSymbols, fontSize: findMax(cellSymbols)}}
		}
	}
	_daf := _gdd.TAttr
	if _daf == sml.ST_CellTypeUnset {
		_daf = sml.ST_CellTypeN
	}
	return lines, _daf
}

// RegisterFontsFromDirectory registers all fonts from the given directory automatically detecting font families and styles.
func RegisterFontsFromDirectory(dirName string) error { return registerFontsFromDirectory(dirName) }
func (ctx *convertContext) getStyleFromRPrElt(cellStyle *sml.CT_RPrElt) *style {
	if cellStyle == nil {
		return nil
	}
	style1 := &style{}
	style1.fontName = &cellStyle.RFont.ValAttr
	if _dbbc := cellStyle.B; _dbbc != nil {
		_ccabf := _dbbc.ValAttr == nil || *_dbbc.ValAttr
		style1.bold = &_ccabf
	}
	if _bcb := cellStyle.I; _bcb != nil {
		_gbfg := _bcb.ValAttr == nil || *_bcb.ValAttr
		style1._gdgae = &_gbfg
	}
	if _ddb := cellStyle.U; _ddb != nil {
		_ddfc := _ddb.ValAttr == sml.ST_UnderlineValuesSingle || _ddb.ValAttr == sml.ST_UnderlineValuesUnset
		style1._bcd = &_ddfc
	}
	if _dgee := cellStyle.VertAlign; _dgee != nil {
		_dgea := _dgee.ValAttr == _bb.ST_VerticalAlignRunSuperscript
		style1._ffab = &_dgea
		_aef := _dgee.ValAttr == _bb.ST_VerticalAlignRunSubscript
		style1._eac = &_aef
	}
	if _cad := cellStyle.Sz; _cad != nil {
		fontsize := _cad.ValAttr / 12 * _gb.DefaultFontSize
		style1.fontSize = &fontsize
	}
	if _bfdd := cellStyle.Color; _bfdd != nil {
		style1._gege = ctx.getColorStringFromSmlColor(_bfdd)
	}
	return style1
}
func (context *convertContext) makeCols() {
	sheet := context.sheet
	x := sheet.X()
	colInfos := []*colInfo{}
	col_width_ranges := []colWidthRange{}
	if ct_col := x.Cols; len(ct_col) > 0 {
		for _, col := range ct_col[0].Col {
			width := 65.0
			if _fbe := col.WidthAttr; _fbe != nil {
				//Chieu rong tinh bang so characters
				if *_fbe > 0.83 {
					*_fbe -= 0.83
				}
				if *_fbe <= 1 {
					width = *_fbe * 11
				} else {
					//width = 5 + *_fbe*6
					width = 5 + *_fbe*5.520
				}
			}
			if col.HiddenAttr != nil && *col.HiddenAttr {
				width = 0
			}
			minAttr := int(col.MinAttr - 1)
			maxAttr := int(col.MaxAttr - 1)
			col_width_ranges = append(
				col_width_ranges,
				colWidthRange{
					minAttr: minAttr,
					maxAttr: maxAttr,
					width:   width,
					style:   context.getStyle(col.StyleAttr)})
		}
	}
	idx := 0
	for i := 0; i <= context._face; i++ {
		var width float64
		var style *style
		if idx >= len(col_width_ranges) {
			width = 65
		} else {
			col := col_width_ranges[idx]
			if i >= col.minAttr && i <= col.maxAttr {
				width = col.width
				style = col.style
				if i == col.maxAttr {
					idx++
				}
			} else {
				width = 65
			}
		}
		if context.scale != 100 {
			width = width * ((float64(context.scale)) / 100)
		}
		colInfos = append(colInfos, &colInfo{width: width, style: style})
	}
	context.colInfo = colInfos
}

type colWidthRange struct {
	minAttr int
	maxAttr int
	width   float64
	style   *style
}

func (_ebd *convertContext) distributeAnchors() {
	for _, _fddd := range _ebd._abba {
		_dec, _dcd := _fddd._ebfd, _fddd._adc
		_gbaf, _fcf := _fddd._afga, _fddd._aeaa
		_fce, _cca := _fddd._cfdc, _fddd._fgad
		_geb, _fac := _fddd._eacd, _fddd._gag
		var _gbb, _aeb, _cedg, _eabf *page
		for _, _eba := range _ebd.pages_span {
			for _, _dbdc := range _eba._eed {
				if _dec >= _dbdc._ebdd._aebb && _dec < _dbdc._ebdd._fbee {
					if _gbaf >= _dbdc._cbf._fcea && _gbaf < _dbdc._cbf._gced {
						_dbdc._ebbf = true
						_gbb = _dbdc
					}
					if _geb >= _dbdc._cbf._fcea && _geb < _dbdc._cbf._gced {
						_dbdc._ebbf = true
						_aeb = _dbdc
					}
				}
				if _fce >= _dbdc._ebdd._aebb && _fce < _dbdc._ebdd._fbee {
					if _gbaf >= _dbdc._cbf._fcea && _gbaf < _dbdc._cbf._gced {
						_dbdc._ebbf = true
						_eabf = _dbdc
					}
					if _geb >= _dbdc._cbf._fcea && _geb < _dbdc._cbf._gced {
						_dbdc._ebbf = true
						_cedg = _dbdc
					}
				}
			}
		}
		_bbcbf := _gbb != _aeb
		_fbb := _gbb != _eabf
		if _bbcbf && _fbb {
			_bdgf := _ebd.colInfo[_gbaf].start + _ae.FromEMU(_fcf)
			_fcd := _gbb._cbf._ebed
			_ccb := _ebd.colInfo[_geb].start + _ae.FromEMU(_fac)
			_beae := _ebd.rowInfo[_dec]._fffd + _ae.FromEMU(_dcd)
			_dac := float64(_gbb._ebdd._afg)
			_agba := _ebd.rowInfo[_fce]._fffd + _ae.FromEMU(_cca)
			_fda := _ccb + _fcd - _bdgf
			_gegf := _agba + _dac - _beae
			_aca := _ebd.imageFromAnchor(_fddd, _fda, _gegf)
			_gbb._agc = append(_gbb._agc, _ebd.getImage(_aca, _gegf, _fda, _bdgf, _beae, _fcd-_bdgf, _dac-_beae, _gb.ImgPart_lt))
			_aeb._agc = append(_aeb._agc, _ebd.getImage(_aca, _gegf, _fda, 0, _beae, _fcd-_bdgf, _dac-_beae, _gb.ImgPart_rt))
			_eabf._agc = append(_eabf._agc, _ebd.getImage(_aca, _gegf, _fda, _bdgf, 0, _fcd-_bdgf, _dac-_beae, _gb.ImgPart_lb))
			_cedg._agc = append(_cedg._agc, _ebd.getImage(_aca, _gegf, _fda, 0, 0, _fcd-_bdgf, _dac-_beae, _gb.ImgPart_rb))
		} else if _bbcbf {
			_gfg := _ebd.rowInfo[_dec]._fffd + _ae.FromEMU(_dcd)
			_bbd := _ebd.rowInfo[_fce]._fffd + _ae.FromEMU(_cca)
			_cfb := _ebd.colInfo[_gbaf].start + _ae.FromEMU(_fcf)
			_gfa := _gbb._cbf._ebed
			_gff := _ebd.colInfo[_geb].start + _ae.FromEMU(_fac)
			_dfd := _gff + _gfa - _cfb
			_bab := _bbd - _gfg
			_fcec := _ebd.imageFromAnchor(_fddd, _dfd, _bab)
			_gbb._agc = append(_gbb._agc, _ebd.getImage(_fcec, _bab, _dfd, _cfb, _gfg, _gfa-_cfb, 0, _gb.ImgPart_l))
			_aeb._agc = append(_aeb._agc, _ebd.getImage(_fcec, _bab, _dfd, 0, _gfg, _gfa-_cfb, 0, _gb.ImgPart_r))
		} else if _fbb {
			_gdb := _ebd.colInfo[_gbaf].start + _ae.FromEMU(_fcf)
			_ddf := _ebd.colInfo[_geb].start + _ae.FromEMU(_fac)
			_cdfc := _ebd.rowInfo[_dec]._fffd + _ae.FromEMU(_dcd)
			_cecg := float64(_gbb._ebdd._afg)
			_acaf := _ebd.rowInfo[_fce]._fffd + _ae.FromEMU(_cca)
			_fed := _ddf - _gdb
			_dbe := _acaf + _cecg - _cdfc
			_dgg := _ebd.imageFromAnchor(_fddd, _fed, _dbe)
			_gbb._agc = append(_gbb._agc, _ebd.getImage(_dgg, _dbe, _fed, _gdb, _cdfc, 0, _cecg-_cdfc, _gb.ImgPart_t))
			_eabf._agc = append(_eabf._agc, _ebd.getImage(_dgg, _dbe, _fed, _gdb, 0, 0, _cecg-_cdfc, _gb.ImgPart_b))
		} else {
			_fbg := _ebd.colInfo[_gbaf].start + _ae.FromEMU(_fcf)
			_agg := _ebd.colInfo[_geb].start + _ae.FromEMU(_fac)
			_dcf := _ebd.rowInfo[_dec]._fffd + _ae.FromEMU(_dcd)
			_gee := _ebd.rowInfo[_fce]._fffd + _ae.FromEMU(_cca)
			_aed := _agg - _fbg
			_abf := _gee - _dcf
			_bcg := _ebd.imageFromAnchor(_fddd, _aed, _abf)
			_gbb._agc = append(_gbb._agc, _ebd.getImage(_bcg, _abf, _aed, _fbg, _dcf, 0, 0, _gb.ImgPart_whole))
		}
	}
}

type pageRow struct {
	_cdff int
	_cbdb []*cell
}
type mergedCell struct {
	startRow    uint32
	startColumn uint32
	endRow      uint32
	endColumn   uint32
	width       float64
	height      float64
}

func (ctx *convertContext) makeMergedCells() {
	mergedCells := []*mergedCell{}
	for _, mergedcell := range ctx.sheet.MergedCells() {
		top, end, err := _f.ParseRangeReference(mergedcell.Reference())
		if err != nil {
			_ce.Log.Debug("\u0065\u0072r\u006f\u0072\u0020\u0070\u0061\u0072\u0073\u0069\u006e\u0067\u0020\u006d\u0065\u0072\u0067\u0065\u0064\u0020\u0063\u0065\u006c\u006c: \u0025\u0073", err)
			continue
		}
		merged_cell := mergedCell{startRow: top.RowIdx, startColumn: top.ColumnIdx, endRow: end.RowIdx, endColumn: end.ColumnIdx}
		for _gf := merged_cell.startColumn; _gf <= merged_cell.endColumn; _gf++ {
			merged_cell.width += ctx.colInfo[_gf].width
		}
		for i := merged_cell.startRow - 1; i < merged_cell.endRow; i++ {
			merged_cell.height += ctx.rowInfo[i].height
		}
		mergedCells = append(mergedCells, &merged_cell)
	}
	ctx.mergedCells = mergedCells
}

type symbol struct {
	value  string
	left   float64
	_fgfag float64
	_fceca float64
	_fae   *creator.TextStyle
	_cbeg  string
}
type rowspan struct {
	_afg  float64
	_aebb int
	_fbee int
}

// FontStyle represents a kind of font styling. It can be FontStyle_Regular, FontStyle_Bold, FontStyle_Italic and FontStyle_BoldItalic.
type FontStyle = _gb.FontStyle

func _fcc(_ffd []*symbol) float64 {
	_ebf := 0.0
	for _, _fec := range _ffd {
		_ebf += _fec._fceca
	}
	return _ebf
}

var _egac = map[uint32]creator.PageSize{1: creator.PageSize{8.5 * _ae.Inch, 11 * _ae.Inch}, 2: creator.PageSize{8.5 * _ae.Inch, 11 * _ae.Inch}, 3: creator.PageSize{11 * _ae.Inch, 17 * _ae.Inch}, 4: creator.PageSize{17 * _ae.Inch, 11 * _ae.Inch}, 5: creator.PageSize{8.5 * _ae.Inch, 14 * _ae.Inch}, 6: creator.PageSize{5.5 * _ae.Inch, 8.5 * _ae.Inch}, 7: creator.PageSize{7.5 * _ae.Inch, 10 * _ae.Inch}, 8: creator.PageSize{_gefg(297), _gefg(420)}, 9: creator.PageSize{_gefg(210), _gefg(297)}, 10: creator.PageSize{_gefg(210), _gefg(297)}, 11: creator.PageSize{_gefg(148), _gefg(210)}, 70: creator.PageSize{_gefg(105), _gefg(148)}, 12: creator.PageSize{_gefg(250), _gefg(354)}, 13: creator.PageSize{_gefg(182), _gefg(257)}, 14: creator.PageSize{8.5 * _ae.Inch, 13 * _ae.Inch}, 20: creator.PageSize{4.125 * _ae.Inch, 9.5 * _ae.Inch}, 27: creator.PageSize{_gefg(110), _gefg(220)}, 28: creator.PageSize{_gefg(162), _gefg(229)}, 34: creator.PageSize{_gefg(250), _gefg(176)}, 29: creator.PageSize{_gefg(324), _gefg(458)}, 30: creator.PageSize{_gefg(229), _gefg(324)}, 31: creator.PageSize{_gefg(114), _gefg(162)}, 37: creator.PageSize{3.88 * _ae.Inch, 7.5 * _ae.Inch}, 43: creator.PageSize{_gefg(100), _gefg(148)}, 69: creator.PageSize{_gefg(200), _gefg(148)}}

type line struct {
	lineSpace float64
	symbols   []*symbol
	fontSize  float64
}

const _ac = 0.25
const _bd = 2

var border_base = _gefg(0.0625)

func (ctx *convertContext) makePages() {
	for _, _dbb := range ctx.pages_span {
		for _, _ddg := range ctx._gbg {
			_dbb._eed = append(_dbb._eed, &page{_bdda: []*pageRow{}, _cbf: _dbb, _ebdd: _ddg})
		}
	}
}
func _becc(_ece *bool) bool { return _ece != nil && *_ece }
func (_bbeg *convertContext) getImage(_cdc _a.Image, _aadd, _fefg, _cebd, _dga, _gdfe, _adag float64, _aeef _gb.ImgPart) *creator.Image {
	_dga += _bbeg._bcag
	_cebd += _bbeg._ebae
	_cdcb, _edg := _gb.GetImage(_bbeg.creator, _cdc, _aadd, _fefg, _cebd, _dga, _gdfe, _adag, _aeef)
	if _edg != nil {
		_ce.Log.Debug("\u0043\u0061\u006eno\u0074\u0020\u0067\u0065\u0074\u0020\u0061\u006e\u0020\u0069\u006d\u0061\u0067\u0065\u003a\u0020\u0025\u0073", _edg)
		return nil
	}
	return _cdcb
}
func (_ebeg *convertContext) getStyle(_dfe *uint32) *style {
	_aagfa := &style{}
	_fdg := false
	if _dfe != nil {
		_cee := _ebeg._aeg.GetCellStyle(*_dfe)
		_faga := _cee.GetFont()
		for _, _gbc := range _faga.Name {
			if _gbc != nil {
				_aagfa.fontName = &_gbc.ValAttr
				_fdg = true
				break
			}
		}
		for _, _gfee := range _faga.B {
			if _gfee != nil {
				_dcb := _gfee.ValAttr == nil || *_gfee.ValAttr
				_aagfa.bold = &_dcb
				_fdg = true
				break
			}
		}
		for _, _cecb := range _faga.I {
			if _cecb != nil {
				_beef := _cecb.ValAttr == nil || *_cecb.ValAttr
				_aagfa._gdgae = &_beef
				_fdg = true
				break
			}
		}
		for _, _cbecd := range _faga.U {
			if _cbecd != nil {
				_add := _cbecd.ValAttr == sml.ST_UnderlineValuesSingle || _cbecd.ValAttr == sml.ST_UnderlineValuesUnset
				_aagfa._bcd = &_add
				_fdg = true
				break
			}
		}
		for _, _cfcf := range _faga.Sz {
			if _cfcf != nil {
				_gfcg := _cfcf.ValAttr / 12 * _gb.DefaultFontSize
				_aagfa.fontSize = &_gfcg
				_fdg = true
				break
			}
		}
		for _, _caag := range _faga.VertAlign {
			if _caag != nil {
				_ffag := _caag.ValAttr == _bb.ST_VerticalAlignRunSuperscript
				_aagfa._ffab = &_ffag
				_fede := _caag.ValAttr == _bb.ST_VerticalAlignRunSubscript
				_aagfa._eac = &_fede
				_fdg = true
				break
			}
		}
		for _, _ccdf := range _faga.Color {
			if _ccdf != nil {
				_aagfa._gege = _ebeg.getColorStringFromSmlColor(_ccdf)
				_fdg = true
				break
			}
		}
		_fdgd := _cee.GetBorder()
		if _fdgd.Top != nil {
			_aagfa.border1 = _ebeg.getBorder(_fdgd.Top)
			_fdg = true
		}
		if _fdgd.Bottom != nil {
			_aagfa.border2 = _ebeg.getBorder(_fdgd.Bottom)
			_fdg = true
		}
		if _fdgd.Left != nil {
			_aagfa.border3 = _ebeg.getBorder(_fdgd.Left)
			_fdg = true
		}
		if _fdgd.Right != nil {
			_aagfa.border4 = _ebeg.getBorder(_fdgd.Right)
			_fdg = true
		}
		if _cee.Wrapped() {
			_aagfa._gea = true
			_fdg = true
		}
		if _dcg := _cee.GetVerticalAlignment(); _dcg != sml.ST_VerticalAlignmentUnset {
			_aagfa._egd = _dcg
			_fdg = true
		}
		if _adfg := _cee.GetHorizontalAlignment(); _adfg != sml.ST_HorizontalAlignmentUnset {
			_aagfa._cbef = _adfg
			_fdg = true
		}
	}
	if _fdg {
		return _aagfa
	}
	return nil
}

const _ga = 0.64

func (_cbee *convertContext) getBorder(ct_border *sml.CT_BorderPr) *border {
	border := &border{}
	switch ct_border.StyleAttr {
	case sml.ST_BorderStyleThin:
		border.thickness = border_base
	case sml.ST_BorderStyleMedium:
		border.thickness = border_base * 2
	case sml.ST_BorderStyleThick:
		border.thickness = border_base * 4
	}
	if border.thickness == 0.0 {
		return nil
	}
	if _edbf := ct_border.Color; _edbf != nil {
		_abfa := _cbee.getColorStringFromSmlColor(_edbf)
		if _abfa != nil {
			border.color = creator.ColorRGBFromHex(*_abfa)
		} else {
			border.color = creator.ColorBlack
		}
	} else {
		border.color = creator.ColorBlack
	}
	return border
}

type rowInfo struct {
	_fffd       float64
	_bgc        bool
	height      float64
	style       *style
	cells       []*cell
	borderWidth float64
}
type convertContext struct {
	creator     *creator.Creator
	workbook    *_d.Workbook
	_ggce       *_be.Theme
	sheet       *_d.Sheet
	_aeg        *_d.StyleSheet
	_dbc        int
	_face       int
	pages_span  []*pagespan
	_dfdc       *page
	colInfo     []*colInfo
	rowInfo     []*rowInfo
	_gbg        []*rowspan
	_bcag       float64
	_ebae       float64
	pageHeight  float64
	pageWidth   float64
	mergedCells []*mergedCell
	_abba       []*anchor
	scale       uint32
}

func _adce(_bdac, _addc *style) {
	if _addc == nil {
		return
	}
	if _bdac == nil {
		_bdac = _addc
		return
	}
	if _bdac.fontName == nil {
		_bdac.fontName = _addc.fontName
	}
	if _bdac._gege == nil {
		_bdac._gege = _addc._gege
	}
	if _bdac.fontSize == nil {
		_bdac.fontSize = _addc.fontSize
	}
	if _bdac.bold == nil {
		_bdac.bold = _addc.bold
	}
	if _bdac._gdgae == nil {
		_bdac._gdgae = _addc._gdgae
	}
	if _bdac._bcd == nil {
		_bdac._bcd = _addc._bcd
	}
	if _bdac._ffab == nil {
		_bdac._ffab = _addc._ffab
	}
	if _bdac._eac == nil {
		_bdac._eac = _addc._eac
	}
	if _bdac.border1 == nil {
		_bdac.border1 = _addc.border1
	}
	if _bdac.border2 == nil {
		_bdac.border2 = _addc.border2
	}
	if _bdac.border3 == nil {
		_bdac.border3 = _addc.border3
	}
	if _bdac.border4 == nil {
		_bdac.border4 = _addc.border4
	}
	if _bdac._egd == sml.ST_VerticalAlignmentUnset {
		_bdac._egd = _addc._egd
	}
	if _bdac._cbef == sml.ST_HorizontalAlignmentUnset {
		_bdac._cbef = _addc._cbef
	}
}
func (_fgce *convertContext) drawSheet() {
	for _cbgg, _eebe := range _fgce.pages_span {
		_ddaf := len(_eebe._eed)
		if _cbgg == len(_fgce.pages_span)-1 {
			for _dgde := len(_eebe._eed) - 1; _dgde >= 0; _dgde-- {
				if !_eebe._eed[_dgde]._ebbf {
					_ddaf = _dgde
				}
			}
		}
		pages := _eebe._eed[:_ddaf]
		for _, _bbee := range pages {
			_fgce.creator.NewPage()
			_fgce.drawPage(_bbee)
		}
	}
}
func (ctx *convertContext) makePagespans() {
	ctx.pages_span = []*pagespan{}
	_bbb := 0.0
	_eeb := 0
	for idx, colInfo := range ctx.colInfo {
		width := colInfo.width
		if _bbb+width <= ctx.pageWidth {
			colInfo.start = _bbb
			_bbb += width
		} else {
			colInfo.start = 0
			ctx.pages_span = append(ctx.pages_span, &pagespan{_ebed: _bbb, _fcea: _eeb, _gced: idx})
			_bbb = width
			_eeb = idx
		}
	}
	ctx.pages_span = append(ctx.pages_span, &pagespan{_ebed: _bbb, _fcea: _eeb, _gced: len(ctx.colInfo)})
}

type pagespan struct {
	_ebed float64
	_eed  []*page
	_fcea int
	_gced int
}
type colInfo struct {
	start float64
	width float64
	style *style
}

func (_gacg *convertContext) alignSymbolsHorizontally(_bfd *cell, _cab sml.ST_HorizontalAlignment) {
	if _cab == sml.ST_HorizontalAlignmentUnset || _cab == sml.ST_HorizontalAlignmentGeneral {
		switch _bfd.cellType {
		case sml.ST_CellTypeB:
			_cab = sml.ST_HorizontalAlignmentCenter
		case sml.ST_CellTypeN:
			_cab = sml.ST_HorizontalAlignmentRight
		default:
			_cab = sml.ST_HorizontalAlignmentLeft
		}
	}
	var _ggd float64
	for _, line := range _bfd.lines {
		switch _cab {
		case sml.ST_HorizontalAlignmentLeft:
			//_ggd = _eb
			_ggd = 0
		case sml.ST_HorizontalAlignmentRight:
			_ebeb := _fcc(line.symbols)
			_ggd = _bfd._bde - _eb - _ebeb
		case sml.ST_HorizontalAlignmentCenter:
			_eccg := _fcc(line.symbols)
			_ggd = (_bfd._bde - _eccg) / 2
		}
		for _, _acd := range line.symbols {
			_acd.left += _ggd
		}
	}
}
func (_aaaf *convertContext) getColorFromTheme(_fcg uint32) string {
	_bddag := _aaaf.workbook.Themes()
	if len(_bddag) != 0 {
		_affg := _bddag[0]
		if _dde := _affg.ThemeElements; _dde != nil {
			if _dcef := _dde.ClrScheme; _dcef != nil {
				switch _fcg {
				case 0:
					return _gb.GetColorStringFromDmlColor(_dcef.Lt1)
				case 1:
					return _gb.GetColorStringFromDmlColor(_dcef.Dk1)
				case 2:
					return _gb.GetColorStringFromDmlColor(_dcef.Lt2)
				case 3:
					return _gb.GetColorStringFromDmlColor(_dcef.Dk2)
				case 4:
					return _gb.GetColorStringFromDmlColor(_dcef.Accent1)
				case 5:
					return _gb.GetColorStringFromDmlColor(_dcef.Accent2)
				case 6:
					return _gb.GetColorStringFromDmlColor(_dcef.Accent3)
				case 7:
					return _gb.GetColorStringFromDmlColor(_dcef.Accent4)
				case 8:
					return _gb.GetColorStringFromDmlColor(_dcef.Accent5)
				case 9:
					return _gb.GetColorStringFromDmlColor(_dcef.Accent6)
				}
			}
		}
	}
	return ""
}

var _ec = _gefg(1)

func (_dbed *convertContext) getStyleFromCell(_abbad _d.Cell, _abe, _cfe *style) *style {
	_gebff := _abbad.X()
	_gdfb := _dbed.getStyle(_gebff.SAttr)
	_adce(_gdfb, _abe)
	_adce(_gdfb, _cfe)
	return _gdfb
}
func (_fddf *convertContext) getSymbolsFromR(_agdd []*sml.CT_RElt, _gdga *style) []*symbol {
	symbols := []*symbol{}
	for _, _fcfd := range _agdd {
		//style := _fddf.combineCellStyleWithRPrElt(_gdga, _fcfd.RPr)
		style := _fddf.getStyleFromRPrElt(_fcfd.RPr)
		for _, _bdf := range _fcfd.T {
			symbols = append(symbols, &symbol{value: string(_bdf), _fae: _fddf.makeTextStyleFromCellStyle(style)})
		}
	}
	return symbols
}

var _ca = 3.025 / _gefg(1)
var _ceag = []string{"\u0030\u0030\u0030\u0030\u0030\u0030", "\u0066\u0066\u0066\u0066\u0066\u0066", "\u0066\u0066\u0030\u0030\u0030\u0030", "\u0030\u0030\u0066\u0066\u0030\u0030", "\u0030\u0030\u0030\u0030\u0066\u0066", "\u0066\u0066\u0066\u0066\u0030\u0030", "\u0066\u0066\u0030\u0030\u0066\u0066", "\u0030\u0030\u0066\u0066\u0066\u0066", "\u0030\u0030\u0030\u0030\u0030\u0030", "\u0066\u0066\u0066\u0066\u0066\u0066", "\u0066\u0066\u0030\u0030\u0030\u0030", "\u0030\u0030\u0066\u0066\u0030\u0030", "\u0030\u0030\u0030\u0030\u0066\u0066", "\u0066\u0066\u0066\u0066\u0030\u0030", "\u0066\u0066\u0030\u0030\u0066\u0066", "\u0030\u0030\u0066\u0066\u0066\u0066", "\u0038\u0030\u0030\u0030\u0030\u0030", "\u0030\u0030\u0038\u0030\u0030\u0030", "\u0030\u0030\u0030\u0030\u0038\u0030", "\u0038\u0030\u0038\u0030\u0030\u0030", "\u0038\u0030\u0030\u0030\u0038\u0030", "\u0030\u0030\u0038\u0030\u0038\u0030", "\u0063\u0030\u0063\u0030\u0063\u0030", "\u0038\u0030\u0038\u0030\u0038\u0030", "\u0039\u0039\u0039\u0039\u0066\u0066", "\u0039\u0039\u0033\u0033\u0036\u0036", "\u0066\u0066\u0066\u0066\u0063\u0063", "\u0063\u0063\u0066\u0066\u0066\u0066", "\u0036\u0036\u0030\u0030\u0036\u0036", "\u0066\u0066\u0038\u0030\u0038\u0030", "\u0030\u0030\u0036\u0036\u0063\u0063", "\u0063\u0063\u0063\u0063\u0066\u0066", "\u0030\u0030\u0030\u0030\u0038\u0030", "\u0066\u0066\u0030\u0030\u0066\u0066", "\u0066\u0066\u0066\u0066\u0030\u0030", "\u0030\u0030\u0066\u0066\u0066\u0066", "\u0038\u0030\u0030\u0030\u0038\u0030", "\u0038\u0030\u0030\u0030\u0030\u0030", "\u0030\u0030\u0038\u0030\u0038\u0030", "\u0030\u0030\u0030\u0030\u0066\u0066", "\u0030\u0030\u0063\u0063\u0066\u0066", "\u0063\u0063\u0066\u0066\u0066\u0066", "\u0063\u0063\u0066\u0066\u0063\u0063", "\u0066\u0066\u0066\u0066\u0039\u0039", "\u0039\u0039\u0063\u0063\u0066\u0066", "\u0066\u0066\u0039\u0039\u0063\u0063", "\u0063\u0063\u0039\u0039\u0066\u0066", "\u0066\u0066\u0063\u0063\u0039\u0039", "\u0033\u0033\u0036\u0036\u0066\u0066", "\u0033\u0033\u0063\u0063\u0063\u0063", "\u0039\u0039\u0063\u0063\u0030\u0030", "\u0066\u0066\u0063\u0063\u0030\u0030", "\u0066\u0066\u0039\u0039\u0030\u0030", "\u0066\u0066\u0036\u0036\u0030\u0030", "\u0036\u0036\u0036\u0036\u0039\u0039", "\u0039\u0036\u0039\u0036\u0039\u0036", "\u0030\u0030\u0033\u0033\u0036\u0036", "\u0033\u0033\u0039\u0039\u0036\u0036", "\u0030\u0030\u0033\u0033\u0030\u0030", "\u0033\u0033\u0033\u0033\u0030\u0030", "\u0039\u0039\u0033\u0033\u0030\u0030", "\u0039\u0039\u0033\u0033\u0036\u0036", "\u0033\u0033\u0033\u0033\u0039\u0039", "\u0033\u0033\u0033\u0033\u0033\u0033"}

func (_ecgd *convertContext) alignSymbolsVertically(_aac *cell, _dff sml.ST_VerticalAlignment) {
	var _bgae float64
	switch _dff {
	case sml.ST_VerticalAlignmentTop:
		_bgae = _bd
		if _aac._cccb {
			_bgae -= _bf
		} else if _aac._aaee {
			_bgae += 4 * _bf
		}
		for _, _gfe := range _aac.lines {
			_bgae += _gfe.fontSize
			_gfe.lineSpace = _bgae
			_bgae += _ec
		}
	case sml.ST_VerticalAlignmentCenter:
		_ebg := 0.0
		for _, _eea := range _aac.lines {
			_ebg += _eea.fontSize + _gefg(1)
		}
		_bgae = 0.5 * (_aac.rowHeight - _ebg)
		if _aac._cccb {
			_bgae -= 2 * _bf
		} else if _aac._aaee {
			_bgae += 2 * _bf
		}
		for _, _feed := range _aac.lines {
			_bgae += _feed.fontSize + 0.5*_ec
			_feed.lineSpace = _bgae
			_bgae += 0.5 * _ec
		}
	default:
		_bgae = _aac.rowHeight - _bd
		if _aac._cccb {
			_bgae -= 4 * _bf
		} else if _aac._aaee {
			_bgae += _bf
		}
		for _ecc := len(_aac.lines) - 1; _ecc >= 0; _ecc-- {
			_aac.lines[_ecc].lineSpace = _bgae
			_bgae -= _aac.lines[_ecc].fontSize
			_bgae -= _ec
		}
	}
}
func (_feg *convertContext) combineCellStyleWithRPrElt(style1 *style, cellStyle *sml.CT_RPrElt) *style {
	style := *style1
	styleFromCell := _feg.getStyleFromRPrElt(cellStyle)
	if styleFromCell == nil {
		return &style
	}
	if styleFromCell._gege != nil {
		style._gege = styleFromCell._gege
	}
	if styleFromCell.fontSize != nil {
		style.fontSize = styleFromCell.fontSize
	}
	if styleFromCell.fontName != nil {
		style.fontName = styleFromCell.fontName
	}
	if styleFromCell.bold != nil {
		style.bold = styleFromCell.bold
	}
	if styleFromCell._gdgae != nil {
		style._gdgae = styleFromCell._gdgae
	}
	if styleFromCell._bcd != nil {
		style._bcd = styleFromCell._bcd
	}
	if styleFromCell._ffab != nil {
		style._ffab = styleFromCell._ffab
	}
	if styleFromCell._eac != nil {
		style._eac = styleFromCell._eac
	}
	return &style
}

const (
	FontStyle_Regular    FontStyle = 0
	FontStyle_Bold       FontStyle = 1
	FontStyle_Italic     FontStyle = 2
	FontStyle_BoldItalic FontStyle = 3
)

func (_bg *convertContext) makeAnchors() {
	_ag, _cg := _bg.sheet.GetDrawing()
	if _ag != nil {
		for _, _gac := range _ag.EG_Anchor {
			_dbf := &anchor{}
			if _ee := _gac.TwoCellAnchor; _ee != nil {
				_de, _bgf := _ee.From, _ee.To
				if _de == nil || _bgf == nil {
					return
				}
				_dbf._ebfd = int(_de.Row)
				_dbf._adc = _gb.FromSTCoordinate(_de.RowOff)
				_dbf._afga = int(_de.Col)
				_dbf._aeaa = _gb.FromSTCoordinate(_de.ColOff)
				_dbf._cfdc = int(_bgf.Row)
				_dbf._fgad = _gb.FromSTCoordinate(_bgf.RowOff)
				_dbf._eacd = int(_bgf.Col)
				_dbf._gag = _gb.FromSTCoordinate(_bgf.ColOff)
				if _gd := _ee.Choice; _gd != nil {
					if _fa := _gd.Pic; _fa != nil {
						if _bbc := _fa.BlipFill; _bbc != nil {
							if _dd := _bbc.Blip; _dd != nil {
								if _fe := _dd.EmbedAttr; _fe != nil {
									for _, _fef := range _cg.X().Relationship {
										if _fef.IdAttr == *_fe {
											for _, _ace := range _bg.workbook.Images {
												if _ace.Target() == _fef.TargetAttr {
													_gg, _fc := _ba.Open(_ace.Path())
													if _fc != nil {
														_ce.Log.Debug("\u004fp\u0065\u006e\u0020\u0069m\u0061\u0067\u0065\u0020\u0066i\u006ce\u0020e\u0072\u0072\u006f\u0072\u003a\u0020\u0025s", _fc)
														continue
													}
													_bdg, _, _fc := _a.Decode(_gg)
													if _fc != nil {
														_ce.Log.Debug("\u0044\u0065\u0063\u006fde\u0020\u0069\u006d\u0061\u0067\u0065\u0020\u0065\u0072\u0072\u006f\u0072\u003a\u0020%\u0073", _fc)
														continue
													}
													_dbf._cgaf = _bdg
												}
											}
										}
									}
								}
							}
						}
					} else if _ggb := _gd.GraphicFrame; _ggb != nil {
						if _dgd := _ggb.Graphic; _dgd != nil {
							if _eef := _dgd.GraphicData; _eef != nil {
								for _, _cgf := range _eef.Any {
									if _fca, _cfa := _cgf.(*_cf.Chart); _cfa {
										for _, _ebe := range _cg.X().Relationship {
											if _ebe.IdAttr == _fca.IdAttr {
												_acg := _bg.workbook.GetChartByTargetId(_ebe.TargetAttr)
												if _acg != nil {
													_dbf._fagc = _acg
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
			if _dbf._cgaf != nil || _dbf._fagc != nil {
				_bg._abba = append(_bg._abba, _dbf)
			}
		}
	}
}
func (_def *convertContext) getColorStringFromSmlColor(_bec *sml.CT_Color) *string {
	var _bbcbg string
	if _bec.RgbAttr != nil {
		_bbcbg = *_bec.RgbAttr
	} else if _bec.IndexedAttr != nil && *_bec.IndexedAttr < 64 {
		_bbcbg = _ceag[*_bec.IndexedAttr]
	} else if _bec.ThemeAttr != nil {
		_egbc := *_bec.ThemeAttr
		_bbcbg = _def.getColorFromTheme(_egbc)
	}
	if _bbcbg == "" {
		return nil
	}
	if len(_bbcbg) > 6 {
		_bbcbg = _bbcbg[(len(_bbcbg) - 6):]
	}
	if _bec.TintAttr != nil {
		_cfeg := *_bec.TintAttr
		_bbcbg = _gb.AdjustColorByTint(_bbcbg, _cfeg)
	}
	_bbcbg = "\u0023" + _bbcbg
	return &_bbcbg
}

// RegisterFont makes a PdfFont accessible for using in converting to PDF.
func RegisterFont(name string, style FontStyle, font *model.PdfFont) {
	_gb.RegisterFont(name, style, font)
}

func RegisterFontFromFile(name string, style FontStyle, file string) {
	font, err := model.NewCompositePdfFontFromTTFFile(file)
	if err != nil {
		return
	}
	_gb.RegisterFont(name, style, font)
}

const _bf = 1.5

func (context *convertContext) makeCells() {
	sheet := context.sheet
	sheetRows := sheet.Rows()
	rowIdx := 0
	for _, row := range context.rowInfo {
		row.cells = []*cell{}
		max_thickness := 0.0
		rowStyle := row.style
		if row._bgc {
			sheetRow := sheetRows[rowIdx]
			rowIdx++
			rowHeight := row.height
			if row.height <= 0 {
				continue
			}
			for _, scell := range sheetRow.Cells() {
				cellref, err := _f.ParseCellReference(scell.Reference())
				if err != nil {
					_ce.Log.Debug("\u0043\u0061\u006e\u006eo\u0074\u0020\u0070\u0061\u0072\u0073\u0065\u0020\u0061\u0020r\u0065f\u0065\u0072\u0065\u006e\u0063\u0065\u003a \u0025\u0073", err)
					continue
				}
				colinfo := context.colInfo[cellref.ColumnIdx]
				columnWidth := colinfo.width
				_baf := columnWidth
				_dc := colinfo.style
				var _bdd, _fcb, _bda, _dea bool
				for _, mergedcell := range context.mergedCells {
					if cellref.RowIdx >= mergedcell.startRow &&
						cellref.RowIdx <= mergedcell.endRow &&
						cellref.ColumnIdx >= mergedcell.startColumn &&
						cellref.ColumnIdx <= mergedcell.endColumn {

						if cellref.ColumnIdx == mergedcell.startColumn && cellref.RowIdx == mergedcell.startRow {
							columnWidth = mergedcell.width
							rowHeight = mergedcell.height
						}
						_bdd = cellref.RowIdx != mergedcell.startRow
						_fcb = cellref.RowIdx != mergedcell.endRow
						_bda = cellref.ColumnIdx != mergedcell.startColumn
						_dea = cellref.ColumnIdx != mergedcell.endColumn
					}
				}
				style := context.getStyleFromCell(scell, rowStyle, _dc)
				var _ed, _eeff, _aad bool
				var border1, border2, border3, border4 *border
				var _eab sml.ST_VerticalAlignment
				var _ceg sml.ST_HorizontalAlignment

				if style != nil {
					if !_bdd {
						border1 = style.border1
					}
					if !_fcb {
						border2 = style.border2
					}
					if !_bda {
						border3 = style.border3
					}
					if !_dea {
						border4 = style.border4
					}
					if border2 != nil && border2.thickness > max_thickness {
						max_thickness = border2.thickness
					}
					_eab = style._egd
					_ceg = style._cbef
					if style._ffab != nil {
						_ed = *style._ffab
					}
					if style._eac != nil {
						_eeff = *style._eac
					}
					_aad = style._gea
				}
				lines, cellType := context.getContentFromCell(scell, style, columnWidth, _aad)

				_affd := &cell{
					cellType: cellType, _bde: columnWidth, _bcf: _baf, rowHeight: rowHeight,
					lines:   lines,
					border1: border1,
					border2: border2,
					border3: border3,
					border4: border4,
					_cccb:   _ed, _aaee: _eeff}
				context.alignSymbolsHorizontally(_affd, _ceg)
				context.alignSymbolsVertically(_affd, _eab)
				row.cells = append(row.cells, _affd)
			}
		}
		row.borderWidth = max_thickness
	}
}

type anchor struct {
	_cgaf _a.Image
	_fagc *_cf.ChartSpace
	_ebfd int
	_adc  int64
	_afga int
	_aeaa int64
	_cfdc int
	_fgad int64
	_eacd int
	_gag  int64
}

func (ctx *convertContext) makeTextStyleFromCellStyle(_style *style) *creator.TextStyle {
	textstyle := ctx.creator.NewTextStyle()

	if _style == nil {
		textstyle.FontSize = _gb.DefaultFontSize
		textstyle.Font = _gb.AssignStdFontByName(textstyle, _gb.StdFontsMap["default"][FontStyle_Regular])
		return &textstyle
	}
	if _becc(_style._bcd) {
		textstyle.Underline = true
		textstyle.UnderlineStyle = creator.TextDecorationLineStyle{Offset: 0.5, Thickness: _gefg(float64(1) / 32)}
	}
	var _aebc FontStyle
	if _becc(_style.bold) && _becc(_style._gdgae) {
		_aebc = FontStyle_BoldItalic
	} else if _becc(_style.bold) {
		_aebc = FontStyle_Bold
	} else if _becc(_style._gdgae) {
		_aebc = FontStyle_Italic
	} else {
		_aebc = FontStyle_Regular
	}
	_eaec := "default"
	if _style.fontName != nil {
		_eaec = *_style.fontName
	}
	delete(_gb.StdFontsMap, "Times New Roman")
	if _fbd, _ebea := _gb.StdFontsMap[_eaec]; _ebea {
		textstyle.Font = _gb.AssignStdFontByName(textstyle, _fbd[_aebc])
	} else if _cgdg := _gb.GetRegisteredFont(_eaec, _aebc); _cgdg != nil {
		textstyle.Font = _cgdg
	} else {
		_ce.Log.Debug("Font %s with style %s is not found, reset to default.", _eaec, _aebc)
		textstyle.Font = _gb.AssignStdFontByName(textstyle, _gb.StdFontsMap["default"][_aebc])
	}

	if _style.fontSize != nil {
		textstyle.FontSize = *_style.fontSize
	}
	if _style._gege != nil {
		textstyle.Color = creator.ColorRGBFromHex(*_style._gege)
	}
	if _style._ffab != nil && *_style._ffab {
		textstyle.FontSize *= _ga
	} else if _style._eac != nil && *_style._eac {
		textstyle.FontSize *= _ga
	}

	if ctx.scale != 100 {
		//textstyle.HorizontalScaling = 89
		//in somecase, decimal font size cause text render incorrectly
		textstyle.FontSize = math.Round(textstyle.FontSize * (float64(ctx.scale)) / 100)
	}

	return &textstyle
}
func _gefg(_ccdb float64) float64 { return _ccdb * _ae.Millimeter }
func (_fdda *convertContext) makeRowspans() {
	var _cbc float64
	_ff := 0
	for _cag, _bafe := range _fdda.rowInfo {
		_cbb := _bafe.height + _bafe.borderWidth
		if _cbc+_cbb <= _fdda.pageHeight {
			_bafe._fffd = _cbc
			_cbc += _cbb
		} else {
			_fdda._gbg = append(_fdda._gbg, &rowspan{_afg: _cbc, _aebb: _ff, _fbee: _cag})
			_ff = _cag
			_bafe._fffd = 0
			_cbc = _cbb
		}
	}
	_fdda._gbg = append(_fdda._gbg, &rowspan{_afg: _cbc, _aebb: _ff, _fbee: len(_fdda.rowInfo)})
}

type border struct {
	thickness float64
	color     creator.Color
}

func (_agf *convertContext) getSymbolsFromString(_gec string, _dadge *style) []*symbol {
	_eefb := []*symbol{}
	_cfca := _agf.makeTextStyleFromCellStyle(_dadge)
	for _, _ddgc := range _gec {
		_eefb = append(_eefb, &symbol{value: string(_ddgc), _fae: _cfca})
	}
	return _eefb
}
func (_aabf *convertContext) addRowToPage(_dgeg []*cell, _aede int) {
	_afa := 0.0
	_cfdg := _aabf.pageWidth
	for _, _ccc := range _dgeg {
		if len(_ccc.lines) != 0 {
			_ccc._bbbc = _afa
			_afa = _ccc._bffg + _ccc._bde
		}
	}
	for _dgdef := len(_dgeg) - 1; _dgdef >= 0; _dgdef-- {
		_gdbg := _dgeg[_dgdef]
		if len(_gdbg.lines) != 0 {
			_gdbg._gbge = _cfdg
			_cfdg = _gdbg._bffg
		}
	}
	_aabf._dfdc._bdda = append(_aabf._dfdc._bdda, &pageRow{_cdff: _aede, _cbdb: _dgeg})
}

const _eb = 3

func (ctx *convertContext) drawPage(_page *page) {
	_dgce := ctx._bcag
	_egc := ctx._ebae
	for _, _cdda := range _page._bdda {
		_aeab := ctx.rowInfo[_cdda._cdff]
		for _, _bee := range _cdda._cbdb {
			_abfd := _bee._bbbc < _bee._bffg
			_bbf := _bee._gbge > _bee._bffg+_bee._bde
			var _fadd, _fdf bool
			for _, _ebad := range _bee.lines {
				for _, _ddfa := range _ebad.symbols {
					if _abfd && !_fadd {
						_fadd = _ddfa.left < 0
					}
					if _bbf && !_fdf {
						_fdf = _bee._bde < _ddfa.left+_ddfa._fceca
					}
					if _bee._bffg+_ddfa.left >= _bee._bbbc && _bee._bffg+_ddfa.left+_ddfa._fceca <= _bee._gbge {
						_ege := ctx.creator.NewStyledParagraph()
						_gda := _egc + _bee._bffg + _ddfa.left
						_fff := _dgce + _aeab._fffd + _ebad.lineSpace - _ddfa._fgfag - _gefg(0.5)
						_ege.SetPos(_gda, _fff)
						var _ccg *creator.TextChunk
						if _ddfa._cbeg != "" {
							_ccg = _ege.AddExternalLink(_ddfa.value, _ddfa._cbeg)
						} else {
							_ccg = _ege.Append(_ddfa.value)
						}
						if _ddfa._fae != nil {
							_ccg.Style = *_ddfa._fae
						}
						ctx.creator.Draw(_ege)
					}
				}
			}
			var _afe, _aae, _gcg, _gcb, _dcc, _bbdc float64
			var _caa, _efa, _dcde, _bfb creator.Color
			if _adfe := _bee.border1; _adfe != nil {
				_afe = _adfe.thickness
				_caa = _adfe.color
			}
			if _gdf := _bee.border2; _gdf != nil {
				_aae = _gdf.thickness
				_efa = _gdf.color
			}
			if _fgg := _bee.border3; _fgg != nil {
				_gcg = _fgg.thickness
				_dcc = _gcg / 2
				_dcde = _fgg.color
			}
			if _gef := _bee.border4; _gef != nil {
				_gcb = _gef.thickness
				_bbdc = _gcb / 2
				_bfb = _gef.color
			}
			var _fba float64
			if _cdda._cdff > 1 {
				_fba = ctx.rowInfo[_cdda._cdff-1].borderWidth
			}
			_bgg := _dgce + _aeab._fffd - 0.5*(_fba-_afe)
			_dcdc := _dgce + _aeab._fffd + _aeab.height + 0.5*(_aeab.borderWidth+_aae)
			_beeb := _egc + _bee._bffg
			_febf := _beeb + _bee._bcf
			_gb.DrawLine(ctx.creator, _beeb, _bgg, _febf, _bgg, _afe, _caa)
			_gb.DrawLine(ctx.creator, _beeb, _dcdc, _febf, _dcdc, _aae, _efa)
			if !_fadd {
				_gb.DrawLine(ctx.creator, _beeb-_dcc, _bgg, _beeb-_dcc, _dcdc, _gcg, _dcde)
			}
			if !_fdf {
				_gb.DrawLine(ctx.creator, _febf-_bbdc, _bgg, _febf-_bbdc, _dcdc, _gcb, _bfb)
			}
		}
	}
	for _, _dae := range _page._agc {
		if _dae != nil {
			ctx.creator.Draw(_dae)
		}
	}
}

type page struct {
	_bdda []*pageRow
	_ebbf bool
	_agc  []*creator.Image
	_cbf  *pagespan
	_ebdd *rowspan
}
type cell struct {
	cellType  sml.ST_CellType
	_gdfg     int
	_bffg     float64
	lines     []*line
	_bde      float64
	_bcf      float64
	rowHeight float64
	_bbbc     float64
	_gbge     float64
	textStyle *creator.TextStyle
	border1   *border
	border2   *border
	border3   *border
	border4   *border
	_cccb     bool
	_aaee     bool
}

func findMax(_cef []*symbol) float64 {
	_accc := 0.0
	for _, _edb := range _cef {
		if _edb._fgfag > _accc {
			_accc = _edb._fgfag
		}
	}
	return _accc
}
func (ctx *convertContext) makeRows() {
	row_info := []*rowInfo{}
	rows := ctx.sheet.Rows()
	_fd := 0
	for _, row := range rows {
		_fd++
		rowNumber := int(row.RowNumber())
		if rowNumber > _fd {
			for _aabc := _fd; _aabc < rowNumber; _aabc++ {
				row_info = append(row_info, &rowInfo{height: 16 / _ca})
			}
			_fd = rowNumber
		}
		var height float64
		if row.X().HtAttr == nil {
			height = 16
			//TODO: if text wrap and is too long. it need to be heigher.
		} else {
			height = *row.X().HtAttr
		}
		hidden := row.X().HiddenAttr
		if hidden != nil && *hidden {
			height = 0
		}
		row_info = append(row_info, &rowInfo{height: height / _ca, _bgc: true, style: ctx.getStyle(row.X().SAttr)})
	}
	for _faa := len(row_info); _faa < ctx._dbc; _faa++ {
		row_info = append(row_info, &rowInfo{height: 16 / _ca})
	}
	ctx.rowInfo = row_info
}
