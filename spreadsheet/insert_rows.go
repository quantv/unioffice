package spreadsheet

import (
	"github.com/unidoc/unioffice"
	"github.com/unidoc/unioffice/schema/soo/sml"
	"github.com/unidoc/unioffice/spreadsheet/formula"
	"github.com/unidoc/unioffice/spreadsheet/reference"
)

/**
* insert multiple rows at once
* update cell referrence in formula
**/

func (sheet *Sheet) InsertRows(rowNum int, rows uint32) uint32 {
	//find cell with formula and update ref if need.
	for _, row := range sheet.Rows() {
		for _, cell := range row.Cells() {
			if cell.HasFormula() {
				//no formula string to calculate
				if len(cell._gbf.F.Content) == 0 {
					continue
				}
				expr := formula.UpdateExpressionCellRef(sheet.Name(), cell.GetFormula(), uint32(rowNum), rows)
				//shared formula
				if cell._gbf.F.TAttr == sml.ST_CellFormulaTypeShared && cell._gbf.F.RefAttr != nil {
					top, bottom, err := reference.ParseRangeReference(*cell.X().F.RefAttr)
					if err != nil {
						continue
					}
					if top.RowIdx >= uint32(rowNum) {
						top.RowIdx += rows
					}
					if bottom.RowIdx >= uint32(rowNum) {
						bottom.RowIdx += rows
					}
					cell._gbf.F.RefAttr = unioffice.String(top.String() + ":" + bottom.String())
				}
				cell._gbf.F.Content = expr.String()
			}
		}
	}

	for range make([]int, rows) {
		sheet.InsertRow(rowNum)
	}

	return rows
}
