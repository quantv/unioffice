package spreadsheet

import (
	"github.com/unidoc/unioffice"
	"github.com/unidoc/unioffice/schema/soo/sml"
	"github.com/unidoc/unioffice/spreadsheet/formula"
	"github.com/unidoc/unioffice/spreadsheet/reference"
)

func _calculate_delta(a, b uint32) int {
	if a >= b {
		return int(a - b)
	} else {
		return -int(b - a)
	}
}
func _add_delta(n uint32, d int) uint32 {
	if d >= 0 {
		return n + uint32(d)
	} else {
		v := uint32(-d)
		if v > n {
			return n
		}
		return n - v
	}
}
func _update_formual(ctx formula.UpdateContext, cell *Cell) {
	if !cell.HasFormula() {
		return
	}
	if len(cell._gbf.F.Content) == 0 {
		return
	}
	expr := formula.UpdateExpressionCellRef(ctx, cell.GetFormula())
	//shared formula
	if cell._gbf.F.TAttr == sml.ST_CellFormulaTypeShared && cell._gbf.F.RefAttr != nil {
		top, bottom, err := reference.ParseRangeReference(*cell.X().F.RefAttr)
		if err != nil {
			return
		}
		if top.RowIdx >= ctx.RowNum {
			top.RowIdx = _add_delta(top.RowIdx, ctx.DeltaR)
		}
		if bottom.RowIdx >= ctx.RowNum {
			bottom.RowIdx = _add_delta(bottom.RowIdx, ctx.DeltaR)
		}

		cell._gbf.F.RefAttr = unioffice.String(top.String() + ":" + bottom.String())
	}
	cell._gbf.F.Content = expr.String()
}

func (cell Cell) Copy(source Cell) {
	x := source.X()
	cell._gbf.SAttr = x.SAttr
	cell._gbf.TAttr = x.TAttr
	cell._gbf.CmAttr = x.CmAttr
	cell._gbf.VmAttr = x.VmAttr
	cell._gbf.PhAttr = x.PhAttr
	cell._gbf.V = x.V

	//TODO: should have a proper copy
	cell._gbf.Is = x.Is
	cell._gbf.ExtLst = x.ExtLst

	if x.F != nil {
		cell._gbf.F = sml.NewCT_CellFormula()
		cell._gbf.F.TAttr = x.F.TAttr
		cell._gbf.F.AcaAttr = x.F.AcaAttr
		cell._gbf.F.RefAttr = x.F.RefAttr
		cell._gbf.F.Dt2DAttr = x.F.Dt2DAttr
		cell._gbf.F.DtrAttr = x.F.DtrAttr
		cell._gbf.F.Del1Attr = x.F.Del1Attr
		cell._gbf.F.Del2Attr = x.F.Del2Attr
		cell._gbf.F.R1Attr = x.F.R1Attr
		cell._gbf.F.R2Attr = x.F.R2Attr
		cell._gbf.F.CaAttr = x.F.CaAttr
		cell._gbf.F.SiAttr = x.F.SiAttr
		cell._gbf.F.BxAttr = x.F.BxAttr
		cell._gbf.F.Content = x.F.Content

		//update cell ref
		if len(cell._gbf.F.Content) != 0 {
			cell_ref, _ := reference.ParseCellReference(cell.Reference())
			source_ref, _ := reference.ParseCellReference(source.Reference())

			ctx := formula.UpdateContext{
				SheetName: cell._becd.Name(),
				DeltaR:    _calculate_delta(cell_ref.RowIdx, source_ref.RowIdx),
				DeltaC:    _calculate_delta(cell_ref.ColumnIdx, source_ref.ColumnIdx),
				RowNum:    0,
			}
			_update_formual(ctx, &cell)
		}
	}
}

func (sheet *Sheet) CopyRows(source, dest uint32, rows int) int {
	sourceRow := sheet.Row(source)
	for i := 0; i < rows; i++ {
		destRow := sheet.Row(dest + uint32(i))
		for _, cell := range sourceRow.Cells() {
			col, _ := cell.Column()
			destRow.Cell(col).Copy(cell)
		}
	}
	return rows
}

// InsertRows insert `rows` rows at `rowNum` and update formula cell referrence
func (sheet *Sheet) InsertRows(rowNum int, rows uint32) uint32 {
	//find cell with formula and update ref if need.
	ctx := formula.UpdateContext{
		SheetName: sheet.Name(),
		DeltaR:    int(rows),
		DeltaC:    0,
		RowNum:    uint32(rowNum),
	}
	for _, row := range sheet.Rows() {
		for _, cell := range row.Cells() {
			_update_formual(ctx, &cell)
		}
	}

	for range make([]int, rows) {
		sheet.InsertRow(rowNum)
	}

	return rows
}
