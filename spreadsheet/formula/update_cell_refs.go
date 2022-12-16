package formula

import (
	"github.com/unidoc/unioffice/spreadsheet/reference"
)

type LinkedExpression interface {
	UpdateCellRef(ctx UpdateContext) Expression
}

func UpdateExpressionCellRef(ctx UpdateContext, s string) Expression {
	expr := ParseString(s)
	return expr.(LinkedExpression).UpdateCellRef(ctx)
}

type UpdateContext struct {
	SheetName string
	DeltaR    int
	DeltaC    int
	RowNum    uint32
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

func (e CellRef) UpdateCellRef(ctx UpdateContext) Expression {
	ref, err := reference.ParseCellReference(e._cbe)
	if err != nil {
		return e
	}
	//do not care cell from other sheet
	if len(ref.SheetName) != 0 && ref.SheetName != ctx.SheetName {
		return e
	}

	//do not update cell above RowNum becuase it does not move
	if ref.RowIdx < ctx.RowNum {
		return e
	}

	if !ref.AbsoluteRow {
		ref.RowIdx = _add_delta(ref.RowIdx, ctx.DeltaR)
	}
	if !ref.AbsoluteColumn {
		ref.ColumnIdx = _add_delta(ref.ColumnIdx, ctx.DeltaC)
		ref.Column = reference.IndexToColumn(ref.ColumnIdx)
	}
	e._cbe = ref.String()
	return e
}
func (e BinaryExpr) UpdateCellRef(ctx UpdateContext) Expression {
	e._baa = e._baa.(LinkedExpression).UpdateCellRef(ctx)
	e._ced = e._ced.(LinkedExpression).UpdateCellRef(ctx)
	return e
}
func (e Number) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e Error) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}

func (e HorizontalRange) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e FunctionCall) UpdateCellRef(ctx UpdateContext) Expression {
	for k, exp := range e._bdegf {
		e._bdegf[k] = exp.(LinkedExpression).UpdateCellRef(ctx)
	}
	return e
}
func (e String) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e PrefixVerticalRange) UpdateCellRef(ctx UpdateContext) Expression {
	e._fcbe = e._fcbe.(LinkedExpression).UpdateCellRef(ctx)
	return e
}
func (e Negate) UpdateCellRef(ctx UpdateContext) Expression {
	e._eaaad = e._eaaad.(LinkedExpression).UpdateCellRef(ctx)
	return e
}
func (e NamedRangeRef) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e PrefixExpr) UpdateCellRef(ctx UpdateContext) Expression {
	e._bdga = e._bdga.(LinkedExpression).UpdateCellRef(ctx)
	e._bebg = e._bebg.(LinkedExpression).UpdateCellRef(ctx)
	return e
}
func (e PrefixHorizontalRange) UpdateCellRef(ctx UpdateContext) Expression {
	e._bedeg = e._bedeg.(LinkedExpression).UpdateCellRef(ctx)
	return e
}
func (e ConstArrayExpr) UpdateCellRef(ctx UpdateContext) Expression {
	for i, row := range e._ccd {
		for k, col := range row {
			e._ccd[i][k] = col.(LinkedExpression).UpdateCellRef(ctx)
		}
	}
	return e
}
func (e SheetPrefixExpr) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e EmptyExpr) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e Bool) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e VerticalRange) UpdateCellRef(ctx UpdateContext) Expression {
	return e
}
func (e PrefixRangeExpr) UpdateCellRef(ctx UpdateContext) Expression {
	e._cfddb = e._cfddb.(LinkedExpression).UpdateCellRef(ctx)
	e._befed = e._befed.(LinkedExpression).UpdateCellRef(ctx)
	return e
}
func (e Range) UpdateCellRef(ctx UpdateContext) Expression {
	e._cdacg = e._cdacg.(LinkedExpression).UpdateCellRef(ctx)
	e._faceba = e._faceba.(LinkedExpression).UpdateCellRef(ctx)
	return e
}
