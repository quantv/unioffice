//
// Copyright 2020 FoxyUtils ehf. All rights reserved.
//
// This is a commercial product and requires a license to operate.
// A trial license can be obtained at https://unidoc.io
//
// DO NOT EDIT: generated by unitwist Go source code obfuscator.
//
// Use of this source code is governed by the UniDoc End User License Agreement
// terms that can be accessed at https://unidoc.io/eula/

package drawing ;import (_e "github.com/unidoc/unioffice";_b "github.com/unidoc/unioffice/color";_ec "github.com/unidoc/unioffice/measurement";_ef "github.com/unidoc/unioffice/schema/soo/dml";);func (_d LineProperties )SetSolidFill (c _b .Color ){_d .clearFill ();_d ._f .SolidFill =_ef .NewCT_SolidColorFillProperties ();_d ._f .SolidFill .SrgbClr =_ef .NewCT_SRgbColor ();_d ._f .SolidFill .SrgbClr .ValAttr =*c .AsRGBString ();};

// SetJoin sets the line join style.
func (_ea LineProperties )SetJoin (e LineJoin ){_ea ._f .Round =nil ;_ea ._f .Miter =nil ;_ea ._f .Bevel =nil ;switch e {case LineJoinRound :_ea ._f .Round =_ef .NewCT_LineJoinRound ();case LineJoinBevel :_ea ._f .Bevel =_ef .NewCT_LineJoinBevel ();case LineJoinMiter :_ea ._f .Miter =_ef .NewCT_LineJoinMiterProperties ();};};

// SetText sets the run's text contents.
func (_ac Run )SetText (s string ){_ac ._eg .Br =nil ;_ac ._eg .Fld =nil ;if _ac ._eg .R ==nil {_ac ._eg .R =_ef .NewCT_RegularTextRun ();};_ac ._eg .R .T =s ;};

// MakeParagraph constructs a new paragraph wrapper.
func MakeParagraph (x *_ef .CT_TextParagraph )Paragraph {return Paragraph {x }};func (_baa LineProperties )SetNoFill (){_baa .clearFill ();_baa ._f .NoFill =_ef .NewCT_NoFillProperties ();};

// Properties returns the paragraph properties.
func (_ca Paragraph )Properties ()ParagraphProperties {if _ca ._a .PPr ==nil {_ca ._a .PPr =_ef .NewCT_TextParagraphProperties ();};return MakeParagraphProperties (_ca ._a .PPr );};func (_ga ShapeProperties )SetNoFill (){_ga .clearFill ();_ga ._be .NoFill =_ef .NewCT_NoFillProperties ();};

// AddRun adds a new run to a paragraph.
func (_fd Paragraph )AddRun ()Run {_gb :=MakeRun (_ef .NewEG_TextRun ());_fd ._a .EG_TextRun =append (_fd ._a .EG_TextRun ,_gb .X ());return _gb ;};

// MakeRun constructs a new Run wrapper.
func MakeRun (x *_ef .EG_TextRun )Run {return Run {x }};

// SetWidth sets the width of the shape.
func (_beg ShapeProperties )SetWidth (w _ec .Distance ){_beg .ensureXfrm ();if _beg ._be .Xfrm .Ext ==nil {_beg ._be .Xfrm .Ext =_ef .NewCT_PositiveSize2D ();};_beg ._be .Xfrm .Ext .CxAttr =int64 (w /_ec .EMU );};

// SetLevel sets the level of indentation of a paragraph.
func (_ae ParagraphProperties )SetLevel (idx int32 ){_ae ._ed .LvlAttr =_e .Int32 (idx )};

// SetFont controls the font of a run.
func (_bgb RunProperties )SetFont (s string ){_bgb ._af .Latin =_ef .NewCT_TextFont ();_bgb ._af .Latin .TypefaceAttr =s ;};

// X returns the inner wrapped XML type.
func (_aa Paragraph )X ()*_ef .CT_TextParagraph {return _aa ._a };

// SetBold controls the bolding of a run.
func (_ee RunProperties )SetBold (b bool ){_ee ._af .BAttr =_e .Bool (b )};

// SetAlign controls the paragraph alignment
func (_cf ParagraphProperties )SetAlign (a _ef .ST_TextAlignType ){_cf ._ed .AlgnAttr =a };

// SetWidth sets the line width, MS products treat zero as the minimum width
// that can be displayed.
func (_ba LineProperties )SetWidth (w _ec .Distance ){_ba ._f .WAttr =_e .Int32 (int32 (w /_ec .EMU ))};

// MakeParagraphProperties constructs a new ParagraphProperties wrapper.
func MakeParagraphProperties (x *_ef .CT_TextParagraphProperties )ParagraphProperties {return ParagraphProperties {x };};func (_bf ShapeProperties )LineProperties ()LineProperties {if _bf ._be .Ln ==nil {_bf ._be .Ln =_ef .NewCT_LineProperties ();};return LineProperties {_bf ._be .Ln };};

// GetPosition gets the position of the shape in EMU.
func (_dd ShapeProperties )GetPosition ()(int64 ,int64 ){_dd .ensureXfrm ();if _dd ._be .Xfrm .Off ==nil {_dd ._be .Xfrm .Off =_ef .NewCT_Point2D ();};return *_dd ._be .Xfrm .Off .XAttr .ST_CoordinateUnqualified ,*_dd ._be .Xfrm .Off .YAttr .ST_CoordinateUnqualified ;};

// SetFlipVertical controls if the shape is flipped vertically.
func (_dg ShapeProperties )SetFlipVertical (b bool ){_dg .ensureXfrm ();if !b {_dg ._be .Xfrm .FlipVAttr =nil ;}else {_dg ._be .Xfrm .FlipVAttr =_e .Bool (true );};};type ShapeProperties struct{_be *_ef .CT_ShapeProperties };

// X returns the inner wrapped XML type.
func (_db ParagraphProperties )X ()*_ef .CT_TextParagraphProperties {return _db ._ed };

// SetBulletChar sets the bullet character for the paragraph.
func (_gf ParagraphProperties )SetBulletChar (c string ){if c ==""{_gf ._ed .BuChar =nil ;}else {_gf ._ed .BuChar =_ef .NewCT_TextCharBullet ();_gf ._ed .BuChar .CharAttr =c ;};};

// SetPosition sets the position of the shape.
func (_fgg ShapeProperties )SetPosition (x ,y _ec .Distance ){_fgg .ensureXfrm ();if _fgg ._be .Xfrm .Off ==nil {_fgg ._be .Xfrm .Off =_ef .NewCT_Point2D ();};_fgg ._be .Xfrm .Off .XAttr .ST_CoordinateUnqualified =_e .Int64 (int64 (x /_ec .EMU ));_fgg ._be .Xfrm .Off .YAttr .ST_CoordinateUnqualified =_e .Int64 (int64 (y /_ec .EMU ));};

// Run is a run within a paragraph.
type Run struct{_eg *_ef .EG_TextRun };func (_dbd ShapeProperties )clearFill (){_dbd ._be .NoFill =nil ;_dbd ._be .BlipFill =nil ;_dbd ._be .GradFill =nil ;_dbd ._be .GrpFill =nil ;_dbd ._be .SolidFill =nil ;_dbd ._be .PattFill =nil ;};

// SetSize sets the font size of the run text
func (_dbb RunProperties )SetSize (sz _ec .Distance ){_dbb ._af .SzAttr =_e .Int32 (int32 (sz /_ec .HundredthPoint ));};

// SetHeight sets the height of the shape.
func (_bab ShapeProperties )SetHeight (h _ec .Distance ){_bab .ensureXfrm ();if _bab ._be .Xfrm .Ext ==nil {_bab ._be .Xfrm .Ext =_ef .NewCT_PositiveSize2D ();};_bab ._be .Xfrm .Ext .CyAttr =int64 (h /_ec .EMU );};func MakeShapeProperties (x *_ef .CT_ShapeProperties )ShapeProperties {return ShapeProperties {x }};

// Properties returns the run's properties.
func (_gc Run )Properties ()RunProperties {if _gc ._eg .R ==nil {_gc ._eg .R =_ef .NewCT_RegularTextRun ();};if _gc ._eg .R .RPr ==nil {_gc ._eg .R .RPr =_ef .NewCT_TextCharacterProperties ();};return RunProperties {_gc ._eg .R .RPr };};func (_edg ShapeProperties )ensureXfrm (){if _edg ._be .Xfrm ==nil {_edg ._be .Xfrm =_ef .NewCT_Transform2D ();};};

// SetBulletFont controls the font for the bullet character.
func (_cc ParagraphProperties )SetBulletFont (f string ){if f ==""{_cc ._ed .BuFont =nil ;}else {_cc ._ed .BuFont =_ef .NewCT_TextFont ();_cc ._ed .BuFont .TypefaceAttr =f ;};};

// SetSize sets the width and height of the shape.
func (_fad ShapeProperties )SetSize (w ,h _ec .Distance ){_fad .SetWidth (w );_fad .SetHeight (h )};func (_cg ShapeProperties )SetSolidFill (c _b .Color ){_cg .clearFill ();_cg ._be .SolidFill =_ef .NewCT_SolidColorFillProperties ();_cg ._be .SolidFill .SrgbClr =_ef .NewCT_SRgbColor ();_cg ._be .SolidFill .SrgbClr .ValAttr =*c .AsRGBString ();};

// X returns the inner wrapped XML type.
func (_fg LineProperties )X ()*_ef .CT_LineProperties {return _fg ._f };

// SetGeometry sets the shape type of the shape
func (_cb ShapeProperties )SetGeometry (g _ef .ST_ShapeType ){if _cb ._be .PrstGeom ==nil {_cb ._be .PrstGeom =_ef .NewCT_PresetGeometry2D ();};_cb ._be .PrstGeom .PrstAttr =g ;};

// RunProperties controls the run properties.
type RunProperties struct{_af *_ef .CT_TextCharacterProperties ;};

// AddBreak adds a new line break to a paragraph.
func (_fc Paragraph )AddBreak (){_fa :=_ef .NewEG_TextRun ();_fa .Br =_ef .NewCT_TextLineBreak ();_fc ._a .EG_TextRun =append (_fc ._a .EG_TextRun ,_fa );};

// SetFlipHorizontal controls if the shape is flipped horizontally.
func (_aec ShapeProperties )SetFlipHorizontal (b bool ){_aec .ensureXfrm ();if !b {_aec ._be .Xfrm .FlipHAttr =nil ;}else {_aec ._be .Xfrm .FlipHAttr =_e .Bool (true );};};type LineProperties struct{_f *_ef .CT_LineProperties };

// SetNumbered controls if bullets are numbered or not.
func (_ce ParagraphProperties )SetNumbered (scheme _ef .ST_TextAutonumberScheme ){if scheme ==_ef .ST_TextAutonumberSchemeUnset {_ce ._ed .BuAutoNum =nil ;}else {_ce ._ed .BuAutoNum =_ef .NewCT_TextAutonumberBullet ();_ce ._ed .BuAutoNum .TypeAttr =scheme ;};};

// X returns the inner wrapped XML type.
func (_ab ShapeProperties )X ()*_ef .CT_ShapeProperties {return _ab ._be };

// Paragraph is a paragraph within a document.
type Paragraph struct{_a *_ef .CT_TextParagraph };func (_g LineProperties )clearFill (){_g ._f .NoFill =nil ;_g ._f .GradFill =nil ;_g ._f .SolidFill =nil ;_g ._f .PattFill =nil ;};const (LineJoinRound LineJoin =iota ;LineJoinBevel ;LineJoinMiter ;);

// SetSolidFill controls the text color of a run.
func (_bg RunProperties )SetSolidFill (c _b .Color ){_bg ._af .NoFill =nil ;_bg ._af .BlipFill =nil ;_bg ._af .GradFill =nil ;_bg ._af .GrpFill =nil ;_bg ._af .PattFill =nil ;_bg ._af .SolidFill =_ef .NewCT_SolidColorFillProperties ();_bg ._af .SolidFill .SrgbClr =_ef .NewCT_SRgbColor ();_bg ._af .SolidFill .SrgbClr .ValAttr =*c .AsRGBString ();};

// LineJoin is the type of line join
type LineJoin byte ;

// MakeRunProperties constructs a new RunProperties wrapper.
func MakeRunProperties (x *_ef .CT_TextCharacterProperties )RunProperties {return RunProperties {x }};

// X returns the inner wrapped XML type.
func (_eb Run )X ()*_ef .EG_TextRun {return _eb ._eg };

// ParagraphProperties allows controlling paragraph properties.
type ParagraphProperties struct{_ed *_ef .CT_TextParagraphProperties ;};