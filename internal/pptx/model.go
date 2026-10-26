package pptx

// Model is the JSON document exchanged with web/slides.js (see the board
// decisions "PPTX MODEL JSON" and "PPTX MODEL ADDITIONS").
// Units: EMU for positions/sizes, points for font sizes, degrees for rotation.
type Model struct {
	Size    Size              `json:"size"`
	Theme   Theme             `json:"theme"`
	Layouts []*Layout         `json:"layouts"`
	Slides  []*Slide          `json:"slides"`
	Media   map[string]string `json:"media,omitempty"`
}

type Size struct {
	CX int64 `json:"cx"`
	CY int64 `json:"cy"`
}

type Theme struct {
	MajorFont string            `json:"majorFont"`
	MinorFont string            `json:"minorFont"`
	Colors    map[string]string `json:"colors"`
}

type Layout struct {
	ID           string   `json:"id"`
	Name         string   `json:"name"`
	Bg           *Bg      `json:"bg,omitempty"`
	Placeholders []*Shape `json:"placeholders,omitempty"`
	Decor        []*Shape `json:"decor,omitempty"`
}

type Slide struct {
	ID     string   `json:"id"`
	Src    *string  `json:"src"`
	Layout string   `json:"layout"`
	Hidden bool     `json:"hidden"`
	Bg     *Bg      `json:"bg"`
	Notes  string   `json:"notes"`
	Decor  []*Shape `json:"decor"`
	Shapes []*Shape `json:"shapes"`
}

type Bg struct {
	Color string `json:"color,omitempty"`
	Image string `json:"image,omitempty"`
}

type Line struct {
	Color   string `json:"color"`
	Width   int64  `json:"width"`
	Dash    string `json:"dash"`
	HeadEnd string `json:"headEnd,omitempty"`
	TailEnd string `json:"tailEnd,omitempty"`
}

type Crop struct {
	L float64 `json:"l"`
	T float64 `json:"t"`
	R float64 `json:"r"`
	B float64 `json:"b"`
}

type TableStyle struct {
	ID       string `json:"id,omitempty"`
	FirstRow bool   `json:"firstRow,omitempty"`
	FirstCol bool   `json:"firstCol,omitempty"`
	LastRow  bool   `json:"lastRow,omitempty"`
	LastCol  bool   `json:"lastCol,omitempty"`
	BandRow  bool   `json:"bandRow,omitempty"`
	BandCol  bool   `json:"bandCol,omitempty"`
}

type Shape struct {
	ID    string  `json:"id"`
	Name  string  `json:"name"`
	Descr string  `json:"descr,omitempty"`
	Kind  string  `json:"kind"` // text | shape | image | table | opaque
	Ph    *string `json:"ph"`
	PhIdx string  `json:"phIdx,omitempty"`

	X     int64   `json:"x"`
	Y     int64   `json:"y"`
	W     int64   `json:"w"`
	H     int64   `json:"h"`
	Rot   float64 `json:"rot"`
	FlipH bool    `json:"flipH"`
	FlipV bool    `json:"flipV"`

	Geom      string  `json:"geom,omitempty"`
	Path      string  `json:"path,omitempty"`
	CustGeom  string  `json:"custGeom,omitempty"`
	Fill      *string `json:"fill"`
	FillImage string  `json:"fillImage,omitempty"`
	Line      *Line   `json:"line"`

	Anchor         string  `json:"anchor,omitempty"`
	Inset          []int64 `json:"inset,omitempty"`
	Autofit        string  `json:"autofit,omitempty"`
	FontScale      float64 `json:"fontScale,omitempty"`
	LnSpcReduction float64 `json:"lnSpcReduction,omitempty"`
	Wrap           string  `json:"wrap,omitempty"`
	Vert           string  `json:"vert,omitempty"`
	Paras          []*Para `json:"paras,omitempty"`

	Src  string `json:"src,omitempty"`
	Crop *Crop  `json:"crop,omitempty"`

	Rows       [][]*Cell   `json:"rows,omitempty"`
	ColW       []int64     `json:"colW,omitempty"`
	RowH       []int64     `json:"rowH,omitempty"`
	TableStyle *TableStyle `json:"tableStyle,omitempty"`

	XML      string         `json:"xml,omitempty"`
	Rels     map[string]Rel `json:"rels,omitempty"`
	Preview  string         `json:"preview,omitempty"`
	Children []*Shape       `json:"children,omitempty"`
}

type Para struct {
	Align          string  `json:"align"`
	Level          int     `json:"level"`
	Bullet         *Bullet `json:"bullet"`
	SpaceBefore    float64 `json:"spaceBefore"`
	SpaceAfter     float64 `json:"spaceAfter"`
	LineSpacing    float64 `json:"lineSpacing"`
	LineSpacingPts float64 `json:"lineSpacingPts,omitempty"`
	Indent         int64   `json:"indent"`
	MarL           int64   `json:"marL"`
	Runs           []*Run  `json:"runs"`
	End            *Run    `json:"end,omitempty"`
}

type Bullet struct {
	Char    string  `json:"char,omitempty"`
	Auto    string  `json:"auto,omitempty"`
	StartAt int     `json:"startAt,omitempty"`
	Font    string  `json:"font,omitempty"`
	Color   string  `json:"color,omitempty"`
	Size    float64 `json:"size,omitempty"`
}

type Run struct {
	Text      string  `json:"text"`
	B         bool    `json:"b"`
	I         bool    `json:"i"`
	U         bool    `json:"u"`
	Strike    bool    `json:"strike"`
	Size      float64 `json:"size"`
	Color     string  `json:"color"`
	Font      string  `json:"font"`
	Highlight *string `json:"highlight"`
	Link      *string `json:"link"`
	Baseline  int     `json:"baseline,omitempty"`
	Field     string  `json:"field,omitempty"`
}

type Cell struct {
	Paras   []*Para  `json:"paras"`
	Fill    *string  `json:"fill"`
	Span    int      `json:"span"`
	RowSpan int      `json:"rowSpan"`
	HMerge  bool     `json:"hMerge,omitempty"`
	VMerge  bool     `json:"vMerge,omitempty"`
	Borders *Borders `json:"borders,omitempty"`
	Anchor  string   `json:"anchor,omitempty"`
	Inset   []int64  `json:"inset,omitempty"`
}

type Borders struct {
	L *Line `json:"l"`
	R *Line `json:"r"`
	T *Line `json:"t"`
	B *Line `json:"b"`
}

func strp(s string) *string { return &s }
