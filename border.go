package protoexcel

// Border adds a Border to cell
type Border int

const (
	// None adds no Border
	None Border = iota
	// Top adds a border to the top
	Top
	// Bottom adds a border to the bottom
	Bottom
	// DoubleBottom adds a double bottom border
	DoubleBottom
	// Left adds a border to the left
	Left
	// Right adds a border to the right
	Right
)
