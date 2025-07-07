package calendardata

// dayRange associates a start and end day with the resulting text
type dayRange struct {
	start int
	end   int
	text  string
}

var dayRanges = []dayRange{
	{1, 5, "6\n\n1"},
	{6, 8, "8\n\n1"},
	{9, 11, "9\n\n1"},
	{12, 12, "10\n\n1"},
	{13, 13, "4\n\n2"},
	{14, 14, "4\n6\n3"},
	{15, 15, "5\n6\n4"},
	{16, 16, "5\n10\n3"},
	{17, 17, "5\n10\n2"},
	{18, 19, "6\n12\n1"},
	{20, 20, "6\n14\n1"},
	{21, 21, "6\n16\n1"},
	{22, 22, "6\n14\n1"},
	{23, 24, "6\n12\n1"},
	{25, 26, "6\n10\n1"},
	{27, 28, "6\n6\n1"},
}

var dayRangesDoc = []dayRange{
	{1, 5, "\x20\x2006\x2000\x2001"},
	{6, 8, "\x20\x2008\x2000\x2001"},
	{9, 11, "\x20\x2009\x2000\x2001"},
	{12, 12, "\x20\x2010\x2000\x2001"},
	{13, 13, "\x20\x2004\x2000\x2002"},
	{14, 14, "\x20\x2004\x2006\x2003"},
	{15, 15, "\x20\x2005\x2006\x2004"},
	{16, 16, "\x20\x2005\x2010\x2003"},
	{17, 17, "\x20\x2005\x2010\x2002"},
	{18, 19, "\x20\x2006\x2012\x2001"},
	{20, 20, "\x20\x2006\x2014\x2001"},
	{21, 21, "\x20\x2006\x2016\x2001"},
	{22, 22, "\x20\x2006\x2014\x2001"},
	{23, 24, "\x20\x2006\x2012\x2001"},
	{25, 26, "\x20\x2006\x2010\x2001"},
	{27, 28, "\x20\x2006\x2006\x2001"},
}

func GetAmountText(day int) string {
	for _, dr := range dayRanges {
		if day >= dr.start && day <= dr.end {
			return dr.text
		}
	}
	// Default if no range matches
	return ""
}
func GetAmountTextDoc(day int) string {
	for _, dr := range dayRangesDoc {
		if day >= dr.start && day <= dr.end {
			return dr.text
		}
	}
	// Default if no range matches
	return ""
}
