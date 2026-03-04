package oxml

// ===========================================================================
// CT_Settings — custom methods
// ===========================================================================

// EvenAndOddHeadersVal returns the value of w:evenAndOddHeaders/@w:val,
// or false if the element is not present.
func (s *CT_Settings) EvenAndOddHeadersVal() bool {
	eaoh := s.EvenAndOddHeaders()
	if eaoh == nil {
		return false
	}
	return eaoh.Val()
}

// SetEvenAndOddHeadersVal sets the evenAndOddHeaders flag.
// Passing false or nil-equivalent removes the element entirely.
func (s *CT_Settings) SetEvenAndOddHeadersVal(v *bool) error {
	if v == nil || !*v {
		s.RemoveEvenAndOddHeaders()
		return nil
	}
	return s.GetOrAddEvenAndOddHeaders().SetVal(true)
}
