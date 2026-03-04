package docx

import (
	"fmt"

	"github.com/vortex/go-docx/pkg/docx/enum"
	"github.com/vortex/go-docx/pkg/docx/oxml"
)

// TabStops provides access to the tab stops defined for a paragraph or style.
// Supports iteration, indexed access, delete, and length.
//
// Mirrors Python TabStops(ElementProxy).
type TabStops struct {
	pPr *oxml.CT_PPr
}

// newTabStops creates a new TabStops proxy.
func newTabStops(pPr *oxml.CT_PPr) *TabStops {
	return &TabStops{pPr: pPr}
}

// Len returns the number of tab stops.
func (ts *TabStops) Len() int {
	tabs := ts.pPr.Tabs()
	if tabs == nil {
		return 0
	}
	return len(tabs.TabList())
}

// Get returns the tab stop at the given index.
func (ts *TabStops) Get(idx int) (*TabStop, error) {
	tabs := ts.pPr.Tabs()
	if tabs == nil {
		return nil, fmt.Errorf("docx: tab index out of range")
	}
	lst := tabs.TabList()
	if idx < 0 || idx >= len(lst) {
		return nil, fmt.Errorf("docx: tab index [%d] out of range", idx)
	}
	return newTabStop(lst[idx]), nil
}

// Delete removes the tab stop at the given index.
func (ts *TabStops) Delete(idx int) error {
	tabs := ts.pPr.Tabs()
	if tabs == nil {
		return fmt.Errorf("docx: tab index out of range")
	}
	lst := tabs.TabList()
	if idx < 0 || idx >= len(lst) {
		return fmt.Errorf("docx: tab index [%d] out of range", idx)
	}
	tabs.RawElement().RemoveChild(lst[idx].RawElement())
	if len(tabs.TabList()) == 0 {
		ts.pPr.RemoveTabs()
	}
	return nil
}

// Iter returns all tab stops in document order.
func (ts *TabStops) Iter() []*TabStop {
	tabs := ts.pPr.Tabs()
	if tabs == nil {
		return nil
	}
	lst := tabs.TabList()
	result := make([]*TabStop, len(lst))
	for i, t := range lst {
		result[i] = newTabStop(t)
	}
	return result
}

// AddTabStop adds a new tab stop at the given position (EMU) with alignment
// and leader. Defaults: alignment=LEFT, leader=SPACES.
//
// Mirrors Python TabStops.add_tab_stop.
func (ts *TabStops) AddTabStop(position int, alignment enum.WdTabAlignment, leader enum.WdTabLeader) (*TabStop, error) {
	tabs := ts.pPr.GetOrAddTabs()
	tab, err := tabs.InsertTabInOrder(position, alignment, leader)
	if err != nil {
		return nil, fmt.Errorf("docx: adding tab stop: %w", err)
	}
	return newTabStop(tab), nil
}

// ClearAll removes all custom tab stops.
//
// Mirrors Python TabStops.clear_all.
func (ts *TabStops) ClearAll() {
	ts.pPr.RemoveTabs()
}

// TabStop represents an individual tab stop.
//
// Mirrors Python TabStop(ElementProxy).
type TabStop struct {
	tab *oxml.CT_TabStop
}

// newTabStop creates a new TabStop proxy.
func newTabStop(tab *oxml.CT_TabStop) *TabStop {
	return &TabStop{tab: tab}
}

// Alignment returns the tab alignment.
func (t *TabStop) Alignment() (enum.WdTabAlignment, error) {
	return t.tab.Val()
}

// SetAlignment sets the tab alignment.
func (t *TabStop) SetAlignment(v enum.WdTabAlignment) error {
	return t.tab.SetVal(v)
}

// Leader returns the tab leader.
func (t *TabStop) Leader() (enum.WdTabLeader, error) {
	return t.tab.Leader()
}

// SetLeader sets the tab leader.
func (t *TabStop) SetLeader(v enum.WdTabLeader) error {
	return t.tab.SetLeader(v)
}

// Position returns the tab position in EMU.
func (t *TabStop) Position() (int, error) {
	return t.tab.Pos()
}

// SetPosition changes the position of this tab stop.
// The tab is re-inserted in position order.
//
// Mirrors Python TabStop.position setter.
func (t *TabStop) SetPosition(v int) error {
	tabs := t.tab.RawElement().Parent()
	if tabs == nil {
		return fmt.Errorf("docx: tab stop has no parent")
	}
	align, err := t.tab.Val()
	if err != nil {
		return err
	}
	leader, err := t.tab.Leader()
	if err != nil {
		return err
	}
	parent := &oxml.CT_TabStops{Element: oxml.WrapElement(tabs)}
	newTab, err := parent.InsertTabInOrder(v, align, leader)
	if err != nil {
		return err
	}
	tabs.RemoveChild(t.tab.RawElement())
	t.tab = newTab
	return nil
}
