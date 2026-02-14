package excel

import (
	"testing"
)

func TestCalculateFixedSizeRanges(t *testing.T) {
	tests := []struct {
		name      string
		dimension string
		pageSize  int
		want      []string
	}{
		{
			name:      "single page fits all",
			dimension: "A1:C10",
			pageSize:  100,
			want:      []string{"A1:C10"},
		},
		{
			name:      "multiple pages",
			dimension: "A1:C10",
			pageSize:  9, // 3 cols * 3 rows = 9 cells per page
			want:      []string{"A1:C3", "A4:C6", "A7:C9", "A10:C10"},
		},
		{
			name:      "single column",
			dimension: "A1:A5",
			pageSize:  2,
			want:      []string{"A1:A2", "A3:A4", "A5:A5"},
		},
		{
			name:      "page size smaller than columns forces 1 row per page",
			dimension: "A1:E3",
			pageSize:  2, // 5 cols but pageSize=2, so rowsPerPage = max(1, 2/5) = 1
			want:      []string{"A1:E1", "A2:E2", "A3:E3"},
		},
		{
			name:      "single cell",
			dimension: "B2:B2",
			pageSize:  100,
			want:      []string{"B2:B2"},
		},
		{
			name:      "invalid dimension returns empty",
			dimension: "invalid",
			pageSize:  100,
			want:      []string{},
		},
		{
			name:      "empty dimension returns empty",
			dimension: "",
			pageSize:  100,
			want:      []string{},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := calculateFixedSizeRanges(tt.dimension, tt.pageSize)
			if len(got) != len(tt.want) {
				t.Errorf("calculateFixedSizeRanges(%q, %d) returned %d ranges, want %d: got %v",
					tt.dimension, tt.pageSize, len(got), len(tt.want), got)
				return
			}
			for i := range got {
				if got[i] != tt.want[i] {
					t.Errorf("calculateFixedSizeRanges(%q, %d)[%d] = %q, want %q",
						tt.dimension, tt.pageSize, i, got[i], tt.want[i])
				}
			}
		})
	}
}

func TestPagingRangeService_FindNextRange(t *testing.T) {
	service := NewPagingRangeService(nil) // strategy not needed for this method

	tests := []struct {
		name         string
		allRanges    []string
		currentRange string
		want         string
	}{
		{
			name:         "returns next range",
			allRanges:    []string{"A1:A10", "A11:A20", "A21:A30"},
			currentRange: "A1:A10",
			want:         "A11:A20",
		},
		{
			name:         "returns empty for last range",
			allRanges:    []string{"A1:A10", "A11:A20"},
			currentRange: "A11:A20",
			want:         "",
		},
		{
			name:         "returns empty for unknown range",
			allRanges:    []string{"A1:A10", "A11:A20"},
			currentRange: "B1:B10",
			want:         "",
		},
		{
			name:         "empty ranges returns empty",
			allRanges:    []string{},
			currentRange: "A1:A10",
			want:         "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := service.FindNextRange(tt.allRanges, tt.currentRange)
			if got != tt.want {
				t.Errorf("FindNextRange(%v, %q) = %q, want %q",
					tt.allRanges, tt.currentRange, got, tt.want)
			}
		})
	}
}

func TestPagingRangeService_FilterRemainingPagingRanges(t *testing.T) {
	service := NewPagingRangeService(nil)

	tests := []struct {
		name        string
		allRanges   []string
		knownRanges []string
		want        []string
	}{
		{
			name:        "no known ranges returns all",
			allRanges:   []string{"A1:A10", "A11:A20"},
			knownRanges: []string{},
			want:        []string{"A1:A10", "A11:A20"},
		},
		{
			name:        "filters known ranges",
			allRanges:   []string{"A1:A10", "A11:A20", "A21:A30"},
			knownRanges: []string{"A1:A10"},
			want:        []string{"A11:A20", "A21:A30"},
		},
		{
			name:        "all known returns empty",
			allRanges:   []string{"A1:A10"},
			knownRanges: []string{"A1:A10"},
			want:        []string{},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := service.FilterRemainingPagingRanges(tt.allRanges, tt.knownRanges)
			if len(got) != len(tt.want) {
				t.Errorf("FilterRemainingPagingRanges() returned %d, want %d: got %v",
					len(got), len(tt.want), got)
				return
			}
			for i := range got {
				if got[i] != tt.want[i] {
					t.Errorf("FilterRemainingPagingRanges()[%d] = %q, want %q", i, got[i], tt.want[i])
				}
			}
		})
	}
}
