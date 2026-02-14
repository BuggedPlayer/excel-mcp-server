package excel

import (
	"testing"
)

func TestParseRange(t *testing.T) {
	tests := []struct {
		name      string
		input     string
		wantSCol  int
		wantSRow  int
		wantECol  int
		wantERow  int
		wantError bool
	}{
		{
			name:     "simple range",
			input:    "A1:C10",
			wantSCol: 1, wantSRow: 1, wantECol: 3, wantERow: 10,
		},
		{
			name:     "single cell",
			input:    "B5",
			wantSCol: 2, wantSRow: 5, wantECol: 2, wantERow: 5,
		},
		{
			name:     "absolute references",
			input:    "$A$1:$C$10",
			wantSCol: 1, wantSRow: 1, wantECol: 3, wantERow: 10,
		},
		{
			name:     "mixed absolute references",
			input:    "$A1:C$10",
			wantSCol: 1, wantSRow: 1, wantECol: 3, wantERow: 10,
		},
		{
			name:     "multi-letter columns",
			input:    "AA1:AZ100",
			wantSCol: 27, wantSRow: 1, wantECol: 52, wantERow: 100,
		},
		{
			name:      "empty string",
			input:     "",
			wantError: true,
		},
		{
			name:      "invalid format",
			input:     "not-a-range",
			wantError: true,
		},
		{
			name:      "missing row number",
			input:     "A:C",
			wantError: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sCol, sRow, eCol, eRow, err := ParseRange(tt.input)
			if tt.wantError {
				if err == nil {
					t.Errorf("ParseRange(%q) expected error, got nil", tt.input)
				}
				return
			}
			if err != nil {
				t.Errorf("ParseRange(%q) unexpected error: %v", tt.input, err)
				return
			}
			if sCol != tt.wantSCol || sRow != tt.wantSRow || eCol != tt.wantECol || eRow != tt.wantERow {
				t.Errorf("ParseRange(%q) = (%d,%d,%d,%d), want (%d,%d,%d,%d)",
					tt.input, sCol, sRow, eCol, eRow, tt.wantSCol, tt.wantSRow, tt.wantECol, tt.wantERow)
			}
		})
	}
}

func TestNormalizeRange(t *testing.T) {
	tests := []struct {
		name  string
		input string
		want  string
	}{
		{
			name:  "already normalized",
			input: "A1:C10",
			want:  "A1:C10",
		},
		{
			name:  "strips absolute references",
			input: "$A$1:$C$10",
			want:  "A1:C10",
		},
		{
			name:  "invalid input returns original",
			input: "not-a-range",
			want:  "not-a-range",
		},
		{
			name:  "empty string returns original",
			input: "",
			want:  "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := NormalizeRange(tt.input)
			if got != tt.want {
				t.Errorf("NormalizeRange(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}
