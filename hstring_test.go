//go:build windows

package ole

import "testing"

func TestHStringRoundTrip(t *testing.T) {
	testCases := []struct {
		name  string
		value string
	}{
		{name: "empty", value: ""},
		{name: "ascii", value: "hello"},
		{name: "bmp unicode", value: "こんにちは世界"},
		{name: "surrogate pair", value: "hello 😀"},
	}

	for _, testCase := range testCases {
		t.Run(testCase.name, func(t *testing.T) {
			hstring, err := NewHString(testCase.value)
			if err != nil {
				t.Fatalf("NewHString failed: %v", err)
			}
			defer func() {
				if err := DeleteHString(hstring); err != nil {
					t.Fatalf("DeleteHString failed: %v", err)
				}
			}()

			if got := hstring.String(); got != testCase.value {
				t.Fatalf("HString round-trip mismatch: got %q want %q", got, testCase.value)
			}
		})
	}
}

func TestZeroHStringStringIsEmpty(t *testing.T) {
	var hstring HString

	if got := hstring.String(); got != "" {
		t.Fatalf("zero HString String() = %q, want empty string", got)
	}
}
