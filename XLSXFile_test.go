package main

import "testing"

func TestString(t *testing.T) {
	o := OneDayClasses{
		Date: "2021-01-02",
		Classes: []string{"1", "2", "3", "4"},
	}

	want := `"2021-01-02", "1", "2", "3", "4"`
	if o.String() != want {
		t.Errorf("want: %s, got: %s", want, o.String())
	}
}
