//go:build windows && arm64
// +build windows,arm64

package ole

import (
	"math"
	"reflect"
	"testing"
	"time"
)

func TestGetVariantDate(t *testing.T) {
	type args struct {
		value uint64
	}
	tests := []struct {
		name    string
		args    args
		want    time.Time
		wantErr bool
	}{
		{
			name:    "2023-10-30 23:30:30:000",
			args:    args{value: math.Float64bits(45229.9795138889)},
			want:    time.Date(2023, 10, 30, 23, 30, 30, 0, time.UTC),
			wantErr: false,
		},
		{
			name:    "2023-10-30 23:30:30:355",
			args:    args{value: math.Float64bits(45229.979518)},
			want:    time.Date(2023, 10, 30, 23, 30, 30, 355000000, time.UTC),
			wantErr: false,
		},
		{
			name:    "2023-10-30 23:30:30:960",
			args:    args{value: math.Float64bits(45229.979525)},
			want:    time.Date(2023, 10, 30, 23, 30, 30, 960000000, time.UTC),
			wantErr: false,
		},
		{
			name:    "min OLE date 0100-01-01 0:0:0:000",
			args:    args{value: math.Float64bits(minOleDate)},
			want:    time.Date(100, 1, 1, 0, 0, 0, 0, time.UTC),
			wantErr: false,
		},
		{
			name:    "max OLE date 9999-12-31 23:59:59:999",
			args:    args{value: math.Float64bits(maxOleDate)},
			want:    time.Date(9999, 12, 31, 23, 59, 59, 999000000, time.UTC),
			wantErr: false,
		},
		{
			name:    "before min date",
			args:    args{value: math.Float64bits(minOleDate - 1)},
			want:    time.Time{},
			wantErr: true,
		},
		{
			name:    "after max date",
			args:    args{value: math.Float64bits(maxOleDate + 1)},
			want:    time.Time{},
			wantErr: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := GetVariantDate(tt.args.value)
			if (err != nil) != tt.wantErr {
				t.Errorf("GetVariantDate() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("GetVariantDate() got = %v, want %v", got, tt.want)
			}
		})
	}
}
