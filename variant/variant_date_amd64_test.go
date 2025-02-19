//go:build windows && amd64
// +build windows,amd64

package variant

import (
	"errors"
	"math"
	"reflect"
	"syscall"
	"testing"
	"time"
	"unsafe"
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

func getVariantDateWithoutMillSeconds(value uint64) (time.Time, error) {
	var st syscall.Systemtime
	r, _, _ := procVariantTimeToSystemTime.Call(uintptr(value), uintptr(unsafe.Pointer(&st)))
	if r != 0 {
		return time.Date(int(st.Year), time.Month(st.Month), int(st.Day), int(st.Hour), int(st.Minute), int(st.Second), int(st.Milliseconds/1000), time.UTC), nil
	}
	return time.Now(), errors.New("Could not convert to time, passing current time.")
}

func TestGetVariantDateWithoutMillSeconds(t *testing.T) {
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
			want:    time.Date(2023, 10, 30, 23, 30, 30, 0, time.UTC),
			wantErr: false,
		},
		{
			name:    "2023-10-30 23:30:30:960",
			args:    args{value: math.Float64bits(45229.979525)},
			want:    time.Date(2023, 10, 30, 23, 30, 31, 0, time.UTC),
			wantErr: false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := getVariantDateWithoutMillSeconds(tt.args.value)
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
