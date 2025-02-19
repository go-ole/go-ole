//go:build windows
// +build windows

package ole

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func Test_safeArrayCreate(t *testing.T) {
	type args struct {
		variantType VT
		dimensions  uint32
		bounds      *SafeArrayBound
	}
	type want struct {
		Dimensions   uint16
		ElementsSize uint32
		LocksAmount  uint32
	}
	tests := []struct {
		name    string
		args    args
		want    want
		wantErr bool
	}{
		{
			name: "Create SafeArray",
			args: args{
				variantType: VT_UI1,
				dimensions:  uint32(1),
				bounds: &SafeArrayBound{
					Elements:   uint32(10000),
					LowerBound: int32(0),
				},
			},
			want: want{
				Dimensions:   uint16(1),
				ElementsSize: uint32(1),
				LocksAmount:  uint32(0),
			},
			wantErr: false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := safeArrayCreate(tt.args.variantType, tt.args.dimensions, tt.args.bounds)
			if (err != nil) != tt.wantErr {
				t.Errorf("safeArrayCreate() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !assert.Equal(t, tt.want.Dimensions, got.Dimensions) {
				t.Errorf("safeArrayCreate() Dimensions not equal: got.Dimensions = %v, tt.want.Dimensions %v", got.Dimensions, tt.want.Dimensions)
			}
			if !assert.Equal(t, tt.want.ElementsSize, got.ElementsSize) {
				t.Errorf("safeArrayCreate() ElementsSize not equal: got.ElementsSize = %v, tt.want.ElementsSize %v", got.ElementsSize, tt.want.ElementsSize)
			}
			if !assert.Equal(t, tt.want.LocksAmount, got.LocksAmount) {
				t.Errorf("safeArrayCreate() LocksAmount not equal: got.LocksAmount = %v, tt.want.LocksAmount %v", got.LocksAmount, tt.want.LocksAmount)
			}
			safeArrayDestroy(got)
		})
	}
}
