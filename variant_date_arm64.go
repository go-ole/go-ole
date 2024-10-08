//go:build windows && arm64
// +build windows,arm64

package ole

import (
	"errors"
	"math"
	"time"
)

// Constants representing the valid range of OLE Automation dates
const (
	// dates taken from microsoft docs
	// https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tooadate?view=net-8.0
	minOleDate float64 = -657434.0        // Represents January 1, 100
	maxOleDate float64 = 2958465.99999999 // Represents December 31, 9999
)

// oleStartTime represents the starting point of OLE date calculation (December 30, 1899)
var oleStartTime = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)

// GetVariantDate converts a uint64 OLE DATE-like value to a Go time.Time structure
func GetVariantDate(value uint64) (time.Time, error) {
	// Convert the uint64 back to a float64 (OLE DATE format)
	oleDateFloat := math.Float64frombits(value)

	// Check if the oleDate is within a valid range
	if oleDateFloat < minOleDate || oleDateFloat > maxOleDate {
		return time.Time{}, errors.New("invalid OLE date range")
	}

	// Separate the integer part (days) and the fractional part (time of day)
	days := int(oleDateFloat)
	fraction := oleDateFloat - float64(days)

	// Calculate the date by adding the integer part to the oleStartTime
	date := oleStartTime.AddDate(0, 0, days)

	// fractional component represents the time on that day divided by 24
	// so mulitplying by 24 gives the total time
	totalTime := fraction * 24

	hours := int(totalTime)
	timeUsedInHours := float64(hours)

	minutes := int((totalTime - timeUsedInHours) * 60)
	timeUsedInMinutes := timeUsedInHours*60 + float64(minutes)

	seconds := int((totalTime*60 - timeUsedInMinutes) * 60)
	timeUsedInSeconds := timeUsedInMinutes*60 + float64(seconds)

	nanoseconds := int((totalTime*60*60 - timeUsedInSeconds) * 1e9)

	// Construct the final time.Time object, rounded to the nearest millisecond
	return time.Date(date.Year(), date.Month(), date.Day(), hours, minutes, seconds, nanoseconds, time.UTC).Round(time.Millisecond), nil
}
