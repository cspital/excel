package excel

import (
	"os"
	"reflect"

	"github.com/pkg/errors"
	"github.com/tealeg/xlsx"
)

var errFileAlreadyExists = errors.New("file already exists and overwrite is not enabled")

// XLWriter ...
// Encapsulates supported abstractions over github.com/tealeg/xlsx.
type XLWriter interface {
	WriteSheet(name string, headers []string, rows []interface{}) (*xlsx.Sheet, error)
	Save() error
}

// New ...
// Returns a new XLWriter.
func New(filename string, overwrite bool) XLWriter {
	return &xlWriter{
		filename:  filename,
		overwrite: overwrite,
		wkbk:      xlsx.NewFile(),
	}
}

func delete(filename string, overwrite bool) error {
	if overwrite {
		if _, err := os.Stat(filename); os.IsNotExist(err) {
			return nil
		}
		return os.Remove(filename)
	}

	if _, err := os.Stat(filename); err == nil {
		return errFileAlreadyExists
	}
	return nil
}

type xlWriter struct {
	filename  string
	overwrite bool
	wkbk      *xlsx.File
}

func (x *xlWriter) WriteSheet(name string, headers []string, rows []interface{}) (*xlsx.Sheet, error) {
	sheet, err := x.wkbk.AddSheet(name)
	if err != nil {
		return sheet, err
	}

	x.writeHeaders(sheet, headers)

	x.writeData(sheet, rows)

	return sheet, nil
}

func (x *xlWriter) Save() error {
	if err := delete(x.filename, x.overwrite); err != nil {
		return err
	}
	return x.wkbk.Save(x.filename)
}

func (x *xlWriter) writeHeaders(sheet *xlsx.Sheet, headers []string) {
	row := sheet.AddRow()
	for _, h := range headers {
		cell := row.AddCell()
		cell.Value = h

		if style := cell.GetStyle(); style != nil {
			style.Border = *xlsx.NewBorder("none", "none", "none", "thin")
			style.ApplyBorder = true
			style.Font.Bold = true
			style.ApplyFont = true
		}
	}
}

func (x *xlWriter) writeData(sheet *xlsx.Sheet, rows []interface{}) {
	for _, r := range rows {
		row := sheet.AddRow()
		val := getValue(r)

		for n := 0; n < val.NumField(); n++ {
			row.AddCell().SetValue(val.Field(n).Interface())
		}
	}
}

func getValue(item interface{}) reflect.Value {
	raw := reflect.ValueOf(item)

	if raw.Kind() == reflect.Ptr {
		return reflect.Indirect(raw)
	}

	return raw
}
