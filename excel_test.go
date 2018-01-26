package excel

import (
	"os"
	"path/filepath"
	"testing"
)

type Test struct {
	A string
	B string
	C int
}

func getFilename(filename string) (string, error) {
	pwd, err := os.Getwd()
	if err != nil {
		return pwd, err
	}
	dir := filepath.Join(pwd, "data")
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		os.Mkdir(dir, os.ModePerm)
	}

	return filepath.Join(pwd, "data", filename), nil
}

func TestOverwrite(t *testing.T) {
	fp, err := getFilename("OverwriteTest.xlsx")
	if err != nil {
		t.Fatal("could not retrieve file path")
	}
	rows := []interface{}{
		Test{
			A: "A1",
			B: "B1",
			C: 1,
		},
		Test{
			A: "A2",
			B: "B2",
			C: 2,
		},
	}

	writer := New(fp, true)
	_, err = writer.WriteSheet("Test", []string{"A", "B", "C"}, rows)
	if err != nil {
		t.Fatal("could not create sheet", err)
	}
	err = writer.Save()
	if err != nil {
		t.Fatal("could not save workbook", err)
	}

	if _, err = os.Stat(fp); os.IsNotExist(err) {
		t.Fatal("file not created", err)
	}
}

func TestNoOverwrite(t *testing.T) {
	fp, err := getFilename("OverwriteTest.xlsx")
	if err != nil {
		t.Fatal("could not retrieve file path")
	}
	rows := []interface{}{
		&Test{
			A: "A1",
			B: "B1",
			C: 1,
		},
		&Test{
			A: "A2",
			B: "B2",
			C: 2,
		},
	}

	writer := New(fp, false)
	_, err = writer.WriteSheet("Test", []string{"A", "B", "C"}, rows)
	if err != nil {
		t.Fatal("could not create sheet", err)
	}
	err = writer.Save()
	if err == nil {
		t.Fatal("should have received an error")
	}
}
