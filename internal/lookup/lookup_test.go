package lookup

import (
	"math"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/openact/kit/cache/v3"
)

func loadTestTable(t *testing.T, path string) *cache.Table {
	t.Helper()
	tbl, err := cache.LoadTable(path, cache.LoadOptions{})
	if err != nil {
		t.Fatal(err)
	}
	return tbl
}

func TestEReadResultUsesTableDimensionOrder(t *testing.T) {
	tbl := loadTestTable(t, filepath.Join("..", "..", "input", "testData", "rbc2512_4", "E_2", "TEST_ACCUM.fac"))

	got, ok := ERead_Result(tbl, "0", "TEST_ACCUM", "0", "DISC_SH_NET_BT_A", "202412")
	if !ok {
		t.Fatal("ERead_Result did not match existing coordinates")
	}

	const want = -10252884024.483696
	if math.Abs(got-want) > 0.000001 {
		t.Fatalf("ERead_Result = %v, want %v", got, want)
	}
}

func TestEReadResultRejectsWrongDimensionCount(t *testing.T) {
	tbl := loadTestTable(t, filepath.Join("..", "..", "input", "testData", "rbc2512_4", "E_2", "TEST_ACCUM.fac"))

	if _, ok := ERead_Result(tbl, "0", "TEST_ACCUM"); ok {
		t.Fatal("ERead_Result matched with too few coordinates")
	}
}

func TestEReadResultWithNineDimFac(t *testing.T) {
	path := filepath.Join(t.TempDir(), "nine.fac")
	data := "!9,CONT_GRP_CODE,PRD_CODE,RPT_DATE,MUTUALIZATION_ID,RUN_NAME,DSCNT_RATE_TYPE,VAR_NAME,VRS,AMT\n" +
		"*,1CNY09A0052021VFA03A2,A204,202201,2,NB,0,EXP_PH_TAX,1,111\n" +
		"*,1CNY09A0052021VFA03A2,A204,202201,2,NB,0,EXP_DIV_INV_UI,1,0\n"
	if err := os.WriteFile(path, []byte(data), 0o600); err != nil {
		t.Fatal(err)
	}
	tbl := loadTestTable(t, path)

	got, ok := ERead_Result(tbl, "1CNY09A0052021VFA03A2", "A204", "202201", "2", "NB", "0", "EXP_PH_TAX", "1", "AMT")
	if !ok {
		t.Fatal("ERead_Result did not match nine-dim coordinates")
	}
	if got != 111 {
		t.Fatalf("ERead_Result = %v, want 111", got)
	}
}

func TestEReadResultAllowsEmptyLeadingCoordinate(t *testing.T) {
	path := filepath.Join(t.TempDir(), "empty-leading.fac")
	data := "!3,DIM_A,DIM_B,AMT\n" +
		"*, ,B1,123\n"
	if err := os.WriteFile(path, []byte(data), 0o600); err != nil {
		t.Fatal(err)
	}
	tbl := loadTestTable(t, path)

	got, ok := ERead_Result(tbl, "", "B1", "AMT")
	if !ok {
		t.Fatal("ERead_Result did not match an empty leading coordinate")
	}
	if got != 123 {
		t.Fatalf("ERead_Result = %v, want 123", got)
	}
}

func TestEReadResultHandlesQuotedCommaKey(t *testing.T) {
	path := filepath.Join(t.TempDir(), "quoted-comma.fac")
	data := "!3,DIM_A,DIM_B,AMT\n" +
		`*,"PREM_INC(1,1)",B1,456` + "\n"
	if err := os.WriteFile(path, []byte(data), 0o600); err != nil {
		t.Fatal(err)
	}
	tbl := loadTestTable(t, path)

	got, ok := ERead_Result(tbl, "PREM_INC(1,1)", "B1", "AMT")
	if !ok {
		t.Fatal("ERead_Result did not match a quoted comma key")
	}
	if got != 456 {
		t.Fatalf("ERead_Result = %v, want 456", got)
	}
}

func TestEReadResultStatusDistinguishesZeroFromNotFound(t *testing.T) {
	tbl := loadTestTable(t, filepath.Join("..", "..", "input", "testData", "rbc2512_4", "E_2", "TEST_ACCUM.fac"))

	got, status := ERead_ResultStatus(tbl, "0", "TEST_ACCUM", "0", "SEG_CFL(1)", "202412")
	if status != StatusOK {
		t.Fatalf("ERead_ResultStatus status = %s, want OK", StatusText(status))
	}
	if got != 0 {
		t.Fatalf("ERead_ResultStatus = %v, want 0", got)
	}

	_, status = ERead_ResultStatus(tbl, "0", "TEST_ACCUM", "0", "NOT_A_VAR", "202412")
	if status != StatusNotFound {
		t.Fatalf("missing lookup status = %s, want not found", StatusText(status))
	}
}

func TestWriteEReadTable(t *testing.T) {
	tbl := loadTestTable(t, filepath.Join("..", "..", "input", "testData", "rbc2512_4", "E_2", "TEST_ACCUM.fac"))
	out := filepath.Join(t.TempDir(), "eread.tsv")
	rows := [][]string{
		{"0", "TEST_ACCUM", "0", "DISC_SH_NET_BT_A"},
		{"0", "TEST_ACCUM", "0", "SEG_CFL(1)"},
		{"0", "TEST_ACCUM", "0", "NOT_A_VAR"},
	}
	cols := []string{"202412", "2025"}

	if status := WriteEReadTable(tbl, rows, cols, out); status != StatusOK {
		t.Fatalf("WriteEReadTable status = %s, want OK", StatusText(status))
	}
	b, err := os.ReadFile(out)
	if err != nil {
		t.Fatal(err)
	}
	got := strings.TrimSpace(string(b))
	want := "-10252884024.483696\t-10678498686.329264\n0\t0\n#N/A\t#N/A"
	if got != want {
		t.Fatalf("WriteEReadTable output:\n%s\nwant:\n%s", got, want)
	}
}

func TestWriteEProjTable(t *testing.T) {
	tbl := loadTestTable(t, filepath.Join("..", "..", "input", "testData", "rbc2512_4", "E_2", "TEST_ACCUM.fac"))
	out := filepath.Join(t.TempDir(), "eproj.tsv")
	rows := [][]string{
		{"0", "DISC_SH_NET_BT_A"},
		{"0", "SEG_CFL(1)"},
	}
	cols := []string{"202412", "2025"}

	if status := WriteEProjTable(tbl, "TEST_ACCUM", "0", rows, cols, out); status != StatusOK {
		t.Fatalf("WriteEProjTable status = %s, want OK", StatusText(status))
	}
	b, err := os.ReadFile(out)
	if err != nil {
		t.Fatal(err)
	}
	got := strings.TrimSpace(string(b))
	want := "-10252884024.483696\t-10678498686.329264\n0\t0"
	if got != want {
		t.Fatalf("WriteEProjTable output:\n%s\nwant:\n%s", got, want)
	}
}
