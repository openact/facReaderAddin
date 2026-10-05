//go:build windows

package main

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/openact/facReaderAddin/internal/lookup"
	"github.com/openact/kit/cache/v3"
)

func TestGetTableMissingFileReturnsNil(t *testing.T) {
	path := filepath.Join(t.TempDir(), "missing.fac")
	if got := getTable(path); got.table != nil || got.status == 0 {
		t.Fatal("getTable returned a table for a missing file")
	}
}

func TestGetTableDirectoryReturnsNil(t *testing.T) {
	if got := getTable(t.TempDir()); got.table != nil || got.status == 0 {
		t.Fatal("getTable returned a table for a directory")
	}
}

func TestGetTableMalformedFileReturnsNil(t *testing.T) {
	path := filepath.Join(t.TempDir(), "bad.fac")
	if err := os.WriteFile(path, []byte("!3,DIM_A,DIM_B,AMT\n*,A,B,1\n*,A,B,2\n"), 0o600); err != nil {
		t.Fatal(err)
	}
	if got := getTable(path); got.table != nil || got.status != lookup.StatusLoadError {
		t.Fatalf("getTable malformed file = (%v, %d), want load error", got.table, got.status)
	}
	if detail := loadError(path); !strings.Contains(detail, "line 3: duplicate row key") {
		t.Fatalf("load error detail = %q, want line number and cause", detail)
	}
	if err := os.WriteFile(path, []byte("!3,DIM_A,DIM_B,AMT\n*,,B,0\n"), 0o600); err != nil {
		t.Fatal(err)
	}
	if got := getTable(path); got.table == nil || got.status != lookup.StatusOK {
		t.Fatalf("getTable corrected file = (%v, %d), want loaded table", got.table, got.status)
	}
	if detail := loadError(path); detail != "" {
		t.Fatalf("stale load error after successful reload: %q", detail)
	}
}

func TestIsRemotePath(t *testing.T) {
	if !isRemotePath(`\\server\share\result.fac`) {
		t.Fatal("UNC path was not detected as remote")
	}
	if isRemotePath(`C:\result\local.fac`) {
		t.Fatal("local drive path was detected as remote")
	}
}

func TestCopyFacToLocalCacheLeavesLocalPathUnchanged(t *testing.T) {
	path := filepath.Join(t.TempDir(), "local.fac")
	if err := os.WriteFile(path, []byte("!1,AMT\n*,123\n"), 0o600); err != nil {
		t.Fatal(err)
	}
	got, err := copyFacToLocalCache(path, 1, 1)
	if err != nil {
		t.Fatal(err)
	}
	if got != path {
		t.Fatalf("copyFacToLocalCache local path = %q, want %q", got, path)
	}
}

func TestGetTableCachedRemoteSkipsSourceStat(t *testing.T) {
	path := `\\server\share\missing-after-cache.fac`
	key := normPath(path)
	table := &cache.Table{}

	mu.Lock()
	tables[key] = cachedTable{table: table, modUnix: 1, size: 1, remote: true}
	mu.Unlock()
	defer func() {
		mu.Lock()
		delete(tables, key)
		mu.Unlock()
	}()

	got := getTable(path)
	if got.status != 0 || got.table != table {
		t.Fatalf("getTable cached remote = (%v, %d), want cached table and OK", got.table, got.status)
	}
}

func TestClearTableCache(t *testing.T) {
	key := normPath(`C:\result\cached.fac`)
	mu.Lock()
	tables[key] = cachedTable{table: &cache.Table{}, modUnix: 1, size: 1}
	mu.Unlock()

	if removed := clearTableCache(); removed == 0 {
		t.Fatal("clearTableCache removed 0 tables")
	}
	mu.RLock()
	_, ok := tables[key]
	mu.RUnlock()
	if ok {
		t.Fatal("clearTableCache left cached table behind")
	}
}

func TestRefreshChangedTableCacheKeepsUnchangedTable(t *testing.T) {
	clearTableCache()
	path := filepath.Join(t.TempDir(), "unchanged.fac")
	if err := os.WriteFile(path, []byte("same"), 0o600); err != nil {
		t.Fatal(err)
	}
	info, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	key := normPath(path)
	table := &cache.Table{}
	mu.Lock()
	tables[key] = cachedTable{table: table, sourcePath: path, modUnix: info.ModTime().UnixNano(), size: info.Size()}
	mu.Unlock()

	if removed := refreshChangedTableCache(); removed != 0 {
		t.Fatalf("refreshChangedTableCache removed %d unchanged tables", removed)
	}
	mu.RLock()
	got := tables[key].table
	mu.RUnlock()
	if got != table {
		t.Fatal("refreshChangedTableCache did not keep unchanged table")
	}
}

func TestRefreshChangedTableCacheRemovesChangedTable(t *testing.T) {
	clearTableCache()
	path := filepath.Join(t.TempDir(), "changed.fac")
	if err := os.WriteFile(path, []byte("new-size"), 0o600); err != nil {
		t.Fatal(err)
	}
	key := normPath(path)
	mu.Lock()
	tables[key] = cachedTable{table: &cache.Table{}, sourcePath: path, modUnix: 1, size: 1}
	mu.Unlock()

	if removed := refreshChangedTableCache(); removed != 1 {
		t.Fatalf("refreshChangedTableCache removed %d tables, want 1", removed)
	}
	mu.RLock()
	_, ok := tables[key]
	mu.RUnlock()
	if ok {
		t.Fatal("refreshChangedTableCache kept changed table")
	}
}

func TestRefreshChangedTableCacheKeepsTableOnStatFailure(t *testing.T) {
	clearTableCache()
	path := filepath.Join(t.TempDir(), "missing.fac")
	key := normPath(path)
	table := &cache.Table{}
	mu.Lock()
	tables[key] = cachedTable{table: table, sourcePath: path, modUnix: 1, size: 1}
	mu.Unlock()

	if removed := refreshChangedTableCache(); removed != 0 {
		t.Fatalf("refreshChangedTableCache removed %d tables after stat failure", removed)
	}
	mu.RLock()
	got := tables[key].table
	mu.RUnlock()
	if got != table {
		t.Fatal("refreshChangedTableCache did not keep table after stat failure")
	}
}

func TestWriteFacCacheMetadata(t *testing.T) {
	localPath := filepath.Join(t.TempDir(), "cached.fac")
	sourcePath := `\\server\share\run\E_1\TEST.fac`
	cachedAt := time.Date(2026, 8, 7, 11, 0, 0, 0, time.UTC)

	writeFacCacheMetadata(sourcePath, localPath, 123456789, 42, cachedAt)

	b, err := os.ReadFile(localPath + ".meta.json")
	if err != nil {
		t.Fatal(err)
	}
	var meta facCacheMetadata
	if err := json.Unmarshal(b, &meta); err != nil {
		t.Fatal(err)
	}
	if meta.SourcePath != filepath.Clean(sourcePath) {
		t.Fatalf("SourcePath = %q, want %q", meta.SourcePath, filepath.Clean(sourcePath))
	}
	if meta.NormalizedSourcePath != normPath(sourcePath) {
		t.Fatalf("NormalizedSourcePath = %q, want %q", meta.NormalizedSourcePath, normPath(sourcePath))
	}
	if meta.LocalPath != localPath {
		t.Fatalf("LocalPath = %q, want %q", meta.LocalPath, localPath)
	}
	if meta.SourceSize != 42 || meta.SourceModUnixNano != 123456789 {
		t.Fatalf("source metadata = size %d mod %d", meta.SourceSize, meta.SourceModUnixNano)
	}
	if meta.CachedAt == "" || meta.LastAccessAt == "" {
		t.Fatal("metadata timestamps were not written")
	}
}

func writeTestFacCacheEntry(t *testing.T, dir, name, source string, size int64, lastAccessAt time.Time) string {
	t.Helper()
	localPath := filepath.Join(dir, name)
	if err := os.WriteFile(localPath, make([]byte, size), 0o600); err != nil {
		t.Fatal(err)
	}
	meta := facCacheMetadata{
		Version:              1,
		SourcePath:           filepath.Clean(source),
		NormalizedSourcePath: normPath(source),
		LocalPath:            localPath,
		SourceSize:           size,
		SourceModUnixNano:    lastAccessAt.UnixNano(),
		CachedAt:             lastAccessAt.Format(time.RFC3339Nano),
		LastAccessAt:         lastAccessAt.Format(time.RFC3339Nano),
	}
	b, err := json.Marshal(meta)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(localPath+facCacheMetaSuffix, b, 0o600); err != nil {
		t.Fatal(err)
	}
	return localPath
}

func TestCleanupFacCacheRemovesOlderVersionsOfSameSource(t *testing.T) {
	clearTableCache()
	dir := t.TempDir()
	source := `\\server\share\run\E_1\Pivot_UC.fac`
	oldPath := writeTestFacCacheEntry(t, dir, "old.fac", source, 4, time.Date(2026, 8, 7, 10, 0, 0, 0, time.UTC))
	currentPath := writeTestFacCacheEntry(t, dir, "current.fac", source, 4, time.Date(2026, 8, 7, 11, 0, 0, 0, time.UTC))

	cleanupFacCache(dir, currentPath, normPath(source), maxFacCacheBytes)

	if _, err := os.Stat(oldPath); !os.IsNotExist(err) {
		t.Fatalf("old cache version still exists, stat err=%v", err)
	}
	if _, err := os.Stat(currentPath); err != nil {
		t.Fatalf("current cache version was removed: %v", err)
	}
}

func TestCleanupFacCacheUsesLRUCap(t *testing.T) {
	clearTableCache()
	dir := t.TempDir()
	oldPath := writeTestFacCacheEntry(t, dir, "old.fac", `\\server\share\old.fac`, 4, time.Date(2026, 8, 7, 10, 0, 0, 0, time.UTC))
	newPath := writeTestFacCacheEntry(t, dir, "new.fac", `\\server\share\new.fac`, 4, time.Date(2026, 8, 7, 11, 0, 0, 0, time.UTC))

	cleanupFacCache(dir, newPath, normPath(`\\server\share\new.fac`), 6)

	if _, err := os.Stat(oldPath); !os.IsNotExist(err) {
		t.Fatalf("LRU cache entry still exists, stat err=%v", err)
	}
	if _, err := os.Stat(newPath); err != nil {
		t.Fatalf("protected current cache was removed: %v", err)
	}
}

func TestCleanupFacCacheDoesNotTrustMetadataLocalPath(t *testing.T) {
	clearTableCache()
	dir := t.TempDir()
	outsidePath := filepath.Join(t.TempDir(), "outside.fac")
	if err := os.WriteFile(outsidePath, []byte("keep"), 0o600); err != nil {
		t.Fatal(err)
	}
	cachePath := filepath.Join(dir, "cache.fac")
	if err := os.WriteFile(cachePath, []byte("cache"), 0o600); err != nil {
		t.Fatal(err)
	}
	meta := facCacheMetadata{
		Version:              1,
		SourcePath:           `\\server\share\old.fac`,
		NormalizedSourcePath: normPath(`\\server\share\old.fac`),
		LocalPath:            outsidePath,
		SourceSize:           4,
		SourceModUnixNano:    1,
		CachedAt:             time.Now().UTC().Format(time.RFC3339Nano),
		LastAccessAt:         time.Now().UTC().Format(time.RFC3339Nano),
	}
	b, err := json.Marshal(meta)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(cachePath+facCacheMetaSuffix, b, 0o600); err != nil {
		t.Fatal(err)
	}

	cleanupFacCache(dir, filepath.Join(dir, "current.fac"), normPath(`\\server\share\current.fac`), 1)

	if _, err := os.Stat(outsidePath); err != nil {
		t.Fatalf("cleanup removed metadata LocalPath outside cache: %v", err)
	}
}
