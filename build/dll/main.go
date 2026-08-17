//go:build windows

// Package main builds a Windows DLL (facReaderAddin.dll) exposing .fac lookup
// functions for Excel VBA.
//
// Build command (run from this directory):
//
//	go build -buildmode=c-shared -o facReaderAddin.dll .
//
// See facReaderAddin.bas for usage from Excel macros / UDFs.
package main

/*
#include <stdlib.h>
#include <string.h>
*/
import "C"

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"syscall"
	"time"
	"unsafe"

	"github.com/openact/facReaderAddin/internal/lookup"
	"github.com/openact/kit/cache/v3"
)

const (
	argSeparator       = "\x1f"
	rowSeparator       = "\x1e"
	maxFacCacheBytes   = int64(100) * 1024 * 1024 * 1024
	facCacheMetaSuffix = ".meta.json"
)

var (
	mu                sync.RWMutex
	tables            = make(map[string]cachedTable)
	kernel32          = syscall.NewLazyDLL("kernel32.dll")
	procGetDriveTypeW = kernel32.NewProc("GetDriveTypeW")
)

type cachedTable struct {
	table      *cache.Table
	sourcePath string
	modUnix    int64
	size       int64
	remote     bool
	local      string
}

type tableResult struct {
	table  *cache.Table
	status int
}

type facCacheMetadata struct {
	Version              int    `json:"version"`
	SourcePath           string `json:"sourcePath"`
	NormalizedSourcePath string `json:"normalizedSourcePath"`
	LocalPath            string `json:"localPath"`
	SourceSize           int64  `json:"sourceSize"`
	SourceModUnixNano    int64  `json:"sourceModUnixNano"`
	CachedAt             string `json:"cachedAt"`
	LastAccessAt         string `json:"lastAccessAt"`
}

type facCacheEntry struct {
	localPath            string
	metaPath             string
	normalizedSourcePath string
	size                 int64
	lastAccessAt         time.Time
}

type cachedTableSnapshot struct {
	key        string
	sourcePath string
	modUnix    int64
	size       int64
}

func normPath(path string) string {
	return strings.ToLower(filepath.Clean(path))
}

func isRemotePath(path string) bool {
	vol := filepath.VolumeName(path)
	if strings.HasPrefix(vol, `\\`) || strings.HasPrefix(vol, `//`) {
		return true
	}
	if len(vol) == 2 && vol[1] == ':' {
		root := vol + `\`
		ptr, err := syscall.UTF16PtrFromString(root)
		if err != nil {
			return false
		}
		const driveRemote = 4
		driveType, _, _ := procGetDriveTypeW.Call(uintptr(unsafe.Pointer(ptr)))
		return driveType == driveRemote
	}
	return false
}

func localFacCacheDir() (string, error) {
	root, err := os.UserCacheDir()
	if err != nil || root == "" {
		root = os.TempDir()
	}
	dir := filepath.Join(root, "facReaderAddin", "fac-cache")
	if err := os.MkdirAll(dir, 0o700); err != nil {
		return "", err
	}
	return dir, nil
}

func localFacCachePath(path string, modUnix int64, size int64) (string, error) {
	dir, err := localFacCacheDir()
	if err != nil {
		return "", err
	}
	sum := sha256.Sum256([]byte(normPath(path)))
	name := hex.EncodeToString(sum[:]) + "_" + strconv.FormatInt(size, 10) + "_" + strconv.FormatInt(modUnix, 10) + ".fac"
	return filepath.Join(dir, name), nil
}

func writeFacCacheMetadata(sourcePath, localPath string, modUnix int64, size int64, cachedAt time.Time) {
	now := time.Now().UTC()
	if cachedAt.IsZero() {
		cachedAt = now
	}
	meta := facCacheMetadata{
		Version:              1,
		SourcePath:           filepath.Clean(sourcePath),
		NormalizedSourcePath: normPath(sourcePath),
		LocalPath:            localPath,
		SourceSize:           size,
		SourceModUnixNano:    modUnix,
		CachedAt:             cachedAt.UTC().Format(time.RFC3339Nano),
		LastAccessAt:         now.Format(time.RFC3339Nano),
	}
	b, err := json.MarshalIndent(meta, "", "  ")
	if err != nil {
		return
	}
	_ = os.WriteFile(localPath+facCacheMetaSuffix, b, 0o600)
}

func readFacCacheEntries(dir string) []facCacheEntry {
	files, err := os.ReadDir(dir)
	if err != nil {
		return nil
	}
	entries := make([]facCacheEntry, 0, len(files)/2)
	for _, file := range files {
		if file.IsDir() || !strings.HasSuffix(file.Name(), ".fac"+facCacheMetaSuffix) {
			continue
		}
		metaPath := filepath.Join(dir, file.Name())
		b, err := os.ReadFile(metaPath)
		if err != nil {
			continue
		}
		var meta facCacheMetadata
		if err := json.Unmarshal(b, &meta); err != nil || meta.LocalPath == "" {
			continue
		}
		localPath := strings.TrimSuffix(metaPath, facCacheMetaSuffix)
		if filepath.Dir(localPath) != dir {
			continue
		}
		info, err := os.Stat(localPath)
		if err != nil || info.IsDir() {
			continue
		}
		lastAccessAt, err := time.Parse(time.RFC3339Nano, meta.LastAccessAt)
		if err != nil {
			lastAccessAt = info.ModTime()
		}
		entries = append(entries, facCacheEntry{
			localPath:            localPath,
			metaPath:             metaPath,
			normalizedSourcePath: meta.NormalizedSourcePath,
			size:                 info.Size(),
			lastAccessAt:         lastAccessAt,
		})
	}
	return entries
}

func removeFacCacheEntry(entry facCacheEntry) {
	_ = os.Remove(entry.localPath)
	_ = os.Remove(entry.metaPath)
}

func cleanupFacCache(dir, currentLocalPath, normalizedSourcePath string, maxBytes int64) {
	currentLocalPath = filepath.Clean(currentLocalPath)
	entries := readFacCacheEntries(dir)
	protected := map[string]struct{}{normPath(currentLocalPath): {}}
	mu.RLock()
	for _, entry := range tables {
		if entry.local != "" {
			protected[normPath(entry.local)] = struct{}{}
		}
	}
	mu.RUnlock()

	kept := entries[:0]
	for _, entry := range entries {
		if _, ok := protected[normPath(entry.localPath)]; ok {
			kept = append(kept, entry)
			continue
		}
		if entry.normalizedSourcePath == normalizedSourcePath {
			removeFacCacheEntry(entry)
			continue
		}
		kept = append(kept, entry)
	}

	if maxBytes <= 0 {
		return
	}
	var total int64
	for _, entry := range kept {
		total += entry.size
	}
	if total <= maxBytes {
		return
	}
	sort.Slice(kept, func(i, j int) bool {
		return kept[i].lastAccessAt.Before(kept[j].lastAccessAt)
	})
	for _, entry := range kept {
		if total <= maxBytes {
			return
		}
		if _, ok := protected[normPath(entry.localPath)]; ok {
			continue
		}
		removeFacCacheEntry(entry)
		total -= entry.size
	}
}

func copyFacToLocalCache(path string, modUnix int64, size int64) (string, error) {
	if !isRemotePath(path) {
		return path, nil
	}
	localPath, err := localFacCachePath(path, modUnix, size)
	if err != nil {
		return "", err
	}
	if info, err := os.Stat(localPath); err == nil && !info.IsDir() && info.Size() == size {
		writeFacCacheMetadata(path, localPath, modUnix, size, info.ModTime())
		return localPath, nil
	}

	src, err := os.Open(path)
	if err != nil {
		return "", err
	}
	defer src.Close()

	tmp := localPath + ".tmp-" + strconv.Itoa(os.Getpid())
	dst, err := os.Create(tmp)
	if err != nil {
		return "", err
	}
	buf := make([]byte, 4<<20)
	written, copyErr := io.CopyBuffer(dst, src, buf)
	closeErr := dst.Close()
	if copyErr != nil {
		_ = os.Remove(tmp)
		return "", copyErr
	}
	if closeErr != nil {
		_ = os.Remove(tmp)
		return "", closeErr
	}
	if written != size {
		_ = os.Remove(tmp)
		return "", fmt.Errorf("copied %d bytes, expected %d", written, size)
	}
	ts := time.Unix(0, modUnix)
	_ = os.Chtimes(tmp, ts, ts)
	_ = os.Remove(localPath)
	if err := os.Rename(tmp, localPath); err != nil {
		_ = os.Remove(tmp)
		return "", err
	}
	writeFacCacheMetadata(path, localPath, modUnix, size, time.Now().UTC())
	return localPath, nil
}

func getTable(path string) tableResult {
	key := normPath(path)

	mu.RLock()
	entry, ok := tables[key]
	if ok && entry.remote {
		table := entry.table
		mu.RUnlock()
		return tableResult{table: table, status: lookup.StatusOK}
	}
	mu.RUnlock()

	remote := isRemotePath(path)
	info, err := os.Stat(path)
	if err != nil {
		return tableResult{status: lookup.StatusFileNotFound}
	}
	if info.IsDir() {
		return tableResult{status: lookup.StatusFileNotFound}
	}
	modUnix := info.ModTime().UnixNano()
	size := info.Size()

	mu.RLock()
	entry, ok = tables[key]
	mu.RUnlock()
	if ok && entry.modUnix == modUnix && entry.size == size {
		return tableResult{table: entry.table, status: lookup.StatusOK}
	}

	mu.Lock()
	if entry, ok = tables[key]; ok && entry.modUnix == modUnix && entry.size == size {
		mu.Unlock()
		return tableResult{table: entry.table, status: lookup.StatusOK}
	}
	loadPath, err := copyFacToLocalCache(path, modUnix, size)
	if err != nil {
		mu.Unlock()
		return tableResult{status: lookup.StatusInvalid}
	}
	loaded, err := cache.LoadTableSafe(loadPath)
	if err != nil || loaded == nil {
		mu.Unlock()
		return tableResult{status: lookup.StatusInvalid}
	}
	tables[key] = cachedTable{table: loaded, sourcePath: path, modUnix: modUnix, size: size, remote: remote, local: loadPath}
	mu.Unlock()
	if remote {
		cleanupFacCache(filepath.Dir(loadPath), loadPath, normPath(path), maxFacCacheBytes)
	}
	return tableResult{table: loaded, status: lookup.StatusOK}
}

// ─── exported functions ───────────────────────────────────────────────────────

func cstr(s *C.char) string {
	if s == nil {
		return ""
	}
	return C.GoString(s)
}

// EProj_Result returns one matched value and a lookup status code.
//
//export EProj_Result
func EProj_Result(filePath, spCode, resultType, variable, timePeriod, simID *C.char, valueOut *C.double) (status C.int) {
	defer func() {
		if recover() != nil {
			status = C.int(lookup.StatusInvalid)
		}
	}()
	_ = resultType // reserved parameter; intentionally not used.
	path := cstr(filePath)
	tr := getTable(path)
	if tr.table == nil {
		return C.int(tr.status)
	}
	product := lookup.ProductFromFilePath(path)
	v, resultStatus := lookup.EProj_ResultStatus(tr.table, product, cstr(spCode), cstr(variable), cstr(timePeriod), cstr(simID))
	if resultStatus == lookup.StatusOK && valueOut != nil {
		*valueOut = C.double(v)
	}
	return C.int(resultStatus)
}

// ERead_Result returns one matched value using direct string arguments.
//
//export ERead_Result
func ERead_Result(filePath *C.char, coordCount C.int, k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12 *C.char, valueOut *C.double) (status C.int) {
	defer func() {
		if recover() != nil {
			status = C.int(lookup.StatusInvalid)
		}
	}()
	n := int(coordCount)
	if n < 0 || n > 12 {
		return C.int(lookup.StatusDimMismatch)
	}

	raw := [12]*C.char{k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12}
	args := make([]string, n)
	for i := 0; i < n; i++ {
		args[i] = cstr(raw[i])
	}
	tr := getTable(cstr(filePath))
	if tr.table == nil {
		return C.int(tr.status)
	}
	v, resultStatus := lookup.ERead_ResultStatus(tr.table, args...)
	if resultStatus == lookup.StatusOK && valueOut != nil {
		*valueOut = C.double(v)
	}
	return C.int(resultStatus)
}

// ERead_Batch writes a TSV matrix for row-coordinate matrix x column-key vector.
//
//export ERead_Batch
func ERead_Batch(filePath, rowKeysPacked, colKeysPacked, outputPath *C.char) (status C.int) {
	defer func() {
		if recover() != nil {
			status = C.int(lookup.StatusInvalid)
		}
	}()
	tr := getTable(cstr(filePath))
	if tr.table == nil {
		return C.int(tr.status)
	}
	rows := unpackRows(cstr(rowKeysPacked))
	cols := unpackFields(cstr(colKeysPacked))
	resultStatus := lookup.WriteEReadTable(tr.table, rows, cols, cstr(outputPath))
	return C.int(resultStatus)
}

// EProj_Batch writes a TSV matrix for SP_CODE/VAR_NAME rows x time-period columns.
//
//export EProj_Batch
func EProj_Batch(filePath, product, simID, projKeysPacked, colKeysPacked, outputPath *C.char) (status C.int) {
	defer func() {
		if recover() != nil {
			status = C.int(lookup.StatusInvalid)
		}
	}()
	_ = product // product key follows EProj_Result: derive it from the .fac filename.
	path := cstr(filePath)
	tr := getTable(path)
	if tr.table == nil {
		return C.int(tr.status)
	}
	rows := unpackRows(cstr(projKeysPacked))
	cols := unpackFields(cstr(colKeysPacked))
	resultStatus := lookup.WriteEProjTable(tr.table, lookup.ProductFromFilePath(path), cstr(simID), rows, cols, cstr(outputPath))
	return C.int(resultStatus)
}

// FacNumDims returns the number of lookup coordinates required by a .fac file.
//
//export FacNumDims
func FacNumDims(filePath *C.char) (status C.int) {
	defer func() {
		if recover() != nil {
			status = 0
		}
	}()
	tr := getTable(cstr(filePath))
	if tr.table == nil {
		return 0
	}
	return C.int(tr.table.NumDims)
}

// FacRelease removes one table from in-memory cache.
// Returns 1 if table existed and was removed; otherwise 0.
//
//export FacRelease
func FacRelease(filePath *C.char) (status C.int) {
	defer func() {
		if recover() != nil {
			status = 0
		}
	}()
	key := normPath(cstr(filePath))
	mu.Lock()
	defer mu.Unlock()
	if _, ok := tables[key]; ok {
		delete(tables, key)
		return 1
	}
	return 0
}

func clearTableCache() int {
	mu.Lock()
	defer mu.Unlock()
	n := len(tables)
	tables = make(map[string]cachedTable)
	return n
}

func refreshChangedTableCache() int {
	mu.RLock()
	snapshots := make([]cachedTableSnapshot, 0, len(tables))
	for key, entry := range tables {
		sourcePath := entry.sourcePath
		if sourcePath == "" {
			sourcePath = key
		}
		snapshots = append(snapshots, cachedTableSnapshot{
			key:        key,
			sourcePath: sourcePath,
			modUnix:    entry.modUnix,
			size:       entry.size,
		})
	}
	mu.RUnlock()

	changed := make([]cachedTableSnapshot, 0)
	for _, snapshot := range snapshots {
		info, err := os.Stat(snapshot.sourcePath)
		if err != nil || info.IsDir() {
			continue
		}
		if info.ModTime().UnixNano() != snapshot.modUnix || info.Size() != snapshot.size {
			changed = append(changed, snapshot)
		}
	}
	if len(changed) == 0 {
		return 0
	}

	removed := 0
	mu.Lock()
	defer mu.Unlock()
	for _, snapshot := range changed {
		entry, ok := tables[snapshot.key]
		if !ok {
			continue
		}
		if entry.modUnix == snapshot.modUnix && entry.size == snapshot.size {
			delete(tables, snapshot.key)
			removed++
		}
	}
	return removed
}

// FacClearCache removes all tables from in-memory cache.
// Returns the number of tables removed.
//
//export FacClearCache
func FacClearCache() (status C.int) {
	defer func() {
		if recover() != nil {
			status = 0
		}
	}()
	return C.int(clearTableCache())
}

// FacRefreshChangedCache removes cached tables whose source file size or
// modification time changed. Stat failures keep the old cache.
// Returns the number of tables removed.
//
//export FacRefreshChangedCache
func FacRefreshChangedCache() (status C.int) {
	defer func() {
		if recover() != nil {
			status = 0
		}
	}()
	return C.int(refreshChangedTableCache())
}

func unpackRows(packed string) [][]string {
	if packed == "" {
		return nil
	}
	parts := strings.Split(packed, rowSeparator)
	rows := make([][]string, len(parts))
	for i, part := range parts {
		rows[i] = unpackFields(part)
	}
	return rows
}

func unpackFields(packed string) []string {
	return strings.Split(packed, argSeparator)
}

func main() {} // required for -buildmode=c-shared
