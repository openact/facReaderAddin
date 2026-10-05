#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <stdarg.h>

typedef int RW;
typedef int COL;
typedef unsigned short WORD16;
typedef struct xloper XLOPER;
typedef XLOPER *LPXLOPER;

typedef struct xlref {
    WORD16 rwFirst;
    WORD16 rwLast;
    BYTE colFirst;
    BYTE colLast;
} XLREF;

typedef struct xlmref {
    WORD16 count;
    XLREF reftbl[1];
} XLMREF;

struct xloper {
    union {
        double num;
        char *str;
        int xbool;
        int err;
        int w;
        struct {
            XLMREF *lpmref;
            int idSheet;
        } mref;
        struct {
            WORD16 rows;
            WORD16 columns;
            LPXLOPER lparray;
        } array;
        struct {
            WORD16 count;
            XLREF ref;
        } sref;
    } val;
    unsigned int xltype;
};

#define xltypeNum      0x0001
#define xltypeStr      0x0002
#define xltypeBool     0x0004
#define xltypeErr      0x0010
#define xltypeMissing  0x0080
#define xltypeNil      0x0100
#define xltypeInt      0x0800

#define xlerrValue 15
#define xlerrNA    42

#define xlfRegister 149

typedef int (__stdcall *Excel4vFn)(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[]);
typedef int (__cdecl *EProjResultFn)(char *filePath, char *spCode, char *resultType, char *variable, char *timePeriod, char *simID, double *valueOut);
typedef int (__cdecl *EReadResultFn)(char *filePath, int coordCount,
    char *k1, char *k2, char *k3, char *k4, char *k5, char *k6,
    char *k7, char *k8, char *k9, char *k10, char *k11, char *k12,
    double *valueOut);
typedef int (__cdecl *FacNumDimsFn)(char *filePath);
typedef int (__cdecl *FacLoadErrorFn)(char *filePath, char *buffer, int capacity);

static HMODULE g_module = NULL;
static HMODULE g_appReader = NULL;
static EProjResultFn g_projResult = NULL;
static EReadResultFn g_readResult = NULL;
static FacNumDimsFn g_facNumDims = NULL;
static FacLoadErrorFn g_facLoadError = NULL;

static __thread XLOPER g_result;
static __thread char g_resultText[256];

BOOL WINAPI DllMain(HINSTANCE hinst, DWORD reason, LPVOID reserved) {
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        g_module = (HMODULE)hinst;
    }
    return TRUE;
}

static void clear_result(void) {
    ZeroMemory(&g_result, sizeof(g_result));
    ZeroMemory(g_resultText, sizeof(g_resultText));
}

static LPXLOPER ret_num(double v) {
    clear_result();
    g_result.xltype = xltypeNum;
    g_result.val.num = v;
    return &g_result;
}

static LPXLOPER ret_text(const char *text) {
    clear_result();
    size_t n = strlen(text);
    if (n > 254) n = 254;
    g_resultText[0] = (char)n;
    memcpy(g_resultText + 1, text, n);
    g_result.xltype = xltypeStr;
    g_result.val.str = g_resultText;
    return &g_result;
}

static const char *status_text(int status) {
    switch (status) {
    case 0: return "OK";
    case 1: return "lookup coordinates not found";
    case 2: return "dimension mismatch";
    case 3: return "invalid argument";
    case 4: return "numeric parse error";
    case 5: return "failed to write batch output";
    case 6: return "file not found";
    case 7: return "FAC load error";
    case 8: return "wildcard is not allowed";
    default: return "unknown status";
    }
}

static LPXLOPER read_error_result(const char *path, int expectedDims, int gotDims, int status) {
    char msg[768];
    switch (status) {
    case 1:
        snprintf(msg, sizeof(msg), "#ERROR: no match found in %s.", path);
        break;
    case 2:
        if (expectedDims > 0) {
            snprintf(msg, sizeof(msg), "#ERROR: expected %d lookup coordinates, got %d for %s.", expectedDims, gotDims, path);
        } else {
            snprintf(msg, sizeof(msg), "#ERROR: dimension mismatch for %s (got %d coordinates).", path, gotDims);
        }
        break;
    case 3:
        snprintf(msg, sizeof(msg), "#ERROR: invalid argument for %s.", path);
        break;
    case 4:
        snprintf(msg, sizeof(msg), "#ERROR: numeric parse error in %s.", path);
        break;
    case 6:
        snprintf(msg, sizeof(msg), "#ERROR: file not found: %s.", path);
        break;
    case 7: {
        char detail[768] = {0};
        g_facLoadError((char *)path, detail, sizeof(detail));
        if (detail[0]) snprintf(msg, sizeof(msg), "#ERROR: %s", detail);
        else snprintf(msg, sizeof(msg), "#ERROR: failed to load FAC %s.", path);
        break;
    }
    case 8:
        snprintf(msg, sizeof(msg), "#ERROR: invalid argument for %s (wildcard \"*\" is not allowed).", path);
        break;
    default:
        snprintf(msg, sizeof(msg), "#ERROR: %s (%d) for %s.", status_text(status), status, path);
        break;
    }
    return ret_text(msg);
}

static LPXLOPER proj_error_result(const char *path, int status, const char *spCode, const char *variable, const char *timePeriod, const char *simID) {
    (void)spCode;
    (void)variable;
    (void)timePeriod;
    (void)simID;
    char msg[768];
    switch (status) {
    case 1:
        snprintf(msg, sizeof(msg), "#ERROR: no match found in %s.", path);
        break;
    case 2:
        snprintf(msg, sizeof(msg), "#ERROR: dimension mismatch for %s.", path);
        break;
    case 3:
        snprintf(msg, sizeof(msg), "#ERROR: invalid argument for %s.", path);
        break;
    case 4:
        snprintf(msg, sizeof(msg), "#ERROR: numeric parse error in %s.", path);
        break;
    case 6:
        snprintf(msg, sizeof(msg), "#ERROR: file not found: %s.", path);
        break;
    case 7: {
        char detail[768] = {0};
        g_facLoadError((char *)path, detail, sizeof(detail));
        if (detail[0]) snprintf(msg, sizeof(msg), "#ERROR: %s", detail);
        else snprintf(msg, sizeof(msg), "#ERROR: failed to load FAC %s.", path);
        break;
    }
    default:
        snprintf(msg, sizeof(msg), "#ERROR: %s (%d) for %s.", status_text(status), status, path);
        break;
    }
    return ret_text(msg);
}

static int is_missing(LPXLOPER x) {
    if (x == NULL) return 1;
    unsigned int t = x->xltype & 0x0FFF;
    return t == xltypeMissing || t == xltypeNil;
}

static char *xl_to_utf8(LPXLOPER x);

static char *ansi_to_utf8(const char *s, int len) {
    int wchars = MultiByteToWideChar(CP_ACP, 0, s, len, NULL, 0);
    if (wchars <= 0) return _strdup("");
    wchar_t *wide = (wchar_t *)calloc((size_t)wchars, sizeof(wchar_t));
    if (!wide) return NULL;
    MultiByteToWideChar(CP_ACP, 0, s, len, wide, wchars);

    int bytes = WideCharToMultiByte(CP_UTF8, 0, wide, wchars, NULL, 0, NULL, NULL);
    if (bytes <= 0) {
        free(wide);
        return _strdup("");
    }
    char *out = (char *)calloc((size_t)bytes + 1, 1);
    if (!out) {
        free(wide);
        return NULL;
    }
    WideCharToMultiByte(CP_UTF8, 0, wide, wchars, out, bytes, NULL, NULL);
    out[bytes] = 0;
    free(wide);
    return out;
}

static void civil_from_days(long long z, int *y, int *m, int *d) {
    z += 719468;
    long long era = (z >= 0 ? z : z - 146096) / 146097;
    unsigned doe = (unsigned)(z - era * 146097);
    unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    long long yy = (long long)yoe + era * 400;
    unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    unsigned mp = (5 * doy + 2) / 153;
    *d = (int)(doy - (153 * mp + 2) / 5 + 1);
    *m = (int)(mp + (mp < 10 ? 3 : -9));
    *y = (int)(yy + (*m <= 2));
}

static char *period_to_utf8(LPXLOPER x) {
    if (is_missing(x)) return _strdup("");
    unsigned int t = x->xltype & 0x0FFF;
    if (t != xltypeNum) return xl_to_utf8(x);

    double n = x->val.num;
    long long rounded = (long long)(n >= 0 ? n + 0.5 : n - 0.5);
    double diff = n - (double)rounded;
    if (diff < 0) diff = -diff;

    char buf[16];
    if (diff < 0.0000001 && ((rounded >= 190001 && rounded <= 299912) || (rounded >= 1900 && rounded <= 2999))) {
        snprintf(buf, sizeof(buf), "%lld", rounded);
        return _strdup(buf);
    }

    long long days = (long long)n;
    int y = 0, m = 0, d = 0;
    civil_from_days(days - 25569, &y, &m, &d);
    (void)d;
    snprintf(buf, sizeof(buf), "%04d%02d", y, m);
    return _strdup(buf);
}

static char *xl_to_utf8(LPXLOPER x) {
    if (is_missing(x)) return _strdup("");
    unsigned int t = x->xltype & 0x0FFF;
    char buf[64];
    switch (t) {
    case xltypeStr:
        if (!x->val.str) return _strdup("");
        return ansi_to_utf8(x->val.str + 1, (unsigned char)x->val.str[0]);
    case xltypeNum:
        snprintf(buf, sizeof(buf), "%.15g", x->val.num);
        return _strdup(buf);
    case xltypeInt:
        snprintf(buf, sizeof(buf), "%d", x->val.w);
        return _strdup(buf);
    case xltypeBool:
        return _strdup(x->val.xbool ? "TRUE" : "FALSE");
    default:
        return _strdup("");
    }
}

static int ends_with_fac(const char *s) {
    size_t n = strlen(s);
    if (n < 4) return 0;
    const char *p = s + n - 4;
    return _stricmp(p, ".fac") == 0;
}

static char *build_fac_path(const char *vault, const char *resID, const char *fileName) {
    size_t nv = strlen(vault), nr = strlen(resID), nf = strlen(fileName);
    int needSlash1 = nv > 0 && vault[nv - 1] != '\\' && vault[nv - 1] != '/';
    int needFac = !ends_with_fac(fileName);
    size_t total = nv + (needSlash1 ? 1 : 0) + nr + 1 + nf + (needFac ? 4 : 0) + 1;
    char *out = (char *)calloc(total, 1);
    if (!out) return NULL;
    strcat(out, vault);
    if (needSlash1) strcat(out, "\\");
    strcat(out, resID);
    strcat(out, "\\");
    strcat(out, fileName);
    if (needFac) strcat(out, ".fac");
    return out;
}

static int load_app_reader(void) {
    if (g_readResult && g_projResult && g_facNumDims && g_facLoadError) return 1;

    wchar_t path[MAX_PATH];
    DWORD n = GetModuleFileNameW(g_module, path, MAX_PATH);
    if (n == 0 || n >= MAX_PATH) return 0;
    wchar_t *slash = wcsrchr(path, L'\\');
    if (!slash) return 0;
    wcscpy(slash + 1, L"facReaderAddin.dll");

    g_appReader = LoadLibraryW(path);
    if (!g_appReader) return 0;

    g_readResult = (EReadResultFn)GetProcAddress(g_appReader, "ERead_Result");
    g_projResult = (EProjResultFn)GetProcAddress(g_appReader, "EProj_Result");
    g_facNumDims = (FacNumDimsFn)GetProcAddress(g_appReader, "FacNumDims");
    g_facLoadError = (FacLoadErrorFn)GetProcAddress(g_appReader, "FacLoadError");
    return g_readResult && g_projResult && g_facNumDims && g_facLoadError;
}

static XLOPER xl4_str(const char *s) {
    XLOPER x;
    ZeroMemory(&x, sizeof(x));
    size_t n = strlen(s);
    if (n > 255) n = 255;
    char *buf = (char *)calloc(n + 2, sizeof(char));
    if (!buf) return x;
    buf[0] = (char)n;
    memcpy(buf + 1, s, n);
    x.xltype = xltypeStr;
    x.val.str = buf;
    return x;
}

static XLOPER xl4_str_wide(const wchar_t *s) {
    XLOPER x;
    ZeroMemory(&x, sizeof(x));
    int bytes = WideCharToMultiByte(CP_ACP, 0, s, -1, NULL, 0, NULL, NULL);
    if (bytes <= 1) return x;
    char *tmp = (char *)calloc((size_t)bytes, sizeof(char));
    if (!tmp) return x;
    WideCharToMultiByte(CP_ACP, 0, s, -1, tmp, bytes, NULL, NULL);
    x = xl4_str(tmp);
    free(tmp);
    return x;
}

static XLOPER xl4_num(double n) {
    XLOPER x;
    ZeroMemory(&x, sizeof(x));
    x.xltype = xltypeNum;
    x.val.num = n;
    return x;
}

static void free_xl4_str(XLOPER *x) {
    if ((x->xltype & 0x0FFF) == xltypeStr && x->val.str) {
        free(x->val.str);
        x->val.str = NULL;
    }
}

static Excel4vFn excel4v(void) {
    HMODULE xlcall = GetModuleHandleW(L"XLCALL32.DLL");
    if (!xlcall) xlcall = LoadLibraryW(L"XLCALL32.DLL");
    if (!xlcall) return NULL;

    return (Excel4vFn)GetProcAddress(xlcall, "Excel4v");
}

static int register_one(Excel4vFn excel, LPXLOPER moduleName, const char *proc, const char *typeText,
    const char *funcText, const char *argText, const char *help, const char **argHelp, int argHelpCount) {
    XLOPER procX = xl4_str(proc);
    XLOPER typeX = xl4_str(typeText);
    XLOPER funcX = xl4_str(funcText);
    XLOPER argX = xl4_str(argText);
    XLOPER macroX = xl4_num(1);
    XLOPER catX = xl4_str("facReaderAddinX");
    XLOPER shortcutX = xl4_str("");
    XLOPER helpTopicX = xl4_str("");
    XLOPER helpX = xl4_str(help);
    XLOPER res;
    XLOPER argHelpX[16];
    LPXLOPER args[27];
    int n = 0;
    args[n++] = moduleName;
    args[n++] = &procX;
    args[n++] = &typeX;
    args[n++] = &funcX;
    args[n++] = &argX;
    args[n++] = &macroX;
    args[n++] = &catX;
    args[n++] = &shortcutX;
    args[n++] = &helpTopicX;
    args[n++] = &helpX;
    for (int i = 0; i < argHelpCount && i < 16; i++) {
        argHelpX[i] = xl4_str(argHelp[i]);
        args[n++] = &argHelpX[i];
    }
    ZeroMemory(&res, sizeof(res));
    int rc = excel(xlfRegister, &res, n, args);
    int ok = rc == 0 && (res.xltype & 0x0FFF) != xltypeErr;
    free_xl4_str(&procX);
    free_xl4_str(&typeX);
    free_xl4_str(&funcX);
    free_xl4_str(&argX);
    free_xl4_str(&catX);
    free_xl4_str(&shortcutX);
    free_xl4_str(&helpTopicX);
    free_xl4_str(&helpX);
    for (int i = 0; i < argHelpCount && i < 16; i++) {
        free_xl4_str(&argHelpX[i]);
    }
    return ok;
}

__declspec(dllexport) int WINAPI xlAutoOpen(void) {
    Excel4vFn excel = excel4v();
    if (!excel) {
        MessageBoxW(NULL, L"facReaderAddinX could not find Excel4v in XLCALL32.DLL. Function metadata registration failed.", L"facReaderAddinX", MB_OK | MB_ICONERROR);
        return 0;
    }

    wchar_t modulePath[MAX_PATH];
    DWORD modulePathLen = GetModuleFileNameW(g_module, modulePath, MAX_PATH);
    if (modulePathLen == 0 || modulePathLen >= MAX_PATH) {
        MessageBoxW(NULL, L"facReaderAddinX could not get the XLL module path. Function metadata registration failed.", L"facReaderAddinX", MB_OK | MB_ICONERROR);
        return 0;
    }

    XLOPER moduleName;
    moduleName = xl4_str_wide(modulePath);
    if (moduleName.xltype == 0) {
        MessageBoxW(NULL, L"facReaderAddinX could not convert the XLL module path. Function metadata registration failed.", L"facReaderAddinX", MB_OK | MB_ICONERROR);
        return 0;
    }

    const char *readHelp[] = {
        "Result vault folder.", "Result ID subfolder.", ".fac file name.",
        "Lookup coordinate 1.", "Lookup coordinate 2.", "Lookup coordinate 3.", "Lookup coordinate 4.",
        "Lookup coordinate 5.", "Lookup coordinate 6.", "Lookup coordinate 7.", "Lookup coordinate 8.",
        "Lookup coordinate 9.", "Lookup coordinate 10.", "Lookup coordinate 11.", "Lookup coordinate 12."
    };
    const char *projHelp[] = {
        "Result vault folder.", "Result ID subfolder.", "Product / .fac file name.",
        "SP_CODE.", "Reserved result type.", "VAR_NAME.", "Time period column.", "SIM_ID, usually 0."
    };

    int readOk = register_one(excel, &moduleName, "ERead_Result", "PPPPPPPPPPPPPPPP!", "ERead_Result",
        "resVault,resID,filename,key1,key2,key3,key4,key5,key6,key7,key8,key9,key10,key11,key12",
        "Read one value from a .fac file using table-order coordinates.", readHelp, 15);
    int projOk = register_one(excel, &moduleName, "EProj_Result", "PPPPPPPPP!", "EProj_Result",
        "resVault,resID,product,spCode,resultType,variable,timePeriod,simID",
        "Read one projected value from a .fac file.", projHelp, 8);

    free_xl4_str(&moduleName);
    if (!readOk || !projOk) {
        MessageBoxW(NULL, L"facReaderAddinX loaded, but Excel rejected function metadata registration.", L"facReaderAddinX", MB_OK | MB_ICONERROR);
        return 0;
    }
    return 1;
}

__declspec(dllexport) LPXLOPER WINAPI ERead_Result(LPXLOPER resVault, LPXLOPER resID, LPXLOPER filename,
    LPXLOPER key1, LPXLOPER key2, LPXLOPER key3, LPXLOPER key4, LPXLOPER key5, LPXLOPER key6,
    LPXLOPER key7, LPXLOPER key8, LPXLOPER key9, LPXLOPER key10, LPXLOPER key11, LPXLOPER key12) {
    if (!load_app_reader()) return ret_text("#ERROR: facReaderAddin.dll not found or incompatible.");

    char *vault = xl_to_utf8(resVault);
    char *rid = xl_to_utf8(resID);
    char *fn = xl_to_utf8(filename);
    if (!vault || !rid || !fn) {
        free(vault); free(rid); free(fn);
        return ret_text("#ERROR: invalid ERead_Result arguments (resVault, resID, filename are required).");
    }
    char *path = build_fac_path(vault, rid, fn);
    if (!path) {
        free(vault); free(rid); free(fn);
        return ret_text("#ERROR: failed to build result file path.");
    }
    LPXLOPER keyOps[12] = { key1, key2, key3, key4, key5, key6, key7, key8, key9, key10, key11, key12 };
    char *keys[12] = {0};
    int dims = g_facNumDims ? g_facNumDims(path) : 0;
    int last = -1;
    for (int i = 0; i < 12; i++) {
        if (!is_missing(keyOps[i])) {
            last = i;
        }
    }
    int n = last + 1;
    if (dims > n && dims <= 12) n = dims;
    for (int i = 0; i < n; i++) {
        keys[i] = xl_to_utf8(keyOps[i]);
    }

    double value = 0;
    int status = g_readResult(path, n, keys[0], keys[1], keys[2], keys[3], keys[4], keys[5],
        keys[6], keys[7], keys[8], keys[9], keys[10], keys[11], &value);

    for (int i = 0; i < 12; i++) free(keys[i]);

    if (status == 0) {
        LPXLOPER result = ret_num(value);
        free(vault); free(rid); free(fn); free(path);
        return result;
    }
    LPXLOPER result = read_error_result(path, dims, n, status);
    free(vault); free(rid); free(fn); free(path);
    return result;
}

__declspec(dllexport) LPXLOPER WINAPI EProj_Result(LPXLOPER resVault, LPXLOPER resID, LPXLOPER product,
    LPXLOPER spCode, LPXLOPER resultType, LPXLOPER variable, LPXLOPER timePeriod, LPXLOPER simID) {
    if (!load_app_reader()) return ret_text("#ERROR: facReaderAddin.dll not found or incompatible.");

    char *vault = xl_to_utf8(resVault);
    char *rid = xl_to_utf8(resID);
    char *prod = xl_to_utf8(product);
    if (!vault || !rid || !prod) {
        free(vault); free(rid); free(prod);
        return ret_text("#ERROR: invalid EProj_Result arguments (resVault, resID, product are required).");
    }
    char *path = build_fac_path(vault, rid, prod);
    char *sp = xl_to_utf8(spCode);
    char *rt = xl_to_utf8(resultType);
    char *var = xl_to_utf8(variable);
    char *period = period_to_utf8(timePeriod);
    char *sim = is_missing(simID) ? _strdup("0") : xl_to_utf8(simID);
    if (!sp || !rt || !var || !period || !sim) {
        free(vault); free(rid); free(prod); free(path);
        free(sp); free(rt); free(var); free(period); free(sim);
        return ret_text("#ERROR: invalid EProj_Result arguments (spCode, resultType, variable, timePeriod, simID).");
    }
    if (!path) {
        free(vault); free(rid); free(prod);
        free(sp); free(rt); free(var); free(period); free(sim);
        return ret_text("#ERROR: failed to build result file path.");
    }
    double value = 0;
    int status = g_projResult(path, sp, rt, var, period, sim, &value);

    if (status == 0) {
        LPXLOPER result = ret_num(value);
        free(vault); free(rid); free(prod); free(path);
        free(sp); free(rt); free(var); free(period); free(sim);
        return result;
    }
    LPXLOPER result = proj_error_result(path, status, sp, var, period, sim);
    free(vault); free(rid); free(prod); free(path);
    free(sp); free(rt); free(var); free(period); free(sim);
    return result;
}
