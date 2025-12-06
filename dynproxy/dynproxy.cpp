#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS 1
#include <windows.h>
#include <oaidl.h>
#include <objbase.h>
#include <unknwn.h>
#include <string>
#include <unordered_map>
#include <vector>
#include <strsafe.h>

#include "dynproxy.h"

/*
'Author:  David Zimmer <dzzie@yahoo.com> + chatgpt
'Site:    http://sandsprite.com
'License: MIT
*/

//debug messages use Elroys:  http://www.vbforums.com/showthread.php?874127-Persistent-Debug-Print-Window

void msg(char);
void msgf(const char*, ...);

bool Warned = false;
HWND hServer = 0;
int DEBUG_MODE = 0;

static BSTR SysAllocStringFromW(const std::wstring& w) {
    return SysAllocStringLen(w.data(), (UINT)w.size());
}

HWND regFindWindow(void) {

    const char* baseKey = "Software\\VB and VBA Program Settings\\dbgWindow\\settings";
    char tmp[20] = { 0 };
    unsigned long l = sizeof(tmp);
    HWND ret = 0;
    HKEY h;

    //printf("regFindWindow triggered\n");

    RegOpenKeyExA(HKEY_CURRENT_USER, baseKey, 0, KEY_READ, &h);
    RegQueryValueExA(h, "hwnd", 0, 0, (unsigned char*)tmp, &l);
    RegCloseKey(h);

    ret = (HWND)atoi(tmp);
    if (!IsWindow(ret)) ret = 0;
    return ret;
}

int FindVBWindow() {
    const char* vbIDEClassName = "ThunderFormDC";
    const char* vbEXEClassName = "ThunderRT6FormDC";
    const char* vbEXEClassName2 = "ThunderRT6Form";
    const char* vbWindowCaption = "Persistent Debug Print Window";

    hServer = FindWindowA(vbIDEClassName, vbWindowCaption);
    if (hServer == 0) hServer = FindWindowA(vbEXEClassName, vbWindowCaption);
    if (hServer == 0) hServer = FindWindowA(vbEXEClassName2, vbWindowCaption);
    if (hServer == 0) hServer = regFindWindow(); //if ide is running as admin

    if (hServer == 0) {
        if (!Warned) {
            //MessageBox(0,"Could not find msg window","",0);
            //printf("Could not find msg window\n");
            Warned = true;
        }
    }
    else {
        if (!Warned) {
            //first time we are being called we could do stuff here...
            //printf("hServer = %x\n", hServer);
            Warned = true;

        }
    }

	return (int)hServer;

}

int msg(char* Buffer, int force=0) {

	if (!DEBUG_MODE && force==0) return 0;
    if (!IsWindow(hServer)) hServer = 0;
    if (hServer == 0) FindVBWindow();

    COPYDATASTRUCT cpStructData;
    memset(&cpStructData, 0, sizeof(struct tagCOPYDATASTRUCT));

    //_snprintf(msgbuf, 0x1000, "%x,%x,%s", myPID, GetCurrentThreadId(), Buffer);

    cpStructData.dwData = 3;
    cpStructData.cbData = strlen(Buffer);
    cpStructData.lpData = (void*)Buffer;
    int ret = SendMessage(hServer, WM_COPYDATA, 0, (LPARAM)&cpStructData);
    return ret; //log ui can send us a response msg to trigger special reaction in ret

}

void msgf(const char* format, ...)
{
    
	if (!DEBUG_MODE) return;
	DWORD dwErr = GetLastError();

    if (format) {
        char buf[1024];
        va_list args;
        va_start(args, format);
        try {
            _vsnprintf(buf, 1024, format, args);
            msg(buf);
        }
        catch (...) {}
    }

    SetLastError(dwErr);
}

// Fallback implementation (uncomment if you don't have msgf):
/*
#include <stdio.h>
#include <stdarg.h>
#include <windows.h>
extern "C" void msgf(const char* format, ...) {
    char buf[2048];
    va_list ap; va_start(ap, format);
    _vsnprintf_s(buf, _countof(buf), _TRUNCATE, format, ap);
    va_end(ap);
    OutputDebugStringA(buf);
    OutputDebugStringA("\r\n");
}
*/


static void DBGW(const wchar_t* fmt, ...) {
    wchar_t wbuf[1024];

	if (!DEBUG_MODE) return;
    va_list ap; va_start(ap, fmt);
    StringCchVPrintfW(wbuf, _countof(wbuf), fmt, ap);
    va_end(ap);

    // Convert to UTF-8 (or ANSI) for msgf
    char buf[2048];
    int n = WideCharToMultiByte(CP_UTF8, 0, wbuf, -1, buf, (int)sizeof(buf), nullptr, nullptr);
    if (n <= 0) { buf[0] = 0; }
    msgf("%s", buf);
}

static void DBGA(const char* fmt, ...) {
    char buf[2048];

	if (!DEBUG_MODE) return;
    va_list ap; va_start(ap, fmt);
    _vsnprintf_s(buf, _countof(buf), _TRUNCATE, fmt, ap);
    va_end(ap);
    msgf("%s", buf);
}



// -------- VB6 resolver bridge (IDispatch with methods by name) ----------
struct VB6ResolverBridge {
    IDispatch* disp;          // not owned here; we AddRef in ctor, Release in dtor
    DISPID dispidGetID = DISPID_UNKNOWN;
    DISPID dispidInvoke = DISPID_UNKNOWN;

    VB6ResolverBridge(IDispatch* p) : disp(p) {
        if (disp) disp->AddRef();
        msgf("[bridge] VB6ResolverBridge() this=%p disp=%p", this, disp);
        cacheDispIDs();
    }

    ~VB6ResolverBridge() { 
        msgf("[bridge] ~VB6ResolverBridge() this=%p disp=%p", this, disp);
        if (disp) disp->Release(); 
    }

    void cacheDispIDs() {
        if (!disp) return;
        LPOLESTR n1 = const_cast<LPOLESTR>(L"ResolveGetID");
        LPOLESTR n2 = const_cast<LPOLESTR>(L"ResolveInvoke");
        disp->GetIDsOfNames(IID_NULL, &n1, 1, LOCALE_USER_DEFAULT, &dispidGetID);
        disp->GetIDsOfNames(IID_NULL, &n2, 1, LOCALE_USER_DEFAULT, &dispidInvoke);
    }

    // Call VB6 ResolveGetID(name) As Long  (negative/positive; your choice)
// In VB6ResolverBridge::ResolveGetID
    bool ResolveGetID(const std::wstring& name, DISPID& out) {
        if (!disp) return false;
        ensureDispIDs();
        if (dispidGetID == DISPID_UNKNOWN) { msgf("[bridge] ResolveGetID not found"); return false; }

        VARIANTARG arg; VariantInit(&arg);
        arg.vt = VT_BSTR; arg.bstrVal = SysAllocStringLen(name.data(), (UINT)name.size());
        DISPPARAMS dp{ &arg, nullptr, 1, 0 };
        VARIANT res; VariantInit(&res);

        msgf("[bridge] calling ResolveGetID('%S')", name.c_str());
        HRESULT hr = disp->Invoke(dispidGetID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &res, nullptr, nullptr);
        VariantClear(&arg);

        if (SUCCEEDED(hr) && (res.vt == VT_I4 || res.vt == VT_I2)) {
            DISPID id = (DISPID)res.lVal;
            VariantClear(&res);
            if (id == 0) {                  // <-- treat 0 as "no override"
                msgf("[bridge] ResolveGetID returned 0 (no override) for '%S'", name.c_str());
                return false;
            }
            out = id;
            msgf("[bridge] ResolveGetID -> %ld", (long)out);
            return true;
        }
        msgf("[bridge] ResolveGetID hr=0x%08X vt=%d", (unsigned)hr, (int)res.vt);
        VariantClear(&res);
        return false;
    }

    void ensureDispIDs() {
        if (!disp) return;
        if (dispidGetID == DISPID_UNKNOWN) {
            LPOLESTR n = const_cast<LPOLESTR>(L"ResolveGetID");
            disp->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &dispidGetID);
            DBGW(L"[bridge] ResolveGetID dispid=%ld", (long)dispidGetID);
        }
        if (dispidInvoke == DISPID_UNKNOWN) {
            LPOLESTR n = const_cast<LPOLESTR>(L"ResolveInvoke");
            disp->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &dispidInvoke);
            DBGW(L"[bridge] ResolveInvoke dispid=%ld", (long)dispidInvoke);
        }
    }

    HRESULT ResolveInvoke(const std::wstring& name, WORD flags, DISPPARAMS* p, VARIANT* pRes)
    {
        if (!disp) return DISP_E_MEMBERNOTFOUND;
        ensureDispIDs();
        if (dispidInvoke == DISPID_UNKNOWN) {
            msgf("[bridge] ResolveInvoke unknown dispid");
            return DISP_E_MEMBERNOTFOUND;
        }

        LONG n = (LONG)p->cArgs;
        SAFEARRAYBOUND b{ (ULONG)n, 0 };
        SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 1, &b);
        if (!psa) return E_OUTOFMEMORY;

        for (LONG i = 0; i < n; ++i) {
            VARIANT v; VariantInit(&v);
            VariantCopyInd(&v, &p->rgvarg[n - 1 - i]); // COM RtL -> LtR
            SafeArrayPutElement(psa, &i, &v);
            VariantClear(&v);
        }

        VARIANTARG argv[4];
        for (int i = 0; i < 4; ++i) VariantInit(&argv[i]);

        //Pass cArgs as 4th argument - (variant array gets initilized funny for 0 args)
        argv[0].vt = VT_I4;
        argv[0].lVal = (LONG)n;  // cArgs count!

        // Build right-to-left call args for VB6
        argv[1].vt = VT_BYREF | VT_ARRAY | VT_VARIANT;   // args() As Variant()  (ByRef SAFEARRAY)
        argv[1].pparray = &psa;

        argv[2].vt = VT_I4;
        argv[2].lVal = (LONG)flags;

        argv[3].vt = VT_BSTR;
        argv[3].bstrVal = SysAllocStringLen(name.data(), (UINT)name.size());

        msgf("[bridge] calling ResolveInvoke name=%S flags=0x%X argc=%ld", name.c_str(), (unsigned)flags, (long)n);

        DISPPARAMS dp{ argv, nullptr, 4, 0 };
        VARIANT res; VariantInit(&res);

        HRESULT hr = disp->Invoke(dispidInvoke, IID_NULL, LOCALE_USER_DEFAULT,
            DISPATCH_METHOD, &dp, &res, nullptr, nullptr);

        if (FAILED(hr)) {
            msgf("[bridge] ResolveInvoke hr=0x%08X", (unsigned)hr);
        }
        else {
            msgf("[bridge] ResolveInvoke succeeded vt=%d", res.vt);
        }

        VariantClear(&argv[2]);   // free BSTR
        SafeArrayDestroy(psa);    // free SAFEARRAY

        if (SUCCEEDED(hr) && pRes) {
            VariantClear(pRes);
            VariantCopy(pRes, &res);
        }

        VariantClear(&res);
        return hr;
    }



};

// -------- Raw proxy implementing IDispatch (no ATL/MFC) ----------
class ProxyDispatch final : public IDispatch {
public:

    ProxyDispatch(IDispatch* inner, IDispatch* resolver, bool resolverWins)
        : m_ref(1), m_inner(inner),
        m_resolver(resolver ? new VB6ResolverBridge(resolver) : nullptr),
        m_resolverWins(resolverWins)
    {
        if (m_inner) m_inner->AddRef();
    }

    ~ProxyDispatch() {
        msgf("[proxy] ~ProxyDispatch() this=%p m_ref=%ld", this, (long)m_ref);
        if (m_inner) m_inner->Release();
        delete m_resolver;
    }

    void SetResolverWins(bool on) {
        m_resolverWins = on;
        msgf("[proxy] resolverWins = %d", (int)m_resolverWins);
    }

    void ClearNameCache() {
        m_nameToDisp.clear();
        m_dispToName.clear();
        msgf("[proxy] name cache cleared");
    }

    // Force/clear per-name override (case-insensitive optional)
    void OverrideName(const std::wstring& nm, DISPID id) {
        std::wstring key = nm; // tolower if you want case-insensitive
        auto it = m_nameToDisp.find(key);
        if (it != m_nameToDisp.end()) {
            // remove old dynamic mapping if any
            m_dispToName.erase(it->second);
            m_nameToDisp.erase(it);
        }
        if (id == 0) { msgf("[proxy] override cleared for '%S'", nm.c_str()); return; }
        m_nameToDisp[key] = id;          // install name->id
        m_dispToName[id] = key;         // mark as dynamic so Invoke routes to resolver
        msgf("[proxy] override set '%S' -> %ld", nm.c_str(), (long)id);
    }


    // IUnknown
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IDispatch) {
            *ppv = static_cast<IDispatch*>(this);
            AddRef(); return S_OK;
        }
        *ppv = nullptr; return E_NOINTERFACE;
    }
    
    ULONG STDMETHODCALLTYPE AddRef() override {
        LONG r = InterlockedIncrement(&m_ref);
        msgf("[proxy] AddRef() this=%p ref=%ld", this, (long)r);
        return (ULONG)r; 
    }

    ULONG STDMETHODCALLTYPE Release() override {
        LONG r = InterlockedDecrement(&m_ref);
        msgf("[proxy] Release() this=%p ref=%ld", this, (long)r);
        if (!r) {
            msgf("[proxy] Deleting this=%p", this); 
            delete this; 
            return 0; 
        }
        return (ULONG)r;
    }

    // IDispatch
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override { if (!pctinfo) return E_POINTER; *pctinfo = 0; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }

    // In ProxyDispatch::GetIDsOfNames
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* names, UINT c, LCID lcid, DISPID* ids) override {
        if (!names || !ids) return E_POINTER;

        for (UINT i = 0; i < c; ++i) {
            std::wstring nm(names[i]);

            // cache hit?
            auto it = m_nameToDisp.find(nm);
            if (it != m_nameToDisp.end()) {
                ids[i] = it->second;
                msgf("[proxy] cached '%S' -> %ld", nm.c_str(), (long)ids[i]);
                continue;
            }

            DISPID id = 0;
            bool have = false;
            bool dynamic_id = false; // <-- NEW: track if it's ours (resolver/invented)

            // resolver-first if enabled
            if (m_resolverWins && m_resolver) {
                if (m_resolver->ResolveGetID(nm, id)) {
                    // treat id==0 as "no override"
                    if (id != 0) {
                        have = true;
                        dynamic_id = true;
                        msgf("[proxy] resolver-wins GetID '%S' -> %ld", nm.c_str(), (long)id);
                    }
                }
            }

            // inner
            if (!have && m_inner) {
                HRESULT hr = m_inner->GetIDsOfNames(IID_NULL, &names[i], 1, lcid, &id);
                if (SUCCEEDED(hr)) {
                    have = true;
                    dynamic_id = false; // inner-owned
                    msgf("[proxy] inner GetIDsOfNames '%S' -> %ld", nm.c_str(), (long)id);
                }
                else {
                    msgf("[proxy] inner unknown '%S' hr=0x%08X", nm.c_str(), (unsigned)hr);
                }
            }

            // resolver-last (only if not resolverWins)
            if (!have && m_resolver) {
                if (m_resolver->ResolveGetID(nm, id)) {
                    if (id != 0) {
                        have = true;
                        dynamic_id = true;
                        msgf("[proxy] resolver GetID '%S' -> %ld", nm.c_str(), (long)id);
                    }
                }
            }

            // invent
            if (!have) {
                id = m_nextFake--;
                dynamic_id = true;
                msgf("[proxy] invented dispid '%S' -> %ld", nm.c_str(), (long)id);
            }

            // Always remember name->id
            m_nameToDisp.emplace(nm, id);
            // Only remember id->name if it's OUR dynamic id
            if (dynamic_id) {
                m_dispToName.emplace(id, nm);
            }

            ids[i] = id;
        }
        return S_OK;
    }



    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispId, REFIID, LCID, WORD flags, DISPPARAMS* p, VARIANT* r, EXCEPINFO* ex, UINT* argerr) override
    {
        if (dispId == DISPID_VALUE) {
            msgf("[proxy] !!!!! dispId IS DISPID_VALUE !!!!!");
            msgf("[proxy] m_resolver = %p", m_resolver);
            msgf("[proxy] p->cArgs = %d", (int)(p ? p->cArgs : -999));

            if (!m_resolver) {
                msgf("[proxy] !!!!! RESOLVER IS NULL !!!!!");
            }
            if (p->cArgs == 0) {
                msgf("[proxy] !!!!! NO ARGS !!!!!");
            }
        }
        
        //NEW: Handle DISPID_VALUE (array indexing) FIRST
        if ((long)dispId == DISPID_VALUE && m_resolver ) {
            msgf("[proxy] DISPID_VALUE with %d args, routing to resolver", p->cArgs);
            return m_resolver->ResolveInvoke(L"<DISPID_0>", flags, p, r);
        }

        auto it = m_dispToName.find(dispId);
        if (it != m_dispToName.end()) {
            DBGW(L"[proxy] dynamic Invoke dispid=%ld name=%s flags=0x%X", (long)dispId, it->second.c_str(), (unsigned)flags);
            if (!m_resolver) return DISP_E_MEMBERNOTFOUND;
            return m_resolver->ResolveInvoke(it->second, flags, p, r);
        }
        if (m_inner) {
            DBGW(L"[proxy] forwarding to inner dispid=%ld flags=0x%X", (long)dispId, (unsigned)flags);
            return m_inner->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, flags, p, r, ex, argerr);
        }
        DBGW(L"[proxy] unknown dispid=%ld (no inner/resolver)", (long)dispId);
        return DISP_E_MEMBERNOTFOUND;
    }


private:
    LONG m_ref;
    IDispatch* m_inner;                 // inner (optional)
    VB6ResolverBridge* m_resolver;      // VB6 callback object bridge (optional)
    std::unordered_map<std::wstring, DISPID> m_nameToDisp;
    std::unordered_map<DISPID, std::wstring> m_dispToName;
    DISPID m_nextFake = -1000;
    bool m_resolverWins = false;
};

// helper: CoCreate inner from ProgID
static HRESULT CoCreateFromProgID(BSTR progId, IDispatch** ppDisp) {
    if (!ppDisp) return E_POINTER;
    *ppDisp = nullptr;
    CLSID clsid; HRESULT hr = CLSIDFromProgID(progId, &clsid);
    if (FAILED(hr)) return hr;
    IDispatch* d = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&d);
    if (SUCCEEDED(hr)) *ppDisp = d;
    return hr;
}

// ---- stdcall exports ----
extern "C" __declspec(dllexport) ULONG_PTR __stdcall CreateProxyForProgIDRaw(BSTR progId, IDispatch * resolverDisp)
{
    // Ensure COM initialized in the caller STA
    // (VB6 main thread is already initialized; if you call from others, you must CoInitializeEx there.)
    IDispatch* inner = nullptr;
    if (progId && progId[0]) {
        CoCreateFromProgID(progId, &inner); // ignore failure -> still create proxy with null inner
    }

    ProxyDispatch* p = new(std::nothrow) ProxyDispatch(inner, resolverDisp, false);
    if (inner) inner->Release();
    if (!p) return 0;

    IDispatch* out = nullptr;
    if (FAILED(p->QueryInterface(IID_IDispatch, (void**)&out))) {
        p->Release();
        return 0;
    }
    // balance initial ref
    p->Release();
    return (ULONG_PTR)out; // refcount 1 owned by caller
}

// ... keep your ProxyDispatch class & VB6ResolverBridge from earlier ...

// Helper: turn a VB6 ObjPtr (IUnknown*) into a real IDispatch* with AddRef via QI.
static IDispatch* QI_IDispatch_FromRawPtr(ULONG_PTR raw)
{
    if (!raw) return nullptr;
    IUnknown* unk = reinterpret_cast<IUnknown*>(raw);
    IDispatch* disp = nullptr;
    // QI adds a ref if successful so lifetime is correct
    if (unk) unk->QueryInterface(IID_IDispatch, (void**)&disp);
    return disp; // may be nullptr if object doesn't support IDispatch (VB6 always does)
}

extern "C" __declspec(dllexport) ULONG_PTR __stdcall CreateProxyForObjectRaw(
    ULONG_PTR innerPtr, ULONG_PTR resolverPtr)
{
    DBGW(L"[api] CreateProxyForObjectRaw inner=0x%p resolver=0x%p", (void*)innerPtr, (void*)resolverPtr);
    IDispatch* inner = QI_IDispatch_FromRawPtr(innerPtr);
    IDispatch* resolver = QI_IDispatch_FromRawPtr(resolverPtr);
    if (innerPtr && !inner) DBGW(L"[api] inner QI for IDispatch FAILED");
    if (resolverPtr && !resolver) DBGW(L"[api] resolver QI for IDispatch FAILED");

    auto p = new(std::nothrow) ProxyDispatch(inner, resolver, false);
    msgf("[api] Created ProxyDispatch at %p", p);
    if (inner) inner->Release();
    if (resolver) resolver->Release();
    if (!p) return 0;

    IDispatch* out = nullptr;
    HRESULT hr = p->QueryInterface(IID_IDispatch, (void**)&out);
    DBGW(L"[api] QI(IID_IDispatch) hr=0x%08X out=0x%p", (unsigned)hr, out);
    p->Release();

    msgf("[api] Returning proxy IDispatch=%p", out );
    return (ULONG_PTR)out;
}



static ProxyDispatch* fromRawProxy(ULONG_PTR p)
{
    return reinterpret_cast<ProxyDispatch*>(p);
}

extern "C" __declspec(dllexport) void __stdcall ReleaseDispatchRaw(ULONG_PTR pDisp)
{
    if (!pDisp) return;
    IDispatch* d = reinterpret_cast<IDispatch*>(pDisp);
    d->Release();
}

extern "C" __declspec(dllexport) void __stdcall SetProxyResolverWins(ULONG_PTR proxyDispPtr, int enable)
{
    if (!proxyDispPtr) return;
    auto* px = fromRawProxy(proxyDispPtr);
    px->SetResolverWins(enable != 0);
}

extern "C" __declspec(dllexport) void __stdcall ClearProxyNameCache(ULONG_PTR proxyDispPtr)
{
    if (!proxyDispPtr) return;
    auto* px = fromRawProxy(proxyDispPtr);
    px->ClearNameCache();
}

// Create with resolverWins at construction time
extern "C" __declspec(dllexport) ULONG_PTR __stdcall CreateProxyForObjectRawEx(
    ULONG_PTR innerPtr, ULONG_PTR resolverPtr, int resolverWins)
{
    IDispatch* inner = QI_IDispatch_FromRawPtr(innerPtr);
    IDispatch* resolver = QI_IDispatch_FromRawPtr(resolverPtr);

    auto p = new(std::nothrow) ProxyDispatch(inner, resolver, resolverWins != 0);
    if (inner) inner->Release();
    if (resolver) resolver->Release();
    if (!p) return 0;

    IDispatch* out = nullptr;
    if (FAILED(p->QueryInterface(IID_IDispatch, (void**)&out))) { p->Release(); return 0; }
    p->Release();
    return (ULONG_PTR)out;
}

// Per-name override: set nonzero dispid to route to resolver; 0 clears override
extern "C" __declspec(dllexport) void __stdcall SetProxyOverride(
    ULONG_PTR proxyPtr, BSTR name, long dispid)
{
    if (!proxyPtr || !name) return;
    fromRawProxy(proxyPtr)->OverrideName(std::wstring(name), (DISPID)dispid);
}



// optional DllMain
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved) {
    if (fdwReason == 1) msg((char*)"<cls>",1);
    return true;
}

//no reason to duplicate the C code in VB when we have the dll loaded...
extern "C" __declspec(dllexport) void __stdcall SendDbgMsg(char* str) {msg(str,1);}

//method calls work for objects or regular return types (vb set limitation)
/*
extern "C" __declspec(dllexport)
HRESULT __stdcall CallByNameEx(
	IUnknown* pUnk,
	const char* memberName,     // VB6 ByVal String -> char*
	WORD invokeFlags,
	VARIANT* pVarArgs,
	VARIANT* pResult,
	BOOL* pIsObject
)
{
	msgf("[CallByNameEx] Entry: pUnk=%p name='%s' flags=%d", pUnk, memberName ? memberName : "(null)", invokeFlags);

	if (!pUnk || !memberName || !pResult || !pIsObject) {
		msgf("[CallByNameEx] E_POINTER: bad params");
		return E_POINTER;
	}

	// QueryInterface for IDispatch
	IDispatch* obj = nullptr;
	HRESULT hr = pUnk->QueryInterface(IID_IDispatch, (void**)&obj);
	msgf("[CallByNameEx] QueryInterface hr=0x%08X obj=%p", hr, obj);

	if (FAILED(hr) || !obj) {
		return E_NOINTERFACE;
	}

	VariantInit(pResult);
	*pIsObject = FALSE;

	// Convert ANSI to Unicode BSTR
	int wlen = MultiByteToWideChar(CP_ACP, 0, memberName, -1, nullptr, 0);
	BSTR wideName = SysAllocStringLen(nullptr, wlen - 1);
	MultiByteToWideChar(CP_ACP, 0, memberName, -1, wideName, wlen);

	// Extract SAFEARRAY from VARIANT
	SAFEARRAY* args = nullptr;
	if (pVarArgs && (pVarArgs->vt & VT_ARRAY)) {
		args = pVarArgs->parray;
		msgf("[CallByNameEx] Got array from VARIANT, vt=0x%X parray=%p", pVarArgs->vt, args);
	}

	// Get DISPID
	DISPID dispid;
	hr = obj->GetIDsOfNames(IID_NULL, &wideName, 1, LOCALE_USER_DEFAULT, &dispid);
	msgf("[CallByNameEx] GetIDsOfNames hr=0x%08X dispid=%ld", hr, dispid);

	if (FAILED(hr)) {
		SysFreeString(wideName);
		obj->Release();
		return hr;
	}

	// Build DISPPARAMS
	DISPPARAMS dp = { 0 };
	if (args) {
		LONG lb = 0, ub = 0;
		SafeArrayGetLBound(args, 1, &lb);
		SafeArrayGetUBound(args, 1, &ub);
		LONG count = ub - lb + 1;
		msgf("[CallByNameEx] Array bounds: lb=%ld ub=%ld count=%ld", lb, ub, count);

		if (count > 0) {
			dp.rgvarg = (VARIANTARG*)CoTaskMemAlloc(count * sizeof(VARIANTARG));
			dp.cArgs = count;

			// Go back to SafeArrayGetElement - it's safer
			for (LONG i = 0; i < count; i++) {
				VARIANT v;
				VariantInit(&v);
				LONG idx = lb + i;

				HRESULT saHr = SafeArrayGetElement(args, &idx, &v);
				msgf("[CallByNameEx] SafeArrayGetElement[%ld] hr=0x%08X vt=%d", i, saHr, v.vt);

				if (SUCCEEDED(saHr)) {
					VariantInit(&dp.rgvarg[count - 1 - i]);
					VariantCopyInd(&dp.rgvarg[count - 1 - i], &v);
					msgf("[CallByNameEx] Copied arg[%ld]", i);
				}
				else {
					msgf("[CallByNameEx] ERROR: Failed to get element %ld", i);
				}
				VariantClear(&v);
			}
		}
	}

	// INVOKE
	msgf("[CallByNameEx] Calling Invoke dispid=%ld flags=%d cArgs=%d...", dispid, invokeFlags, dp.cArgs);
	hr = obj->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
		invokeFlags, &dp, pResult, nullptr, nullptr);
	msgf("[CallByNameEx] Invoke hr=0x%08X result.vt=%d", hr, pResult->vt);

	// Check if result is object
	if (SUCCEEDED(hr)) {
		*pIsObject = (pResult->vt == VT_DISPATCH ||
			pResult->vt == VT_UNKNOWN ||
			pResult->vt == (VT_BYREF | VT_DISPATCH) ||
			pResult->vt == (VT_BYREF | VT_UNKNOWN));
		msgf("[CallByNameEx] isObject=%d", *pIsObject);
	}

	// Cleanup
	if (dp.rgvarg) {
		for (UINT i = 0; i < dp.cArgs; i++) {
			VariantClear(&dp.rgvarg[i]);
		}
		CoTaskMemFree(dp.rgvarg);
	}

	SysFreeString(wideName);

	msgf("[CallByNameEx] Exit hr=0x%08X", hr);
	return hr;
}
*/

extern "C" __declspec(dllexport)
HRESULT __stdcall CallByNameEx(
	IUnknown* pUnk,
	const char* memberName,
	WORD invokeFlags,
	void* pArgsOrVariant,
	VARIANT* pResult,
	VARIANT_BOOL* pIsObject
)
{
	msgf("[CallByNameEx] Entry: pUnk=%p name='%s' flags=%d pResult=%p", pUnk, memberName ? memberName : "(null)", invokeFlags, pResult);

	if (!pUnk || !memberName || !pResult || !pIsObject) {
		msgf("[CallByNameEx] E_POINTER: bad params");
		return E_POINTER;
	}

	IDispatch* obj = nullptr;
	HRESULT hr = pUnk->QueryInterface(IID_IDispatch, (void**)&obj);
	msgf("[CallByNameEx] QueryInterface hr=0x%08X obj=%p", hr, obj);

	if (FAILED(hr) || !obj) return E_NOINTERFACE;

	// DON'T clear pResult - let Invoke write to it directly
	msgf("[CallByNameEx] Before operations: pResult=%p vt=%d", pResult, pResult->vt);
	*pIsObject = FALSE;

	int wlen = MultiByteToWideChar(CP_ACP, 0, memberName, -1, nullptr, 0);
	BSTR wideName = SysAllocStringLen(nullptr, wlen - 1);
	MultiByteToWideChar(CP_ACP, 0, memberName, -1, wideName, wlen);

	DISPID dispid;
	hr = obj->GetIDsOfNames(IID_NULL, &wideName, 1, LOCALE_USER_DEFAULT, &dispid);
	msgf("[CallByNameEx] GetIDsOfNames hr=0x%08X dispid=%ld", hr, dispid);

	if (FAILED(hr)) {
		SysFreeString(wideName);
		obj->Release();
		return hr;
	}

	// Detect array type
	SAFEARRAY* args = nullptr;
	SAFEARRAY** ppArgs = (SAFEARRAY**)pArgsOrVariant;

	if (ppArgs && *ppArgs) {
		SAFEARRAY* candidate = *ppArgs;
		msgf("[CallByNameEx] Trying as SAFEARRAY**: %p cDims=%d", candidate, candidate->cDims);
		if (candidate->cDims > 0) {
			args = candidate;
			msgf("[CallByNameEx] SUCCESS: Valid SAFEARRAY**");
		}
	}

	if (!args) {
		VARIANT* pVar = (VARIANT*)pArgsOrVariant;
		msgf("[CallByNameEx] Trying as VARIANT*: vt=0x%X", pVar->vt);
		if ((pVar->vt & VT_ARRAY) == VT_ARRAY) {
			args = pVar->parray;
			msgf("[CallByNameEx] SUCCESS: Extracted from VARIANT");
		}
	}

	if (!args) {
		msgf("[CallByNameEx] WARNING: No valid array found, proceeding with 0 args");
	}

	// Build DISPPARAMS
	DISPPARAMS dp = { 0 };
	VARIANTARG* pAllocatedArgs = nullptr;
	DISPID dispidNamed = DISPID_PROPERTYPUT;  // <--- ADD THIS

	if (args) {
		LONG lb = 0, ub = 0;
		SafeArrayGetLBound(args, 1, &lb);
		SafeArrayGetUBound(args, 1, &ub);
		LONG count = ub - lb + 1;
		msgf("[CallByNameEx] Array bounds: lb=%ld ub=%ld count=%ld", lb, ub, count);

		if (count > 0) {
			pAllocatedArgs = (VARIANTARG*)CoTaskMemAlloc(count * sizeof(VARIANTARG));
			if (!pAllocatedArgs) {
				msgf("[CallByNameEx] E_OUTOFMEMORY");
				SysFreeString(wideName);
				obj->Release();
				return E_OUTOFMEMORY;
			}

			dp.rgvarg = pAllocatedArgs;
			dp.cArgs = count;

			//For property put, mark the value argument
			if (invokeFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
				dp.rgdispidNamedArgs = &dispidNamed;
				dp.cNamedArgs = 1;
				msgf("[CallByNameEx] Property put/putref - setting named arg");
			}

			for (LONG i = 0; i < count; i++) {
				VariantInit(&dp.rgvarg[i]);
			}

			for (LONG i = 0; i < count; i++) {
				VARIANT v;
				VariantInit(&v);
				LONG idx = lb + i;

				HRESULT saHr = SafeArrayGetElement(args, &idx, &v);
				msgf("[CallByNameEx] arg[%ld] hr=0x%08X vt=%d", i, saHr, v.vt);

				if (SUCCEEDED(saHr)) {
					hr = VariantCopyInd(&dp.rgvarg[count - 1 - i], &v);
					if (FAILED(hr)) {
						msgf("[CallByNameEx] VariantCopyInd failed hr=0x%08X", hr);
					}
				}
				VariantClear(&v);
			}
		}
	}

	msgf("[CallByNameEx] Before Invoke: pResult=%p vt=%d", pResult, pResult->vt);
	msgf("[CallByNameEx] Calling Invoke dispid=%ld flags=%d cArgs=%d...", dispid, invokeFlags, dp.cArgs);

	hr = obj->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, invokeFlags, &dp, pResult, nullptr, nullptr);

	msgf("[CallByNameEx] After Invoke: hr=0x%08X pResult=%p vt=%d", hr, pResult, pResult ? pResult->vt : -1);

	if (SUCCEEDED(hr) && pResult) {
		*pIsObject = (pResult->vt == VT_DISPATCH || pResult->vt == VT_UNKNOWN);
		msgf("[CallByNameEx] isObject=%d pResult->vt=%d", *pIsObject, pResult->vt);

		if (pResult->vt == VT_DISPATCH) {
			msgf("[CallByNameEx] result.pdispVal=%p", pResult->pdispVal);
		}
	}

	msgf("[CallByNameEx] Before cleanup: pResult=%p vt=%d", pResult, pResult ? pResult->vt : -1);

	// Cleanup
	if (pAllocatedArgs) {
		for (UINT i = 0; i < dp.cArgs; i++) {
			VariantClear(&pAllocatedArgs[i]);
		}
		CoTaskMemFree(pAllocatedArgs);
	}

	SysFreeString(wideName);
	obj->Release();

	msgf("[CallByNameEx] After cleanup: pResult=%p vt=%d", pResult, pResult ? pResult->vt : -1);
	msgf("[CallByNameEx] Exit hr=0x%08X", hr);

	return hr;
}

extern "C" __declspec(dllexport)
void __stdcall IPCDebugMode(int enabled) { DEBUG_MODE = enabled; }

extern "C" __declspec(dllexport) HRESULT __stdcall StartDbgWnd(int dbgEnabled)
{

	DEBUG_MODE = dbgEnabled;
	if (FindVBWindow() != 0) return 0;

	STARTUPINFOA si = { sizeof(si) };
	PROCESS_INFORMATION pi = { 0 };
	char dllPath[MAX_PATH];
	char exePath[MAX_PATH];

	// Get the DLL's full path
	HMODULE hModule = NULL;
	if (!GetModuleHandleExA(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
		GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
		(LPCSTR)&StartDbgWnd, &hModule)) {
		msgf("[StartDbgWnd] Failed to get module handle: %d", GetLastError());
		return E_FAIL;
	}

	if (!GetModuleFileNameA(hModule, dllPath, MAX_PATH)) {
		msgf("[StartDbgWnd] Failed to get DLL path: %d", GetLastError());
		return E_FAIL;
	}

	// Strip filename to get directory
	char* lastSlash = strrchr(dllPath, '\\');
	if (!lastSlash) {
		msgf("[StartDbgWnd] Invalid DLL path: %s", dllPath);
		return E_FAIL;
	}

	*(lastSlash + 1) = '\0';  // Keep the trailing backslash

	// Build path to dbgwindow.exe
	strcpy_s(exePath, MAX_PATH, dllPath);
	strcat_s(exePath, MAX_PATH, "dbgwindow.exe");

	msgf("[StartDbgWnd] Looking for: %s", exePath);

	// Check if file exists
	if (GetFileAttributesA(exePath) == INVALID_FILE_ATTRIBUTES) {
		msgf("[StartDbgWnd] File not found: %s", exePath);
		return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
	}

	if (!CreateProcessA(exePath, NULL, NULL, NULL, FALSE, 0, NULL, dllPath, &si, &pi))
	{
		msgf("[StartDbgWnd] CreateProcess failed: %d", GetLastError());
		return HRESULT_FROM_WIN32(GetLastError());
	}

	msgf("[StartDbgWnd] Launched PID=%d, waiting for init...", pi.dwProcessId);

	// Wait for the app to be ready (up to 5 seconds)
	DWORD result = WaitForInputIdle(pi.hProcess, 5000);

	if (result == 0) {
		msgf("[StartDbgWnd] Window ready");
	}
	else if (result == WAIT_TIMEOUT) {
		msgf("[StartDbgWnd] WARNING: Timeout waiting for window");
	}
	else {
		msgf("[StartDbgWnd] WARNING: WaitForInputIdle failed");
	}

	CloseHandle(pi.hThread);
	CloseHandle(pi.hProcess);

	return S_OK;
}