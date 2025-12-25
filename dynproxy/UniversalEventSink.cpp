// UniversalEventSink.cpp — single-file COM event sink for forwarding to callbacks
#define UNICODE
#define _UNICODE
#include <windows.h>
#include <oleauto.h>
#include <oaidl.h>
#include <ocidl.h>
#include <olectl.h>
#include <objbase.h>
#include <stdarg.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

// ========================== Debug Output ==========================
extern int msg(const char* Buffer, int force = 0);
extern void msgf(const char* format, ...);

// Forward declaration - ComTypeName from typename.cpp
extern "C" VARIANT __stdcall ComTypeName(IUnknown* pUnk);

/*
 * NOTE: VB6 Intrinsic Control Events (CommandButton, TextBox, etc.)
 *
 * VB6 intrinsic controls use vtable-based event interfaces (derives from IUnknown),
 * NOT IDispatch-based. When they fire events, they call directly into vtable slots:
 *     pSink->vtable[3]()  // Click
 *     pSink->vtable[4]()  // MouseDown, etc.
 *
 * Our universal sink fakes IDispatch interfaces, which works for COM/ActiveX controls.
 * But intrinsic controls expect a real vtable layout - calling our fake sink crashes.
 *
 * To support intrinsics would require dynamically generating vtables with thunks
 * that capture the slot index and forward to a common handler. Doable but gnarly.
 *
 * For now, intrinsic controls require native VB6 WithEvents. Everything else works:
 *   - ActiveX controls (ListView, TreeView, etc. via .object)
 *   - COM automation servers (Excel, Word, etc.)
 *   - VB6 classes with Event declarations
 
	 Address    Stack      Procedure                             Called from                   Frame
	0019F2D8   660123D3   ? MSVBVM60.CallProcWithArgs           MSVBVM60.660123CE             0019F2C4
	0019F2F0   66011B2C   ? MSVBVM60.InvokeVtblEvent            MSVBVM60.66011B27
	0019F324   66005F17   ? MSVBVM60.InvokeEvent                MSVBVM60.66005F12
	0019F3F8   66005D89   MSVBVM60.EvtErrFireWorker             MSVBVM60.66005D84             0019F3F4
	0019F41C   66055FAE   MSVBVM60.EvtErrFire                   MSVBVM60.66055FA9             0019F418
	0019F434   6602E991   MSVBVM60._DoClick                     MSVBVM60.6602E98C             0019F448
	0019F44C   66004738   Includes MSVBVM60.6602E991            MSVBVM60.66004735             0019F448
	0019F474   6600400E   MSVBVM60.CommonGizWndProc             MSVBVM60.66004009             0019F470
	0019F4D0   660570AE   MSVBVM60.StdCtlWndProc                MSVBVM60.660570A9             0019F4CC
	0019F4F4   66003ABF   MSVBVM60._DefWmCommand                MSVBVM60.66003ABA             0019F4F0
	0019F560   66004F29   MSVBVM60.VBDefControlProc             MSVBVM60.66004F24             0019F55C
	0019F6E0   66004738   Includes MSVBVM60.66004F29            MSVBVM60.66004735             0019F6DC
	0019F708   6600400E   MSVBVM60.CommonGizWndProc             MSVBVM60.66004009             0019F704

 */

// ========================== VB6 Intrinsic Control Event IIDs ==========================
// These controls don't expose IProvideClassInfo or EnumConnectionPoints properly,
// but DO support connection points if you know the IID. Extracted from VB6.OLB.

static const struct { const wchar_t* typeName; GUID eventIID; } g_VB6Intrinsics[] = {
	{ L"TextBox",       { 0x33AD4EE2, 0x6699, 0x11CF, { 0xB7, 0x0C, 0x00, 0xAA, 0x00, 0x60, 0xD3, 0x93 } } },
	{ L"ListBox",       { 0x33AD4F12, 0x6699, 0x11CF, { 0xB7, 0x0C, 0x00, 0xAA, 0x00, 0x60, 0xD3, 0x93 } } },
	{ L"Timer",         { 0x33AD4F2A, 0x6699, 0x11CF, { 0xB7, 0x0C, 0x00, 0xAA, 0x00, 0x60, 0xD3, 0x93 } } },
	{ L"Form",          { 0x33AD4F3A, 0x6699, 0x11CF, { 0xB7, 0x0C, 0x00, 0xAA, 0x00, 0x60, 0xD3, 0x93 } } },
	{ L"CommandButton", { 0x33AD4EF2, 0x6699, 0x11CF, { 0xB7, 0x0C, 0x00, 0xAA, 0x00, 0x60, 0xD3, 0x93 } } },
	{ nullptr, { 0 } }  // Sentinel
};

// ========================== Callback Interface ==========================
// Your JS engine would implement this to receive events.
// For now, a simple IDispatch-based callback or function pointer.

// Callback signature: includes source name for multi-object routing
// void __stdcall EventCallback(BSTR sourceName, BSTR eventName, DISPID dispid, 
//                               DISPPARAMS* pParams, IUnknown* pSourceObj, void* userData)
typedef void(__stdcall* PFN_EVENT_CALLBACK)(BSTR sourceName, BSTR eventName, DISPID dispid,
	DISPPARAMS* pParams, IUnknown* pSourceObj, void* userData);

// ========================== Universal Event Sink ==========================
class UniversalEventSink : public IDispatch {
	LONG m_refCount = 1;
	IID m_eventIID = {};              // IID we pretend to support
	ITypeInfo* m_pEventTypeInfo = nullptr;  // For dispid → name lookup
	PFN_EVENT_CALLBACK m_pfnCallback = nullptr;
	void* m_userData = nullptr;
	IDispatch* m_pDispatchCallback = nullptr;  // Alternative: forward via IDispatch::Invoke

	// Source identification
	BSTR m_bstrSourceName = nullptr;  // User-provided name, e.g., "myExcelApp", "wdDoc1"
	IUnknown* m_pSourceObject = nullptr;  // Optional: ref to source for passing through
	DISPID m_dispidOnEvent = DISPID_UNKNOWN;  // Cached dispid for "OnEvent" method

public:
	UniversalEventSink(const IID& eventIID, ITypeInfo* pEventTI, PFN_EVENT_CALLBACK pfn, void* userData,
		BSTR sourceName = nullptr, IUnknown* pSourceObj = nullptr)
		: m_eventIID(eventIID), m_pEventTypeInfo(pEventTI), m_pfnCallback(pfn), m_userData(userData)
	{
		if (m_pEventTypeInfo) m_pEventTypeInfo->AddRef();
		if (sourceName) m_bstrSourceName = SysAllocString(sourceName);
		if (pSourceObj) { m_pSourceObject = pSourceObj; m_pSourceObject->AddRef(); }
	}

	// Alternative ctor: forward to IDispatch (e.g., your JS function wrapper)
	UniversalEventSink(const IID& eventIID, ITypeInfo* pEventTI, IDispatch* pCallback,
		BSTR sourceName = nullptr, IUnknown* pSourceObj = nullptr)
		: m_eventIID(eventIID), m_pEventTypeInfo(pEventTI), m_pDispatchCallback(pCallback)
	{
		if (m_pEventTypeInfo) m_pEventTypeInfo->AddRef();
		if (m_pDispatchCallback) m_pDispatchCallback->AddRef();
		if (sourceName) m_bstrSourceName = SysAllocString(sourceName);
		if (pSourceObj) { m_pSourceObject = pSourceObj; m_pSourceObject->AddRef(); }
	}

	~UniversalEventSink() {
		if (m_pEventTypeInfo) m_pEventTypeInfo->Release();
		if (m_pDispatchCallback) m_pDispatchCallback->Release();
		if (m_bstrSourceName) SysFreeString(m_bstrSourceName);
		if (m_pSourceObject) m_pSourceObject->Release();
	}

	// Accessors for external use
	BSTR GetSourceName() const { return m_bstrSourceName; }
	IUnknown* GetSourceObject() const { return m_pSourceObject; }

	// ---------------------- IUnknown ----------------------
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
		if (!ppv) return E_POINTER;
		*ppv = nullptr;

		if (riid == IID_IUnknown || riid == IID_IDispatch) {
			*ppv = static_cast<IDispatch*>(this);
		}
		else if (riid == m_eventIID) {
			// THE LIE: we claim to be the event interface
			*ppv = static_cast<IDispatch*>(this);
		}
		else {
			return E_NOINTERFACE;
		}
		AddRef();
		return S_OK;
	}

	ULONG STDMETHODCALLTYPE AddRef() override {
		return InterlockedIncrement(&m_refCount);
	}

	ULONG STDMETHODCALLTYPE Release() override {
		LONG r = InterlockedDecrement(&m_refCount);
		if (r == 0) delete this;
		return r;
	}

	// ---------------------- IDispatch (stub most, Invoke is the star) ----------------------
	HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* pctinfo) override {
		if (pctinfo) *pctinfo = 0;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override {
		return E_NOTIMPL;
	}

	HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override {
		return E_NOTIMPL;  // We're a sink, not called by name
	}

	HRESULT STDMETHODCALLTYPE Invoke(
		DISPID dispid,
		REFIID /*riid*/,
		LCID lcid,
		WORD /*wFlags*/,
		DISPPARAMS* pParams,
		VARIANT* pResult,
		EXCEPINFO* /*pExcepInfo*/,
		UINT* /*puArgErr*/) override
	{
		// *** ALL EVENTS LAND HERE ***
		msgf(">>> SINK INVOKE: dispid=%ld cArgs=%u", (long)dispid, pParams ? pParams->cArgs : 0);

		// 1) Resolve dispid → event name (optional but nice)
		BSTR bstrName = nullptr;
		if (m_pEventTypeInfo) {
			UINT cNames = 0;
			m_pEventTypeInfo->GetNames(dispid, &bstrName, 1, &cNames);
			// bstrName is now the event method name, e.g., "QueryClose", "Click"
		}
		if (!bstrName) {
			// Fallback: synthesize name from dispid
			wchar_t buf[32];
			wsprintfW(buf, L"Event_%d", dispid);
			bstrName = SysAllocString(buf);
		}

		msgf(">>> SINK INVOKE: eventName=%S sourceName=%S", bstrName ? bstrName : L"(null)", m_bstrSourceName ? m_bstrSourceName : L"(null)");

		// 2) Forward to callback
		if (m_pfnCallback) {
			msg(">>> SINK: calling pfnCallback");
			m_pfnCallback(m_bstrSourceName, bstrName, dispid, pParams, m_pSourceObject, m_userData);
		}
		else if (m_pDispatchCallback) {
			msg(">>> SINK: calling ForwardToDispatch");
			ForwardToDispatch(bstrName, dispid, pParams);
		}
		else {
			msg(">>> SINK: NO CALLBACK SET!");
		}

		SysFreeString(bstrName);

		if (pResult) VariantInit(pResult);
		return S_OK;
	}

private:
	void ForwardToDispatch(BSTR eventName, DISPID dispid, DISPPARAMS* pSrcParams) {
		msg("ForwardToDispatch: enter");
		if (!m_pDispatchCallback) {
			msg("ForwardToDispatch: NO CALLBACK!");
			return;
		}

		// Get dispid for "OnEvent" if not cached
		if (m_dispidOnEvent == DISPID_UNKNOWN) {
			LPOLESTR name = (wchar_t*)L"OnEvent";
			HRESULT hr = m_pDispatchCallback->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &m_dispidOnEvent);
			msgf("ForwardToDispatch: GetIDsOfNames('OnEvent') = 0x%08X dispid=%ld", hr, (long)m_dispidOnEvent);
			if (FAILED(hr)) {
				msg("ForwardToDispatch: OnEvent method not found on callback!");
				return;
			}
		}

		// Build new DISPPARAMS: [sourceName, sourceObj, eventName, dispid, ...originalArgs...]
		// Original args are in reverse order in pSrcParams->rgvarg
		UINT cSrcArgs = pSrcParams ? pSrcParams->cArgs : 0;

		// Sanity check — no legit COM event has 1000 args
		if (cSrcArgs > 256) {
			msg("ForwardToDispatch: too many args, bailing");
			return;
		}

		UINT cNewArgs = cSrcArgs + 4;  // +sourceName +sourceObj +eventName +dispid
		msgf("ForwardToDispatch: cSrcArgs=%u cNewArgs=%u", cSrcArgs, cNewArgs);

		VARIANT* rgNew = (VARIANT*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, cNewArgs * sizeof(VARIANT));
		if (!rgNew) {
			msg("ForwardToDispatch: HeapAlloc failed");
			return;  // OOM, just drop the event
		}

		// Args go in reverse: last param = index 0
		// We want: callback(sourceName, sourceObj, eventName, dispid, arg0, arg1, ...)
		// So in rgvarg: [argN-1, ..., arg0, dispid, eventName, sourceObj, sourceName]

		// Copy source args (already reversed)
		for (UINT i = 0; i < cSrcArgs; i++) {
			rgNew[i] = pSrcParams->rgvarg[i];  // shallow copy ok, we don't own
		}

		// dispid as VT_I4
		rgNew[cSrcArgs].vt = VT_I4;
		rgNew[cSrcArgs].lVal = dispid;

		// eventName as VT_BSTR
		rgNew[cSrcArgs + 1].vt = VT_BSTR;
		rgNew[cSrcArgs + 1].bstrVal = eventName;  // borrowed, don't free

		// sourceObj as VT_UNKNOWN (or VT_DISPATCH if available)
		if (m_pSourceObject) {
			IDispatch* pDisp = nullptr;
			if (SUCCEEDED(m_pSourceObject->QueryInterface(IID_IDispatch, (void**)&pDisp)) && pDisp) {
				rgNew[cSrcArgs + 2].vt = VT_DISPATCH;
				rgNew[cSrcArgs + 2].pdispVal = pDisp;  // AddRef'd by QI, caller must release
			}
			else {
				rgNew[cSrcArgs + 2].vt = VT_UNKNOWN;
				rgNew[cSrcArgs + 2].punkVal = m_pSourceObject;
				m_pSourceObject->AddRef();  // caller must release
			}
		}
		else {
			rgNew[cSrcArgs + 2].vt = VT_EMPTY;
		}

		// sourceName as VT_BSTR
		rgNew[cSrcArgs + 3].vt = VT_BSTR;
		rgNew[cSrcArgs + 3].bstrVal = m_bstrSourceName;  // borrowed, don't free

		DISPPARAMS dp = { rgNew, nullptr, cNewArgs, 0 };
		VARIANT vResult;
		VariantInit(&vResult);

		msgf("ForwardToDispatch: calling callback->Invoke(%ld)...", (long)m_dispidOnEvent);
		HRESULT hr = m_pDispatchCallback->Invoke(m_dispidOnEvent, IID_NULL, LOCALE_USER_DEFAULT,
			DISPATCH_METHOD, &dp, &vResult, nullptr, nullptr);
		msgf("ForwardToDispatch: Invoke returned 0x%08X", hr);

		VariantClear(&vResult);

		// Clean up the source object ref we added
		if (rgNew[cSrcArgs + 2].vt == VT_DISPATCH && rgNew[cSrcArgs + 2].pdispVal) {
			rgNew[cSrcArgs + 2].pdispVal->Release();
		}
		else if (rgNew[cSrcArgs + 2].vt == VT_UNKNOWN && rgNew[cSrcArgs + 2].punkVal) {
			rgNew[cSrcArgs + 2].punkVal->Release();
		}
		// Don't clear other rgNew entries — we borrowed the BSTRs

		HeapFree(GetProcessHeap(), 0, rgNew);
		msg("ForwardToDispatch: done");
	}
};


static void DumpIID(const char* label, const IID& iid) {
	LPOLESTR sz = nullptr;
	StringFromIID(iid, &sz);
	if (sz) {
		msgf("%s: %S", label, sz);
		CoTaskMemFree(sz);
	}
}

// Lookup event IID for VB6 intrinsic controls by type name
static HRESULT GetVB6IntrinsicEventIID(IUnknown* pUnk, IID* pOutIID) {
	msg("GetVB6IntrinsicEventIID: enter");

	// Get the type name using our ComTypeName helper
	VARIANT vName = ComTypeName(pUnk);
	if (vName.vt != VT_BSTR || !vName.bstrVal) {
		msg("GetVB6IntrinsicEventIID: couldn't get type name");
		VariantClear(&vName);
		return E_FAIL;
	}

	msgf("GetVB6IntrinsicEventIID: typeName=%S", vName.bstrVal);

	for (int i = 0; g_VB6Intrinsics[i].typeName != nullptr; i++) {
		if (_wcsicmp(vName.bstrVal, g_VB6Intrinsics[i].typeName) == 0) {
			*pOutIID = g_VB6Intrinsics[i].eventIID;
			DumpIID("GetVB6IntrinsicEventIID: found", *pOutIID);
			VariantClear(&vName);
			return S_OK;
		}
	}

	msgf("GetVB6IntrinsicEventIID: '%S' not in intrinsic table", vName.bstrVal);
	VariantClear(&vName);
	return E_FAIL;
}

// ========================== Helper: Enumerate Connection Points ==========================
// Fallback when IProvideClassInfo isn't available - enumerate what's there
static HRESULT GetConnectionPointByEnum(IUnknown* pUnk, IID* pOutIID, ITypeInfo** ppOutEventTI) {
	msg("GetConnectionPointByEnum: enter");

	IConnectionPointContainer* pCPC = nullptr;
	HRESULT hr = pUnk->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
	msgf("GetConnectionPointByEnum: QI CPC = 0x%08X", hr);
	if (FAILED(hr) || !pCPC) return hr;

	IEnumConnectionPoints* pEnum = nullptr;
	hr = pCPC->EnumConnectionPoints(&pEnum);
	msgf("GetConnectionPointByEnum: EnumConnectionPoints = 0x%08X", hr);
	if (FAILED(hr) || !pEnum) {
		pCPC->Release();
		return hr;
	}

	IConnectionPoint* pCP = nullptr;
	ULONG fetched = 0;
	hr = E_FAIL;
	int cpCount = 0;

	while (pEnum->Next(1, &pCP, &fetched) == S_OK && fetched == 1) {
		cpCount++;
		IID iid = IID_NULL;
		if (SUCCEEDED(pCP->GetConnectionInterface(&iid))) {
			DumpIID("GetConnectionPointByEnum: found CP", iid);

			// Take the first one (usually the default event interface)
			if (hr != S_OK) {
				*pOutIID = iid;
				if (ppOutEventTI) *ppOutEventTI = nullptr;  // No typeinfo available this way
				hr = S_OK;
				// Don't break - keep enumerating to log all available CPs
			}
		}
		pCP->Release();
	}

	msgf("GetConnectionPointByEnum: found %d connection points", cpCount);

	pEnum->Release();
	pCPC->Release();

	if (hr == S_OK) {
		msg("GetConnectionPointByEnum: success");
	}
	else {
		msg("GetConnectionPointByEnum: no connection points found");
	}
	return hr;
}

// ========================== Helper: Get Event IID from TypeLib ==========================
// Try to find the source interface by loading the object's typelib
static HRESULT GetEventIIDFromTypeLib(IUnknown* pUnk, IID* pOutIID, ITypeInfo** ppOutEventTI) {
	msg("GetEventIIDFromTypeLib: enter");

	// Get IDispatch to access type info
	IDispatch* pDisp = nullptr;
	HRESULT hr = pUnk->QueryInterface(IID_IDispatch, (void**)&pDisp);
	msgf("GetEventIIDFromTypeLib: QI IDispatch = 0x%08X", hr);
	if (FAILED(hr) || !pDisp) return hr;

	ITypeInfo* pTI = nullptr;
	hr = pDisp->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTI);
	msgf("GetEventIIDFromTypeLib: GetTypeInfo = 0x%08X", hr);
	if (FAILED(hr) || !pTI) {
		pDisp->Release();
		return hr;
	}

	// Get the interface GUID - we'll use this to search
	TYPEATTR* pTIAttr = nullptr;
	hr = pTI->GetTypeAttr(&pTIAttr);
	if (FAILED(hr) || !pTIAttr) {
		pTI->Release();
		pDisp->Release();
		return hr;
	}
	GUID dispIID = pTIAttr->guid;
	DumpIID("GetEventIIDFromTypeLib: dispatch IID", dispIID);
	pTI->ReleaseTypeAttr(pTIAttr);

	// Try to get containing typelib
	ITypeLib* pTL = nullptr;
	UINT idx = 0;
	hr = pTI->GetContainingTypeLib(&pTL, &idx);
	msgf("GetEventIIDFromTypeLib: GetContainingTypeLib = 0x%08X idx=%u", hr, idx);

	if (FAILED(hr) || !pTL) {
		// Fallback: Try loading MSVBVM60.DLL typelib for VB6 intrinsic controls
		msg("GetEventIIDFromTypeLib: trying MSVBVM60.DLL...");
		hr = LoadTypeLib(L"MSVBVM60.DLL", &pTL);
		msgf("GetEventIIDFromTypeLib: LoadTypeLib MSVBVM60 = 0x%08X", hr);

		if (FAILED(hr) || !pTL) {
			// Try full path
			wchar_t sysPath[MAX_PATH];
			GetSystemDirectoryW(sysPath, MAX_PATH);
			wcscat_s(sysPath, L"\\MSVBVM60.DLL");
			hr = LoadTypeLib(sysPath, &pTL);
			msgf("GetEventIIDFromTypeLib: LoadTypeLib full path = 0x%08X", hr);
		}

		if (FAILED(hr) || !pTL) {
			pTI->Release();
			pDisp->Release();
			return hr;
		}
	}

	pTI->Release();  // Done with the dispatch typeinfo

	// Now search the typelib for a coclass that implements our dispatch interface
	UINT count = pTL->GetTypeInfoCount();
	msgf("GetEventIIDFromTypeLib: typelib has %u types", count);

	for (UINT i = 0; i < count; i++) {
		TYPEKIND tk;
		if (FAILED(pTL->GetTypeInfoType(i, &tk))) continue;
		if (tk != TKIND_COCLASS) continue;

		ITypeInfo* pCoClassTI = nullptr;
		if (FAILED(pTL->GetTypeInfo(i, &pCoClassTI)) || !pCoClassTI) continue;

		TYPEATTR* pTA = nullptr;
		if (FAILED(pCoClassTI->GetTypeAttr(&pTA)) || !pTA) {
			pCoClassTI->Release();
			continue;
		}

		// Check if this coclass implements our dispatch interface
		bool foundDispatch = false;
		UINT sourceIdx = (UINT)-1;

		for (UINT j = 0; j < pTA->cImplTypes; j++) {
			HREFTYPE hRef = 0;
			if (FAILED(pCoClassTI->GetRefTypeOfImplType(j, &hRef))) continue;

			ITypeInfo* pImplTI = nullptr;
			if (FAILED(pCoClassTI->GetRefTypeInfo(hRef, &pImplTI)) || !pImplTI) continue;

			TYPEATTR* pImplTA = nullptr;
			if (SUCCEEDED(pImplTI->GetTypeAttr(&pImplTA)) && pImplTA) {
				// Check if this is our dispatch interface
				if (IsEqualGUID(pImplTA->guid, dispIID)) {
					foundDispatch = true;
					msgf("GetEventIIDFromTypeLib: found coclass[%u] implements our dispatch", i);
				}
				pImplTI->ReleaseTypeAttr(pImplTA);
			}

			// Also check impl flags for [default, source]
			INT implFlags = 0;
			if (SUCCEEDED(pCoClassTI->GetImplTypeFlags(j, &implFlags))) {
				if ((implFlags & IMPLTYPEFLAG_FDEFAULT) && (implFlags & IMPLTYPEFLAG_FSOURCE)) {
					sourceIdx = j;
				}
			}

			pImplTI->Release();
		}

		// If this coclass implements our dispatch AND has a source interface
		if (foundDispatch && sourceIdx != (UINT)-1) {
			msgf("GetEventIIDFromTypeLib: coclass[%u] has source at impl[%u]", i, sourceIdx);

			HREFTYPE hRef = 0;
			if (SUCCEEDED(pCoClassTI->GetRefTypeOfImplType(sourceIdx, &hRef))) {
				ITypeInfo* pEventTI = nullptr;
				if (SUCCEEDED(pCoClassTI->GetRefTypeInfo(hRef, &pEventTI)) && pEventTI) {
					TYPEATTR* pEventTA = nullptr;
					if (SUCCEEDED(pEventTI->GetTypeAttr(&pEventTA)) && pEventTA) {
						*pOutIID = pEventTA->guid;
						DumpIID("GetEventIIDFromTypeLib: event IID", *pOutIID);
						pEventTI->ReleaseTypeAttr(pEventTA);
						if (ppOutEventTI) {
							*ppOutEventTI = pEventTI;
						}
						else {
							pEventTI->Release();
						}
						pCoClassTI->ReleaseTypeAttr(pTA);
						pCoClassTI->Release();
						pTL->Release();
						pDisp->Release();
						msg("GetEventIIDFromTypeLib: success");
						return S_OK;
					}
					pEventTI->Release();
				}
			}
		}

		pCoClassTI->ReleaseTypeAttr(pTA);
		pCoClassTI->Release();
	}

	pTL->Release();
	pDisp->Release();
	msg("GetEventIIDFromTypeLib: no event interface found");
	return E_FAIL;
}

// ========================== Helper: Find Default Source Interface ==========================
// Walks the coclass typeinfo to find [default, source] interface IID
static HRESULT GetDefaultSourceIID(IUnknown* pUnk, IID* pOutIID, ITypeInfo** ppOutEventTI) {
	msg("GetDefaultSourceIID: enter");

	if (!pUnk || !pOutIID) {
		msg("GetDefaultSourceIID: E_POINTER");
		return E_POINTER;
	}
	*pOutIID = IID_NULL;
	if (ppOutEventTI) *ppOutEventTI = nullptr;

	// Try IProvideClassInfo2 first (easy path)
	msg("GetDefaultSourceIID: trying IProvideClassInfo2...");
	IProvideClassInfo2* pci2 = nullptr;
	HRESULT hrQI = pUnk->QueryInterface(IID_IProvideClassInfo2, (void**)&pci2);
	msgf("GetDefaultSourceIID: QI IProvideClassInfo2 = 0x%08X", hrQI);

	if (SUCCEEDED(hrQI) && pci2) {
		HRESULT hr = pci2->GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID, pOutIID);
		msgf("GetDefaultSourceIID: GetGUID = 0x%08X", hr);
		if (SUCCEEDED(hr) && *pOutIID != IID_NULL) {
			DumpIID("GetDefaultSourceIID: found event IID", *pOutIID);
			// Now get the ITypeInfo for this IID from the typelib
			ITypeInfo* pCoClassTI = nullptr;
			if (SUCCEEDED(pci2->GetClassInfo(&pCoClassTI)) && pCoClassTI) {
				ITypeLib* pTL = nullptr;
				UINT idx = 0;
				if (SUCCEEDED(pCoClassTI->GetContainingTypeLib(&pTL, &idx)) && pTL) {
					ITypeInfo* pEventTI = nullptr;
					if (SUCCEEDED(pTL->GetTypeInfoOfGuid(*pOutIID, &pEventTI)) && pEventTI) {
						msg("GetDefaultSourceIID: got event ITypeInfo");
						if (ppOutEventTI) *ppOutEventTI = pEventTI;
						else pEventTI->Release();
					}
					pTL->Release();
				}
				pCoClassTI->Release();
			}
			pci2->Release();
			msg("GetDefaultSourceIID: success via IProvideClassInfo2");
			return S_OK;
		}
		pci2->Release();
	}

	// Fallback: IProvideClassInfo → walk impl types
	msg("GetDefaultSourceIID: trying IProvideClassInfo...");
	IProvideClassInfo* pci = nullptr;
	hrQI = pUnk->QueryInterface(IID_IProvideClassInfo, (void**)&pci);
	msgf("GetDefaultSourceIID: QI IProvideClassInfo = 0x%08X", hrQI);

	if (SUCCEEDED(hrQI) && pci) {
		ITypeInfo* pCoClassTI = nullptr;
		HRESULT hr = pci->GetClassInfo(&pCoClassTI);
		msgf("GetDefaultSourceIID: GetClassInfo = 0x%08X", hr);
		pci->Release();

		if (SUCCEEDED(hr) && pCoClassTI) {
			TYPEATTR* pTA = nullptr;
			hr = pCoClassTI->GetTypeAttr(&pTA);
			if (SUCCEEDED(hr) && pTA) {
				msgf("GetDefaultSourceIID: scanning %d impl types...", pTA->cImplTypes);

				for (UINT i = 0; i < pTA->cImplTypes; i++) {
					INT implFlags = 0;
					if (FAILED(pCoClassTI->GetImplTypeFlags(i, &implFlags))) continue;

					msgf("GetDefaultSourceIID: impl[%d] flags=0x%04X", i, implFlags);

					// Looking for [default, source]
					if ((implFlags & IMPLTYPEFLAG_FDEFAULT) && (implFlags & IMPLTYPEFLAG_FSOURCE)) {
						msg("GetDefaultSourceIID: found [default, source]!");
						HREFTYPE hRef = 0;
						if (SUCCEEDED(pCoClassTI->GetRefTypeOfImplType(i, &hRef))) {
							ITypeInfo* pEventTI = nullptr;
							if (SUCCEEDED(pCoClassTI->GetRefTypeInfo(hRef, &pEventTI)) && pEventTI) {
								TYPEATTR* pEventTA = nullptr;
								if (SUCCEEDED(pEventTI->GetTypeAttr(&pEventTA)) && pEventTA) {
									*pOutIID = pEventTA->guid;
									DumpIID("GetDefaultSourceIID: event IID", *pOutIID);
									pEventTI->ReleaseTypeAttr(pEventTA);
									if (ppOutEventTI) {
										*ppOutEventTI = pEventTI;  // caller owns
									}
									else {
										pEventTI->Release();
									}
									pCoClassTI->ReleaseTypeAttr(pTA);
									pCoClassTI->Release();
									msg("GetDefaultSourceIID: success via IProvideClassInfo");
									return S_OK;
								}
								else {
									pEventTI->Release();
								}
							}
						}
						break;  // Found default source
					}
				}
				pCoClassTI->ReleaseTypeAttr(pTA);
			}
			pCoClassTI->Release();
		}
	}

	// Fallback: Enumerate connection points directly
	msg("GetDefaultSourceIID: trying connection point enumeration...");
	HRESULT hr = GetConnectionPointByEnum(pUnk, pOutIID, ppOutEventTI);
	if (SUCCEEDED(hr)) {
		msg("GetDefaultSourceIID: success via enumeration");
		return hr;
	}
	msgf("GetDefaultSourceIID: enumeration returned 0x%08X", hr);

	// Last resort: Try to find event IID from typelib
	msg("GetDefaultSourceIID: trying typelib scan...");
	hr = GetEventIIDFromTypeLib(pUnk, pOutIID, ppOutEventTI);
	if (SUCCEEDED(hr)) {
		msg("GetDefaultSourceIID: success via typelib scan");
		return hr;
	}
	msgf("GetDefaultSourceIID: typelib scan returned 0x%08X", hr);

	// Final fallback : VB6 intrinsic control lookup
	// DISABLED: VB6 intrinsics use vtable-based events, not IDispatch. See header comment.
	// msg("GetDefaultSourceIID: trying VB6 intrinsic lookup...");
	// hr = GetVB6IntrinsicEventIID(pUnk, pOutIID);
	// if (SUCCEEDED(hr)) {
	//     if (ppOutEventTI) *ppOutEventTI = nullptr;
	//     msg("GetDefaultSourceIID: success via VB6 intrinsic lookup");
	//     return hr;
	// }
	// msgf("GetDefaultSourceIID: VB6 intrinsic lookup returned 0x%08X", hr);


	return E_FAIL;
}

// ========================== Connection Point Helpers ==========================
struct EventConnection {
	IConnectionPoint* pCP = nullptr;
	DWORD dwCookie = 0;
	UniversalEventSink* pSink = nullptr;
};

static HRESULT ConnectEvents(IUnknown* pSource, const IID& eventIID, ITypeInfo* pEventTI,
	PFN_EVENT_CALLBACK pfn, void* userData,
	BSTR sourceName, EventConnection* pConn)
{
	if (!pSource || !pConn) return E_POINTER;
	memset(pConn, 0, sizeof(*pConn));

	IConnectionPointContainer* pCPC = nullptr;
	HRESULT hr = pSource->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
	if (FAILED(hr) || !pCPC) return hr;

	IConnectionPoint* pCP = nullptr;
	hr = pCPC->FindConnectionPoint(eventIID, &pCP);
	pCPC->Release();
	if (FAILED(hr) || !pCP) return hr;

	UniversalEventSink* pSink = new UniversalEventSink(eventIID, pEventTI, pfn, userData, sourceName, pSource);

	DWORD dwCookie = 0;
	hr = pCP->Advise(pSink, &dwCookie);
	if (FAILED(hr)) {
		pSink->Release();
		pCP->Release();
		return hr;
	}

	pConn->pCP = pCP;
	pConn->dwCookie = dwCookie;
	pConn->pSink = pSink;
	return S_OK;
}

// Overload for IDispatch callback (your JS wrapper)
static HRESULT ConnectEventsDispatch(IUnknown* pSource, const IID& eventIID, ITypeInfo* pEventTI,
	IDispatch* pCallback, BSTR sourceName, EventConnection* pConn)
{
	msg("ConnectEventsDispatch: enter");
	if (!pSource || !pConn) return E_POINTER;
	memset(pConn, 0, sizeof(*pConn));

	msg("ConnectEventsDispatch: QI for IConnectionPointContainer...");
	IConnectionPointContainer* pCPC = nullptr;
	HRESULT hr = pSource->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
	msgf("ConnectEventsDispatch: QI IConnectionPointContainer = 0x%08X", hr);
	if (FAILED(hr) || !pCPC) return hr;

	msg("ConnectEventsDispatch: FindConnectionPoint...");
	DumpIID("ConnectEventsDispatch: looking for", eventIID);
	IConnectionPoint* pCP = nullptr;
	hr = pCPC->FindConnectionPoint(eventIID, &pCP);
	msgf("ConnectEventsDispatch: FindConnectionPoint = 0x%08X", hr);
	pCPC->Release();
	if (FAILED(hr) || !pCP) return hr;

	msg("ConnectEventsDispatch: creating sink...");
	UniversalEventSink* pSink = new UniversalEventSink(eventIID, pEventTI, pCallback, sourceName, pSource);

	msg("ConnectEventsDispatch: Advise...");
	DWORD dwCookie = 0;
	hr = pCP->Advise(pSink, &dwCookie);
	msgf("ConnectEventsDispatch: Advise = 0x%08X, cookie=%u", hr, dwCookie);
	if (FAILED(hr)) {
		pSink->Release();
		pCP->Release();
		return hr;
	}

	pConn->pCP = pCP;
	pConn->dwCookie = dwCookie;
	pConn->pSink = pSink;
	msg("ConnectEventsDispatch: success");
	return S_OK;
}

static void DisconnectEvents(EventConnection* pConn) {
	if (!pConn) return;
	if (pConn->pCP && pConn->dwCookie) {
		pConn->pCP->Unadvise(pConn->dwCookie);
	}
	if (pConn->pCP) pConn->pCP->Release();
	if (pConn->pSink) pConn->pSink->Release();
	memset(pConn, 0, sizeof(*pConn));
}

// ========================== Exports for VB6 / Your Engine ==========================
// Opaque handle for connection
typedef struct EventConnectionHandle {
	EventConnection conn;
	IID eventIID;
	ITypeInfo* pEventTI;
	BSTR bstrSourceName;  // Keep copy for GetSourceName export
} EventConnectionHandle;

extern "C" {

	// Auto-discover default source and connect
	// Uses raw pointers like CreateProxyForObjectRaw pattern
	__declspec(dllexport) HRESULT __stdcall SinkEventsAuto(
		ULONG_PTR pSourceRaw,
		BSTR sourceName,
		ULONG_PTR pCallbackRaw,
		ULONG_PTR* ppHandle)  // Return handle as ULONG_PTR*
	{
		msg("=== SinkEventsAuto ===");
		msgf("SinkEventsAuto: pSourceRaw=0x%08X, pCallbackRaw=0x%08X", (DWORD)pSourceRaw, (DWORD)pCallbackRaw);
		if (sourceName) msgf("SinkEventsAuto: sourceName=%S", sourceName);

		if (!pSourceRaw || !pCallbackRaw || !ppHandle) {
			msg("SinkEventsAuto: E_POINTER (null arg)");
			return E_POINTER;
		}
		*ppHandle = 0;

		// QI source for IUnknown
		IUnknown* pSource = reinterpret_cast<IUnknown*>(pSourceRaw);

		// QI callback for IDispatch (like QI_IDispatch_FromRawPtr pattern)
		msg("SinkEventsAuto: QI callback for IDispatch...");
		IUnknown* pCallbackUnk = reinterpret_cast<IUnknown*>(pCallbackRaw);
		IDispatch* pCallback = nullptr;
		HRESULT hr = pCallbackUnk->QueryInterface(IID_IDispatch, (void**)&pCallback);
		msgf("SinkEventsAuto: QI IDispatch = 0x%08X", hr);
		if (FAILED(hr) || !pCallback) {
			msg("SinkEventsAuto: callback doesn't support IDispatch!");
			return hr;
		}

		IID eventIID = IID_NULL;
		ITypeInfo* pEventTI = nullptr;
		hr = GetDefaultSourceIID(pSource, &eventIID, &pEventTI);
		msgf("SinkEventsAuto: GetDefaultSourceIID returned 0x%08X", hr);
		if (FAILED(hr)) {
			pCallback->Release();
			return hr;
		}

		EventConnectionHandle* h = new EventConnectionHandle();
		h->eventIID = eventIID;
		h->pEventTI = pEventTI;  // Takes ownership
		h->bstrSourceName = sourceName ? SysAllocString(sourceName) : nullptr;

		hr = ConnectEventsDispatch(pSource, eventIID, pEventTI, pCallback, sourceName, &h->conn);
		msgf("SinkEventsAuto: ConnectEventsDispatch returned 0x%08X", hr);

		pCallback->Release();  // ConnectEventsDispatch AddRefs if it needs it

		if (FAILED(hr)) {
			if (pEventTI) pEventTI->Release();
			if (h->bstrSourceName) SysFreeString(h->bstrSourceName);
			delete h;
			return hr;
		}

		*ppHandle = (ULONG_PTR)h;
		msg("SinkEventsAuto: success!");
		return S_OK;
	}

	// Connect to specific IID (if you know it)
	__declspec(dllexport) HRESULT __stdcall SinkEventsIID(
		IUnknown* pSource,
		GUID* pEventIID,
		BSTR sourceName,
		IUnknown* pCallbackUnk,   // VB6 passes Object as IUnknown*
		void** ppHandle)
	{
		if (!pSource || !pEventIID || !pCallbackUnk || !ppHandle) return E_POINTER;
		*ppHandle = nullptr;

		// QI callback for IDispatch
		IDispatch* pCallback = nullptr;
		HRESULT hr = pCallbackUnk->QueryInterface(IID_IDispatch, (void**)&pCallback);
		if (FAILED(hr) || !pCallback) return hr;

		EventConnectionHandle* h = new EventConnectionHandle();
		h->eventIID = *pEventIID;
		h->pEventTI = nullptr;  // No type info for name resolution
		h->bstrSourceName = sourceName ? SysAllocString(sourceName) : nullptr;

		hr = ConnectEventsDispatch(pSource, *pEventIID, nullptr, pCallback, sourceName, &h->conn);

		pCallback->Release();

		if (FAILED(hr)) {
			if (h->bstrSourceName) SysFreeString(h->bstrSourceName);
			delete h;
			return hr;
		}

		*ppHandle = h;
		return S_OK;
	}

	// Disconnect and free
	__declspec(dllexport) HRESULT __stdcall SinkEventsDisconnect(void* pHandle) {
		if (!pHandle) return S_OK;
		EventConnectionHandle* h = (EventConnectionHandle*)pHandle;
		DisconnectEvents(&h->conn);
		if (h->pEventTI) h->pEventTI->Release();
		if (h->bstrSourceName) SysFreeString(h->bstrSourceName);
		delete h;
		return S_OK;
	}

	// Get the event IID we connected to (informational)
	__declspec(dllexport) HRESULT __stdcall SinkEventsGetIID(void* pHandle, GUID* pOutIID) {
		if (!pHandle || !pOutIID) return E_POINTER;
		EventConnectionHandle* h = (EventConnectionHandle*)pHandle;
		*pOutIID = h->eventIID;
		return S_OK;
	}

	// Get the source name (informational)
	__declspec(dllexport) HRESULT __stdcall SinkEventsGetSourceName(void* pHandle, BSTR* pOutName) {
		if (!pHandle || !pOutName) return E_POINTER;
		EventConnectionHandle* h = (EventConnectionHandle*)pHandle;
		*pOutName = h->bstrSourceName ? SysAllocString(h->bstrSourceName) : nullptr;
		return S_OK;
	}

}  // extern "C"

// Dead simple test - does VB6 even call into us?
extern "C" __declspec(dllexport) HRESULT __stdcall SinkTest(void) {
	msg("SinkTest called!");
	return S_OK;
}

// Test with raw Long pointers - bypass VB6 COM marshaling
extern "C" __declspec(dllexport) HRESULT __stdcall SinkTestRaw(
	LONG_PTR pSourceRaw,
	BSTR sourceName,
	LONG_PTR pCallbackRaw,
	void** ppHandle)
{
	msg("SinkTestRaw: enter");
	msgf("SinkTestRaw: pSourceRaw=0x%08X", (DWORD)pSourceRaw);
	msgf("SinkTestRaw: pCallbackRaw=0x%08X", (DWORD)pCallbackRaw);
	if (sourceName) msgf("SinkTestRaw: sourceName=%S", sourceName);
	msgf("SinkTestRaw: ppHandle=0x%p", ppHandle);

	if (ppHandle) *ppHandle = (void*)0xDEADBEEF;

	msg("SinkTestRaw: exit");
	return S_OK;
}

// Test each param individually
extern "C" __declspec(dllexport) HRESULT __stdcall SinkTest1(IUnknown* pSource) {
	msgf("SinkTest1: pSource=0x%p", pSource);
	return S_OK;
}

extern "C" __declspec(dllexport) HRESULT __stdcall SinkTest2(IUnknown* pSource, BSTR sourceName) {
	msgf("SinkTest2: pSource=0x%p", pSource);
	if (sourceName) msgf("SinkTest2: sourceName=%S", sourceName);
	else msg("SinkTest2: sourceName=NULL");
	return S_OK;
}

extern "C" __declspec(dllexport) HRESULT __stdcall SinkTest3(IUnknown* pSource, BSTR sourceName, IUnknown* pCallback) {
	msgf("SinkTest3: pSource=0x%p, pCallback=0x%p", pSource, pCallback);
	return S_OK;
}

extern "C" __declspec(dllexport) HRESULT __stdcall SinkTest4(IUnknown* pSource, BSTR sourceName, IUnknown* pCallback, void** ppHandle) {
	msgf("SinkTest4: pSource=0x%p, pCallback=0x%p, ppHandle=0x%p", pSource, pCallback, ppHandle);
	if (ppHandle) *ppHandle = (void*)0x12345678;
	return S_OK;
}

/*
=== VB6 Usage ===

Private Declare Function SinkEventsAuto Lib "ComEventSink.dll" _
	(ByVal pSource As IUnknown, ByVal sourceName As String, _
	 ByVal pCallback As IDispatch, ByRef ppHandle As LongPtr) As Long
Private Declare Function SinkEventsDisconnect Lib "ComEventSink.dll" (ByVal pHandle As LongPtr) As Long

' Hook up multiple objects with different names
Dim hSinkExcel As LongPtr, hSinkWord As LongPtr
hr = SinkEventsAuto(xlApp, "xlApp", myEventRouter, hSinkExcel)
hr = SinkEventsAuto(wdApp, "wdApp", myEventRouter, hSinkWord)

' Same callback handles both — sourceName distinguishes them!

' Later:
SinkEventsDisconnect hSinkExcel
SinkEventsDisconnect hSinkWord


=== Callback receives (in order): ===

  arg[n+3] = sourceName  (BSTR)   ← "xlApp", "wdApp", etc.
  arg[n+2] = sourceObj   (IDispatch/IUnknown) ← the actual COM object
  arg[n+1] = eventName   (BSTR)   ← "QueryClose", "Click", etc.
  arg[n]   = dispid      (Long)
  arg[0..n-1] = original event params (reversed)


=== In your JS Engine ===

Your master event router function:

function eventRouter(sourceName, sourceObj, eventName, dispid, ...args) {
	// Look up the JS variable by sourceName
	var target = globals[sourceName];  // e.g., globals["xlApp"]

	// Check for handler: xlApp.onQueryClose or xlApp.QueryClose
	var handler = target["on" + eventName] || target[eventName];
	if (typeof handler === "function") {
		handler.apply(target, args);
	}
}

Or with your vtCom extension pattern:

// When sinking, the sourceName IS the JS variable name
// So events auto-route to: xlApp.onWorkbookOpen(wb)

*/