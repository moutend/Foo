#include <shlwapi.h>
#include <strsafe.h>
#include <windows.h>

#include "Foo.h"

const TCHAR ProgIDStr[] = TEXT("Foo v0.1.0");
LONG LockCount{};
HINSTANCE FooDLLInstance{};
TCHAR FooCLSIDStr[256]{};
TCHAR LibraryIDStr[256]{};

// CFoo
CFoo::CFoo() : mReferenceCount(1), mTypeInfo(nullptr) {
  ITypeLib *pTypeLib{};
  HRESULT hr{};

  LockModule(true);

  hr = LoadRegTypeLib(LIBID_FooLib, 1, 0, 0, &pTypeLib);

  if (SUCCEEDED(hr)) {
    pTypeLib->GetTypeInfoOfGuid(IID_IFoo, &mTypeInfo);
    pTypeLib->Release();
  }
}

CFoo::~CFoo() {
  LockModule(false);
}

STDMETHODIMP CFoo::QueryInterface(REFIID riid, void **ppvObject) {
  *ppvObject = nullptr;

  if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch) ||
      IsEqualIID(riid, IID_IFoo)) {
    *ppvObject = static_cast<IFoo *>(this);
  } else {
    return E_NOINTERFACE;
  }

  AddRef();

  return S_OK;
}

STDMETHODIMP_(ULONG) CFoo::AddRef() {
  return InterlockedIncrement(&mReferenceCount);
}

STDMETHODIMP_(ULONG) CFoo::Release() {
  if (InterlockedDecrement(&mReferenceCount) == 0) {
    delete this;

    return 0;
  }

  return mReferenceCount;
}

STDMETHODIMP CFoo::GetTypeInfoCount(UINT *pctinfo) {
  *pctinfo = mTypeInfo != nullptr ? 1 : 0;

  return S_OK;
}

STDMETHODIMP CFoo::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo) {
  if (mTypeInfo == nullptr) {
    return E_NOTIMPL;
  }

  mTypeInfo->AddRef();
  *ppTInfo = mTypeInfo;

  return S_OK;
}

STDMETHODIMP CFoo::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames,
                                 LCID lcid, DISPID *rgDispId) {
  if (mTypeInfo == nullptr) {
    return E_NOTIMPL;
  }

  return mTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP CFoo::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                          WORD wFlags, DISPPARAMS *pDispParams,
                          VARIANT *pVarResult, EXCEPINFO *pExcepInfo,
                          UINT *puArgErr) {
  void *p = static_cast<IFoo *>(this);

  if (mTypeInfo == nullptr) {
    return E_NOTIMPL;
  }

  return mTypeInfo->Invoke(p, dispIdMember, wFlags, pDispParams, pVarResult,
                           pExcepInfo, puArgErr);
}

STDMETHODIMP CFoo::Start() { return E_NOTIMPL; }

STDMETHODIMP CFoo::Stop() { return E_NOTIMPL; }

STDMETHODIMP CFoo::SetUIEventHandler(UIEventHandler uiEventHandler, AvailableAccessibilityAPI eAPI) {
  return E_NOTIMPL;
}

// CFooFactory
STDMETHODIMP CFooFactory::QueryInterface(REFIID riid, void **ppvObject) {
  *ppvObject = nullptr;

  if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory)) {
    *ppvObject = static_cast<IClassFactory *>(this);
  } else {
    return E_NOINTERFACE;
  }

  AddRef();

  return S_OK;
}

STDMETHODIMP_(ULONG) CFooFactory::AddRef() {
  LockModule(true);

  return 2;
}

STDMETHODIMP_(ULONG) CFooFactory::Release() {
  LockModule(false);

  return 1;
}

STDMETHODIMP CFooFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid,
                                         void **ppvObject) {
  CFoo *p{};
  HRESULT hr{};

  *ppvObject = nullptr;

  if (pUnkOuter != nullptr) {
    return CLASS_E_NOAGGREGATION;
  }

  p = new CFoo();

  if (p == nullptr) {
    return E_OUTOFMEMORY;
  }

  hr = p->QueryInterface(riid, ppvObject);
  p->Release();

  return hr;
}

STDMETHODIMP CFooFactory::LockServer(BOOL fLock) {
  LockModule(fLock);

  return S_OK;
}

// DLL Export
STDAPI DllCanUnloadNow(void) { return LockCount == 0 ? S_OK : S_FALSE; }

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv) {
  static CFooFactory serverFactory;
  HRESULT hr{};
  *ppv = nullptr;

  if (IsEqualCLSID(rclsid, CLSID_Foo)) {
    hr = serverFactory.QueryInterface(riid, ppv);
  } else {
    hr = CLASS_E_CLASSNOTAVAILABLE;
  }

  return hr;
}

STDAPI DllRegisterServer(void) {
  TCHAR szModulePath[256]{};
  TCHAR szKey[256]{};
  WCHAR szTypeLibPath[256]{};
  HRESULT hr{};
  ITypeLib *pTypeLib{};

  wsprintf(szKey, TEXT("CLSID\\%s"), FooCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         TEXT("Foo v0.1.0"))) {
    return E_FAIL;
  }

  GetModuleFileName(FooDLLInstance, szModulePath,
                    sizeof(szModulePath) / sizeof(TCHAR));
  wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), FooCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr, szModulePath)) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), FooCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, TEXT("ThreadingModel"),
                         TEXT("Apartment"))) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("CLSID\\%s\\ProgID"), FooCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         (LPTSTR)ProgIDStr)) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("%s"), ProgIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         TEXT("Foo v0.1.0"))) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("%s\\CLSID"), ProgIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         (LPTSTR)FooCLSIDStr)) {
    return E_FAIL;
  }

  GetModuleFileNameW(FooDLLInstance, szTypeLibPath,
                     sizeof(szTypeLibPath) / sizeof(TCHAR));

  hr = LoadTypeLib(szTypeLibPath, &pTypeLib);

  if (SUCCEEDED(hr)) {
    hr = RegisterTypeLib(pTypeLib, szTypeLibPath, nullptr);

    if (SUCCEEDED(hr)) {
      wsprintf(szKey, TEXT("CLSID\\%s\\TypeLib"), FooCLSIDStr);

      if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr, LibraryIDStr)) {
        pTypeLib->Release();

        return E_FAIL;
      }
    }

    pTypeLib->Release();
  }

  return S_OK;
}

STDAPI DllUnregisterServer(void) {
  TCHAR szKey[256]{};

  wsprintf(szKey, TEXT("CLSID\\%s"), FooCLSIDStr);
  SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

  wsprintf(szKey, TEXT("%s"), ProgIDStr);
  SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

  UnRegisterTypeLib(LIBID_FooLib, 1, 0, LOCALE_NEUTRAL, SYS_WIN32);

  return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved) {
  switch (dwReason) {
  case DLL_PROCESS_ATTACH:
    FooDLLInstance = hInstance;
    DisableThreadLibraryCalls(hInstance);
    GetGuidString(CLSID_Foo, FooCLSIDStr);
    GetGuidString(LIBID_FooLib, LibraryIDStr);

    return TRUE;
  }

  return TRUE;
}

// Helper function
void LockModule(BOOL bLock) {
  if (bLock) {
    InterlockedIncrement(&LockCount);
  } else {
    InterlockedDecrement(&LockCount);
  }
}

BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue,
                       LPTSTR lpszData) {
  HKEY hKey{};
  LONG lResult{};
  DWORD dwSize{};

  lResult =
      RegCreateKeyEx(hKeyRoot, lpszKey, 0, nullptr, REG_OPTION_NON_VOLATILE,
                     KEY_WRITE, nullptr, &hKey, nullptr);

  if (lResult != ERROR_SUCCESS) {
    return FALSE;
  }
  if (lpszData != nullptr) {
    dwSize = (lstrlen(lpszData) + 1) * sizeof(TCHAR);
  } else {
    dwSize = 0;
  }

  RegSetValueEx(hKey, lpszValue, 0, REG_SZ, (LPBYTE)lpszData, dwSize);
  RegCloseKey(hKey);

  return TRUE;
}

void GetGuidString(REFGUID rguid, LPTSTR lpszGuid) {
  LPWSTR lpsz;
  StringFromCLSID(rguid, &lpsz);

#ifdef UNICODE
  lstrcpyW(lpszGuid, lpsz);
#else
  WideCharToMultiByte(CP_ACP, 0, lpsz, -1, lpszGuid, 256, nullptr, nullptr);
#endif

  CoTaskMemFree(lpsz);
}
