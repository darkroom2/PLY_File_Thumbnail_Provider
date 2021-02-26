#include <objbase.h>
#include <Shlwapi.h>
#include <thumbcache.h> // For IThumbnailProvider.
#include <ShlObj.h>     // For SHChangeNotify
#include <new>

extern HRESULT CPointCloudThumbProvider_CreateInstance(REFIID riid, void** ppv);

#define SZ_CLSID_POINTCLOUDTHUMBHANDLER L"{dc45e622-49ba-4f08-bda3-28bd4db772b0}"
#define SZ_POINTCLOUDTHUMBHANDLER       L"Point Cloud Thumbnail Handler"

const CLSID CLSID_PointCloudThumbHandler = {
        0xdc45e622, 0x49ba, 0x4f08,
        {0xbd, 0xa3, 0x28, 0xbd, 0x4d, 0xb7, 0x72, 0xb0}
};

typedef HRESULT(*PFNCREATEINSTANCE)(REFIID riid, void** ppvObject);

struct CLASS_OBJECT_INIT {
    const CLSID* pClsid;
    PFNCREATEINSTANCE pfnCreate;
};

// add classes supported by this module here
const CLASS_OBJECT_INIT c_rgClassObjectInit[] = {
        {&CLSID_PointCloudThumbHandler, CPointCloudThumbProvider_CreateInstance}
};

long g_cRefModule = 0;

// Handle the the DLL's module
HINSTANCE g_hInst = nullptr;

// Standard DLL functions
STDAPI_(BOOL) DllMain(HINSTANCE hInstance, DWORD dwReason, void*) {
    if (dwReason == DLL_PROCESS_ATTACH) {
        g_hInst = hInstance;
        DisableThreadLibraryCalls(hInstance);
    }
    return TRUE;
}

STDAPI DllCanUnloadNow() {
    // Only allow the DLL to be unloaded after all outstanding references have been released
    return (g_cRefModule == 0) ? S_OK : S_FALSE;
}

void DllAddRef() {
    InterlockedIncrement(&g_cRefModule);
}

void DllRelease() {
    InterlockedDecrement(&g_cRefModule);
}

class CClassFactory : public IClassFactory {
public:
    static HRESULT
        CreateInstance(REFCLSID clsid, const CLASS_OBJECT_INIT* pClassObjectInits, size_t cClassObjectInits,
            REFIID riid, void** ppv) {
        *ppv = nullptr;
        HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;
        for (size_t i = 0; i < cClassObjectInits; i++) {
            if (clsid == *pClassObjectInits[i].pClsid) {
                IClassFactory* pClassFactory = new(std::nothrow) CClassFactory(pClassObjectInits[i].pfnCreate);
                hr = pClassFactory ? S_OK : E_OUTOFMEMORY;
                if (SUCCEEDED(hr)) {
                    hr = pClassFactory->QueryInterface(riid, ppv);
                    pClassFactory->Release();
                }
                break; // match found
            }
        }
        return hr;
    }

    explicit CClassFactory(PFNCREATEINSTANCE pfnCreate) : _cRef(1), _pfnCreate(pfnCreate) {
        DllAddRef();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        static const QITAB qit[] = {
                QITABENT(CClassFactory, IClassFactory),
                {nullptr}
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        long cRef = InterlockedDecrement(&_cRef);
        if (cRef == 0) {
            delete this;
        }
        return cRef;
    }

    // IClassFactory
    IFACEMETHODIMP CreateInstance(IUnknown* punkOuter, REFIID riid, void** ppv) override {
        return punkOuter ? CLASS_E_NOAGGREGATION : _pfnCreate(riid, ppv);
    }

    IFACEMETHODIMP LockServer(BOOL fLock) override {
        if (fLock) {
            DllAddRef();
        }
        else {
            DllRelease();
        }
        return S_OK;
    }

private:
    ~CClassFactory() {
        DllRelease();
    }

    long _cRef;
    PFNCREATEINSTANCE _pfnCreate;
};

STDAPI DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv) {
    return CClassFactory::CreateInstance(clsid, c_rgClassObjectInit, ARRAYSIZE(c_rgClassObjectInit), riid, ppv);
}

// A struct to hold the information required for a registry entry
struct REGISTRY_ENTRY {
    HKEY hkeyRoot;
    PCWSTR pszKeyName;
    PCWSTR pszValueName;
    PCWSTR pszData;
};

// Creates a registry key (if needed) and sets the default value of the key
HRESULT CreateRegKeyAndSetValue(const REGISTRY_ENTRY* pRegistryEntry) {
    HKEY hKey;
    HRESULT hr = HRESULT_FROM_WIN32(
        RegCreateKeyExW(pRegistryEntry->hkeyRoot, pRegistryEntry->pszKeyName, 0, nullptr, REG_OPTION_NON_VOLATILE,
            KEY_SET_VALUE, nullptr, &hKey, nullptr));
    if (SUCCEEDED(hr)) {
        hr = HRESULT_FROM_WIN32(
            RegSetValueExW(hKey, pRegistryEntry->pszValueName, 0, REG_SZ, (LPBYTE)pRegistryEntry->pszData,
                ((DWORD)wcslen(pRegistryEntry->pszData) + 1) * sizeof(WCHAR)));
        RegCloseKey(hKey);
    }
    return hr;
}

//
// Rejestracja serwera COM
//
STDAPI DllRegisterServer() {
    HRESULT hr;

    WCHAR szModuleName[MAX_PATH];

    if (!GetModuleFileNameW(g_hInst, szModuleName, ARRAYSIZE(szModuleName))) {
        hr = HRESULT_FROM_WIN32(GetLastError());
    }
    else {
        // Lista wpisów do rejestru, które chcemy utworzyæ
        const REGISTRY_ENTRY rgRegistryEntries[] = {
                {HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_POINTCLOUDTHUMBHANDLER,                     nullptr, SZ_POINTCLOUDTHUMBHANDLER},
                {HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_POINTCLOUDTHUMBHANDLER L"\\InProcServer32", nullptr,           szModuleName},
                {HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\" SZ_CLSID_POINTCLOUDTHUMBHANDLER L"\\InProcServer32", L"ThreadingModel", L"Apartment"},
                {HKEY_CURRENT_USER, L"Software\\Classes\\.ply\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}",       nullptr, SZ_CLSID_POINTCLOUDTHUMBHANDLER}
        };

        hr = S_OK;
        for (int i = 0; i < ARRAYSIZE(rgRegistryEntries) && SUCCEEDED(hr); i++) {
            hr = CreateRegKeyAndSetValue(&rgRegistryEntries[i]);
        }
    }
    if (SUCCEEDED(hr)) {
        // To wywo³anie mówi pow³oce, aby uniewa¿niæ pamiêæ podrêczn¹ miniatur.
        // Jest to wa¿ne, poniewa¿ wszelkie pliki *.ply przegl¹dane przed rejestracj¹ tego handlera pokazywa³yby zbuforowane puste miniatury.
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);
    }
    return hr;
}

//
// Unregisters this COM server
//
STDAPI DllUnregisterServer() {
    HRESULT hr = S_OK;

    const PCWSTR rgpszKeys[] = {
            L"Software\\Classes\\CLSID\\" SZ_CLSID_POINTCLOUDTHUMBHANDLER,
            L"Software\\Classes\\.ply\\ShellEx\\{e357fccd-a995-4576-b01f-234630154e96}"
    };

    // Delete the registry entries
    for (int i = 0; i < ARRAYSIZE(rgpszKeys) && SUCCEEDED(hr); i++) {
        hr = HRESULT_FROM_WIN32(RegDeleteTreeW(HKEY_CURRENT_USER, rgpszKeys[i]));
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) {
            // If the registry entry has already been deleted, say S_OK.
            hr = S_OK;
        }
    }
    return hr;
}
