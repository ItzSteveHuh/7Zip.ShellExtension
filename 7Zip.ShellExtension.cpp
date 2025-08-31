
// 7Zip.ShellExtension.cpp
// In-proc COM server (DLL) implementing Windows 11 modern ExplorerCommand for 7-Zip.
//
// Build (MSVC x64):
//   cl /EHsc /W4 /permissive- /std:c++17 /LD 7Zip.ShellExtension.cpp ^
//      Ole32.lib Shell32.lib Shlwapi.lib
//
// If the linker doesn't export the COM entry points automatically, this file
// uses #pragma comment(linker, "/EXPORT:...") for x64 builds. Alternatively,
// you can provide a .def file exporting DllGetClassObject and DllCanUnloadNow.
//
// Manifest (MSIX) should register this DLL via <com:SurrogateServer> with the
// same CLSID as below, and hook the Windows 11 menu via
// <desktop4:Extension Category="windows.fileExplorerContextMenus"> with a single
// <desktop5:Verb Type="*"> entry pointing to this CLSID.
//
// Behavior parity:
//  - "Extract Here" is SMART for multi-archives (NanaZip-style): each archive
//    goes into its own "<ArchiveName>\\" folder to avoid mixing files.
//  - "Extract to \\<Folder>\\" always creates per-archive folders (classic).
//  - Default archive naming matches classic 7-Zip:
//      * Single item  -> <ItemName>.ext
//      * Multi items  -> <ParentName>.ext if all from same parent; else Archive.ext
//  - CRC submenu with CRC-32/CRC-64/SHA-1/SHA-256.
//  - Add/Email entries available for files/dirs/archives, like classic.
//
// NOTE: This DLL assumes 7zFM.exe, 7zG.exe, 7z.exe are either next to the DLL
//       (inside your MSIX) or discoverable via PATH.

#define NOMINMAX
#include <Windows.h>
#include <shobjidl_core.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <objbase.h>
#include <filesystem>
#include <string>
#include <vector>
#include <algorithm>

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")

// Export COM entry points on x64 without a .def (harmless if also supplied via .def)
#if defined(_M_X64) || defined(_WIN64)
#pragma comment(linker, "/EXPORT:DllGetClassObject")
#pragma comment(linker, "/EXPORT:DllCanUnloadNow")
#endif

// {7C9F3AA6-4D07-4E1D-8B86-0F5A4E4F44AC}
static const CLSID CLSID_SevenZipExplorer =
{ 0x7c9f3aa6, 0x4d07, 0x4e1d, { 0x8b, 0x86, 0x0f, 0x5a, 0x4e, 0x4f, 0x44, 0xac } };

static HMODULE g_hMod = nullptr;
static LONG g_ObjCount = 0;
static LONG g_LockCount = 0;

// ---------- helpers ----------
static std::wstring GetModuleDir(HMODULE mod) {
    wchar_t buf[MAX_PATH]{};
    GetModuleFileNameW(mod, buf, MAX_PATH);
    PathRemoveFileSpecW(buf);
    return buf;
}
static std::wstring Combine(const std::wstring& a, const std::wstring& b) {
    wchar_t out[MAX_PATH]{};
    wcscpy_s(out, a.c_str());
    PathAppendW(out, b.c_str());
    return out;
}
static bool FileExists(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return (a != INVALID_FILE_ATTRIBUTES) && !(a & FILE_ATTRIBUTE_DIRECTORY);
}
static void ShellRun(const std::wstring& exe, const std::wstring& args, const std::wstring& cwd = L"") {
    SHELLEXECUTEINFOW sei{ sizeof(sei) };
    sei.fMask = SEE_MASK_NOASYNC | SEE_MASK_FLAG_NO_UI;
    sei.lpFile = exe.c_str();
    sei.lpParameters = args.empty() ? nullptr : args.c_str();
    sei.lpDirectory = cwd.empty() ? nullptr : cwd.c_str();
    sei.nShow = SW_SHOWNORMAL;
    ShellExecuteExW(&sei);
}
static std::wstring Find7zTool(const std::wstring& name) {
    auto here = GetModuleDir(g_hMod);
    auto p = Combine(here, name);
    if (FileExists(p)) return p;
    return name; // PATH fallback
}
static std::wstring GetItemPath(IShellItem* it) {
    LPWSTR s = nullptr; std::wstring out;
    if (SUCCEEDED(it->GetDisplayName(SIGDN_FILESYSPATH, &s)) && s) { out = s; CoTaskMemFree(s); }
    return out;
}
static bool IsDirectoryPath(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return (a != INVALID_FILE_ATTRIBUTES) && (a & FILE_ATTRIBUTE_DIRECTORY);
}
static bool IsArchiveExt(const std::wstring& ext) {
    static const wchar_t* exts[] = {
        L".7z",L".zip",L".rar",L".tar",L".gz",L".xz",L".bz2",L".cab",L".wim",L".lzma",L".zst",L".arj"
    };
    for (auto e : exts) if (_wcsicmp(ext.c_str(), e) == 0) return true;
    return false;
}
static std::wstring BaseName(const std::wstring& path) {
    std::filesystem::path p(path);
    return IsDirectoryPath(path) ? p.filename().wstring() : p.stem().wstring();
}
static std::wstring DefaultArchiveName(const std::vector<std::wstring>& paths, const wchar_t* ext) {
    if (paths.empty()) return L"Archive" + std::wstring(ext);
    if (paths.size() == 1) return BaseName(paths[0]) + ext;

    // Multi-selection → use parent directory name if common
    std::filesystem::path parent = std::filesystem::path(paths[0]).parent_path();
    bool allSameParent = true;
    for (auto& p : paths) {
        if (std::filesystem::path(p).parent_path() != parent) { allSameParent = false; break; }
    }
    if (allSameParent && !parent.filename().empty())
        return parent.filename().wstring() + ext;

    // Fallback
    return std::wstring(L"Archive") + ext;
}

// ---------- Command IDs ----------
enum class CommandID {
    None,
    Open, Test, ExtractFiles, ExtractHere, ExtractTo,
    AddToArchive, AddTo7z, AddToZip,
    EmailArchive, Email7z, EmailZip,
    CRCMenu, CRC32, CRC64, SHA1, SHA256
};

// ---------- IEnumExplorerCommand ----------
struct CommandEnum : IEnumExplorerCommand {
    LONG m_ref{ 1 };
    std::vector<IExplorerCommand*> items;
    size_t idx{ 0 };

    CommandEnum(const std::vector<IExplorerCommand*>& v) : items(v) {
        for (auto* i : items) if (i) i->AddRef();
    }
    ~CommandEnum() {
        for (auto* i : items) if (i) i->Release();
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IEnumExplorerCommand) { *ppv = this; AddRef(); return S_OK; }
        *ppv = nullptr; return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG c = InterlockedDecrement(&m_ref); if (!c) delete this; return c;
    }

    // IEnumExplorerCommand
    IFACEMETHODIMP Next(ULONG celt, IExplorerCommand** rgelt, ULONG* pceltFetched) override {
        ULONG f = 0;
        while (f < celt && idx < items.size()) {
            rgelt[f] = items[idx++]; rgelt[f]->AddRef(); ++f;
        }
        if (pceltFetched) *pceltFetched = f;
        return f == celt ? S_OK : S_FALSE;
    }
    IFACEMETHODIMP Skip(ULONG celt) override { idx = std::min(idx + size_t(celt), items.size()); return S_OK; }
    IFACEMETHODIMP Reset() override { idx = 0; return S_OK; }
    IFACEMETHODIMP Clone(IEnumExplorerCommand** pp) override {
        if (!pp) return E_POINTER;
        auto* e = new CommandEnum(items);
        e->idx = idx;
        *pp = e;
        return S_OK;
    }
};

// ---------- Base command ----------
struct ExplorerCommandBase : IExplorerCommand {
    LONG m_ref{ 1 };
    CommandID m_id{};
    std::wstring m_title;

    ExplorerCommandBase(CommandID id, std::wstring title) : m_id(id), m_title(std::move(title)) { InterlockedIncrement(&g_ObjCount); }
    virtual ~ExplorerCommandBase() { InterlockedDecrement(&g_ObjCount); }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IExplorerCommand) { *ppv = this; AddRef(); return S_OK; }
        *ppv = nullptr; return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
    IFACEMETHODIMP_(ULONG) Release() override { ULONG c = InterlockedDecrement(&m_ref); if (!c) delete this; return c; }

    // IExplorerCommandBase(ie.submenu)
    IFACEMETHODIMP GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName) override {
    if (!ppszName) return E_POINTER;

    std::vector<std::wstring> paths;
    CollectPaths(psiItemArray, paths);

    if (m_id == CommandID::AddTo7z || m_id == CommandID::AddToZip ||
        m_id == CommandID::Email7z || m_id == CommandID::EmailZip) {

        std::wstring ext = (m_id == CommandID::AddToZip || m_id == CommandID::EmailZip) ? L".zip" : L".7z";
        std::wstring base = DefaultArchiveName(paths, ext.c_str());

        std::wstring text;
        if (m_id == CommandID::AddTo7z || m_id == CommandID::AddToZip)
            text = L"Add to \"" + base + L"\"";
        else
            text = L"Compress to \"" + base + L"\" and email";

        return SHStrDupW(text.c_str(), ppszName);
    }

    if (m_id == CommandID::ExtractTo) {
        if (!paths.empty()) {
            std::wstring folder = BaseName(paths[0]);
            std::wstring text = L"Extract to \"" + folder + L"\\\"";
            return SHStrDupW(text.c_str(), ppszName);
        }
        return SHStrDupW(L"Extract to \\<Folder>\\", ppszName);
    }

    // fallback to static title
    return SHStrDupW(m_title.c_str(), ppszName);
}
    IFACEMETHODIMP GetIcon(IShellItemArray*, LPWSTR* ppszIcon) override { *ppszIcon = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetToolTip(IShellItemArray*, LPWSTR* ppszInfotip) override { *ppszInfotip = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetCanonicalName(GUID* pguidCommandName) override { *pguidCommandName = GUID_NULL; return E_NOTIMPL; }

    static void CollectPaths(IShellItemArray* arr, std::vector<std::wstring>& out) {
        out.clear(); if (!arr) return;
        DWORD c = 0; if (FAILED(arr->GetCount(&c))) return;
        for (DWORD i = 0; i < c; ++i) {
            IShellItem* it = nullptr;
            if (SUCCEEDED(arr->GetItemAt(i, &it)) && it) {
                auto p = GetItemPath(it);
                it->Release();
                if (!p.empty()) out.push_back(std::move(p));
            }
        }
    }

    IFACEMETHODIMP GetState(IShellItemArray* psiItemArray, BOOL, EXPCMDSTATE* pState) override {
        *pState = ECS_HIDDEN;
        std::vector<std::wstring> paths;
        CollectPaths(psiItemArray, paths);
        if (paths.empty()) return S_OK;

        bool allArchives = true;
        for (auto& p : paths) {
            auto ext = std::filesystem::path(p).extension().wstring();
            if (!IsArchiveExt(ext)) { allArchives = false; break; }
        }

        switch (m_id) {
        case CommandID::Open:
            if (paths.size() == 1 && allArchives) *pState = ECS_ENABLED;
            break;
        case CommandID::Test:
        case CommandID::ExtractFiles:
        case CommandID::ExtractHere:
        case CommandID::ExtractTo:
            if (allArchives) *pState = ECS_ENABLED;
            break;
        default:
            *pState = ECS_ENABLED; // Add/Email/CRC always available
            break;
        }
        return S_OK;
    }

    IFACEMETHODIMP Invoke(IShellItemArray* psiItemArray, IBindCtx*) override {
        std::vector<std::wstring> paths;
        CollectPaths(psiItemArray, paths);
        if (paths.empty()) return S_OK;

        const auto sevenZG = Find7zTool(L"7zG.exe");
        const auto sevenZ  = Find7zTool(L"7z.exe");
        const auto sevenFM = Find7zTool(L"7zFM.exe");

        auto quoteJoin = [](const std::vector<std::wstring>& v) {
            std::wstring s; for (auto& p : v) { s += L"\""; s += p; s += L"\" "; } return s;
        };

        switch (m_id) {
        case CommandID::Open:
            ShellRun(sevenFM, L"\"" + paths[0] + L"\"");
            break;

        case CommandID::Test:
            ShellRun(sevenZG, L"t " + quoteJoin(paths));
            break;

        case CommandID::ExtractFiles:
            // GUI extract dialog
            ShellRun(sevenZG, L"x " + quoteJoin(paths));
            break;

        case CommandID::ExtractHere:
            if (paths.size() == 1) {
                // Classic single-archive behavior
                ShellRun(sevenZG, L"x -y \"" + paths[0] + L"\"");
            } else {
                // SMART multi-archive behavior: each into its own folder
                for (auto& p : paths) {
                    std::wstring folder = BaseName(p);
                    std::wstring args = L"x -y -o\"" + folder + L"\\\" \"" + p + L"\"";
                    ShellRun(sevenZG, args);
                }
            }
            break;

        case CommandID::ExtractTo:
            // Classic: always into <ArchiveName>\ (multi-select creates per-archive dirs)
            for (auto& p : paths) {
                std::wstring folder = BaseName(p);
                std::wstring args = L"x -y -o\"" + folder + L"\\\" \"" + p + L"\"";
                ShellRun(sevenZG, args);
            }
            break;

        case CommandID::AddToArchive:
            ShellRun(sevenZG, L"a " + quoteJoin(paths));
            break;

        case CommandID::AddTo7z: {
            std::wstring out = DefaultArchiveName(paths, L".7z");
            ShellRun(sevenZG, L"a \"" + out + L"\" " + quoteJoin(paths));
            break;
        }

        case CommandID::AddToZip: {
            std::wstring out = DefaultArchiveName(paths, L".zip");
            ShellRun(sevenZG, L"a -tzip \"" + out + L"\" " + quoteJoin(paths));
            break;
        }

        case CommandID::EmailArchive:
            ShellRun(sevenZG, L"a " + quoteJoin(paths));
            break;

        case CommandID::Email7z: {
            std::wstring out = DefaultArchiveName(paths, L".7z");
            ShellRun(sevenZG, L"a \"" + out + L"\" " + quoteJoin(paths));
            break;
        }

        case CommandID::EmailZip: {
            std::wstring out = DefaultArchiveName(paths, L".zip");
            ShellRun(sevenZG, L"a -tzip \"" + out + L"\" " + quoteJoin(paths));
            break;
        }

        case CommandID::CRC32:
            ShellRun(sevenZ, L"h -scrcCRC32 " + quoteJoin(paths));
            break;
        case CommandID::CRC64:
            ShellRun(sevenZ, L"h -scrcCRC64 " + quoteJoin(paths));
            break;
        case CommandID::SHA1:
            ShellRun(sevenZ, L"h -scrcSHA1 " + quoteJoin(paths));
            break;
        case CommandID::SHA256:
            ShellRun(sevenZ, L"h -scrcSHA256 " + quoteJoin(paths));
            break;

        default:
            break;
        }
        return S_OK;
    }

    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* pFlags) override {
        *pFlags = (m_id == CommandID::CRCMenu) ? ECF_HASSUBCOMMANDS : ECF_DEFAULT;
        return S_OK;
    }
    virtual IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand**) { return E_NOTIMPL; }
};

// CRC submenu parent
struct CRCMenuParent : ExplorerCommandBase {
    CRCMenuParent() : ExplorerCommandBase(CommandID::CRCMenu, L"CRC SHA") {}
    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** ppEnum) override {
        std::vector<IExplorerCommand*> v;
        v.push_back(new ExplorerCommandBase(CommandID::CRC32,  L"CRC-32"));
        v.push_back(new ExplorerCommandBase(CommandID::CRC64,  L"CRC-64"));
        v.push_back(new ExplorerCommandBase(CommandID::SHA1,   L"SHA-1"));
        v.push_back(new ExplorerCommandBase(CommandID::SHA256, L"SHA-256"));
        *ppEnum = new CommandEnum(v);
        (*ppEnum)->AddRef();
        return S_OK;
    }
};

// Root flyout
struct ExplorerCommandRoot : IExplorerCommand, IEnumExplorerCommand {
    LONG m_ref{1}; std::vector<IExplorerCommand*> subs; size_t idx{0};

    ExplorerCommandRoot() {
        InterlockedIncrement(&g_ObjCount);
        subs.push_back(new ExplorerCommandBase(CommandID::Open,         L"Open archive"));
        subs.push_back(new ExplorerCommandBase(CommandID::ExtractFiles, L"Extract files..."));
        subs.push_back(new ExplorerCommandBase(CommandID::ExtractHere,  L"Extract Here"));
        subs.push_back(new ExplorerCommandBase(CommandID::ExtractTo,    L"Extract to \\<Folder>\\"));
        subs.push_back(new ExplorerCommandBase(CommandID::Test,         L"Test archive"));
        subs.push_back(new ExplorerCommandBase(CommandID::AddToArchive, L"Add to archive..."));
        subs.push_back(new ExplorerCommandBase(CommandID::AddTo7z,      L"Add to \"<Name>.7z\""));
        subs.push_back(new ExplorerCommandBase(CommandID::AddToZip,     L"Add to \"<Name>.zip\""));
        subs.push_back(new ExplorerCommandBase(CommandID::EmailArchive, L"Compress and email..."));
        subs.push_back(new ExplorerCommandBase(CommandID::Email7z,      L"Compress to \"<Name>.7z\" and email"));
        subs.push_back(new ExplorerCommandBase(CommandID::EmailZip,     L"Compress to \"<Name>.zip\" and email"));
        subs.push_back(new CRCMenuParent());
    }
    ~ExplorerCommandRoot(){for(auto*c:subs)if(c)c->Release();InterlockedDecrement(&g_ObjCount);}

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid,void**ppv)override{
        if(!ppv) return E_POINTER;
        if(riid==IID_IUnknown||riid==IID_IExplorerCommand){*ppv=(IExplorerCommand*)this;AddRef();return S_OK;}
        if(riid==IID_IEnumExplorerCommand){*ppv=(IEnumExplorerCommand*)this;AddRef();return S_OK;}
        *ppv=nullptr;return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG)AddRef()override{return InterlockedIncrement(&m_ref);}
    IFACEMETHODIMP_(ULONG)Release()override{ULONG c=InterlockedDecrement(&m_ref);if(!c)delete this;return c;}

    // IExplorerCommand root
    IFACEMETHODIMP GetTitle(IShellItemArray*,LPWSTR*pp)override{return SHStrDupW(L"7-Zip",pp);}
    IFACEMETHODIMP GetIcon(IShellItemArray*, LPWSTR* ppszIcon) override {
    if (!ppszIcon) return E_POINTER;
    *ppszIcon = nullptr;
    std::wstring iconPath = L"C:\\Program Files\\7-Zip\\7zFM.exe,0";
    return SHStrDupW(iconPath.c_str(), ppszIcon);
}

    IFACEMETHODIMP GetToolTip(IShellItemArray*,LPWSTR*pp)override{*pp=nullptr;return E_NOTIMPL;}
    IFACEMETHODIMP GetCanonicalName(GUID*pg)override{*pg=CLSID_SevenZipExplorer;return S_OK;}
    IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* st) override {
    *st = ECS_ENABLED; // the root command is always shown
    return S_OK;
}
    IFACEMETHODIMP Invoke(IShellItemArray*,IBindCtx*)override{return S_OK;}
    IFACEMETHODIMP GetFlags(EXPCMDFLAGS*f)override{*f=ECF_HASSUBCOMMANDS;return S_OK;}
    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand**pp)override{*pp=new CommandEnum(subs);(*pp)->AddRef();return S_OK;}

    // IEnumExplorerCommand (unused here)
    IFACEMETHODIMP Next(ULONG,IExplorerCommand**,ULONG*)override{return E_NOTIMPL;}
    IFACEMETHODIMP Skip(ULONG)override{return E_NOTIMPL;}
    IFACEMETHODIMP Reset()override{return E_NOTIMPL;}
    IFACEMETHODIMP Clone(IEnumExplorerCommand**)override{return E_NOTIMPL;}
};

// ---------- ClassFactory ----------
struct ClassFactory : IClassFactory {
    LONG m_ref{1};
    IFACEMETHODIMP QueryInterface(REFIID riid,void**ppv)override{
        if(!ppv) return E_POINTER;
        if(riid==IID_IUnknown||riid==IID_IClassFactory){*ppv=this;AddRef();return S_OK;}
        *ppv=nullptr;return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG)AddRef()override{return InterlockedIncrement(&m_ref);}
    IFACEMETHODIMP_(ULONG)Release()override{ULONG c=InterlockedDecrement(&m_ref);if(!c)delete this;return c;}
    IFACEMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override {
        if (pUnkOuter) return CLASS_E_NOAGGREGATION;
        auto* obj = new ExplorerCommandRoot();
        HRESULT hr = obj->QueryInterface(riid, ppv);
        obj->Release();
        return hr;
    }
    IFACEMETHODIMP LockServer(BOOL fLock) override {
        if (fLock) InterlockedIncrement(&g_LockCount); else InterlockedDecrement(&g_LockCount);
        return S_OK;
    }
};

// ---------- Exports ----------
extern "C" BOOL APIENTRY DllMain(HMODULE hModule, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        g_hMod = hModule;
        DisableThreadLibraryCalls(hModule);
    }
    return TRUE;
}

// extern "C" __declspec(dllexport) HRESULT __stdcall DllCanUnloadNow(void) {
//     return (g_ObjCount == 0 && g_LockCount == 0) ? S_OK : S_FALSE;
// }

// extern "C" __declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {
//     if (rclsid != CLSID_SevenZipExplorer) return CLASS_E_CLASSNOTAVAILABLE;
//     auto* cf = new ClassFactory();
//     HRESULT hr = cf->QueryInterface(riid, ppv);
//     cf->Release();
//     return hr;
// }
HRESULT __stdcall DllCanUnloadNow(void) {
    return (g_ObjCount == 0 && g_LockCount == 0) ? S_OK : S_FALSE;
}

HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {
    if (rclsid != CLSID_SevenZipExplorer) return CLASS_E_CLASSNOTAVAILABLE;
    auto* cf = new ClassFactory();
    HRESULT hr = cf->QueryInterface(riid, ppv);
    cf->Release();
    return hr;
}
