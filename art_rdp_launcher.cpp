/*
  RDP Launcher
  ------------
  Scans Documents\RDP_Connect for *.rdp files, shows a list,
  copies selected file to Default.rdp and launches mstsc.exe.

  Build (MSVC Developer Command Prompt):
    rc resources.rc
    cl /MT /EHsc /O1 /GS art_rdp_launcher.cpp resources.res /FeArt_RDP_Launcher.exe /link /subsystem:windows /DYNAMICBASE /HIGHENTROPYVA /NXCOMPAT /GUARD:CF

  Build (MinGW):
    windres resources.rc -O coff -o resources.res
    g++ -o Art_RDP_Launcher.exe art_rdp_launcher.cpp resources.res -lshell32 -lcomctl32 -luser32 -lgdi32 -mwindows -std=c++17 -Wl,--dynamicbase,--nxcompat,--high-entropy-va
*/

#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <commctrl.h>
#include <cwchar>
#include <string>
#include <vector>
#include <algorithm>
#include "version.h"

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")

// ---------------------------------------------------------------------------
// IDs
// ---------------------------------------------------------------------------
#define IDI_APP_ICON    1
#define IDC_LIST        101
#define IDC_BTN_CONNECT 102
#define IDC_BTN_CANCEL  103
#define IDC_LABEL       104
#define IDC_STATUS      105

static const wchar_t* RDP_SUBFOLDER = L"RDP_Connect";
static const wchar_t* WND_CLASS     = L"RdpLauncherWnd";
static const wchar_t* WND_TITLE     = L"ART RDP Launcher v" VERSION_STRING_W;

static const int WND_WIDTH  = 392;
static const int WND_HEIGHT = 358;

// ---------------------------------------------------------------------------
// Globals
// ---------------------------------------------------------------------------
static HWND g_hList   = NULL;
static HWND g_hStatus = NULL;
static HICON g_hAppIcon = NULL;
static HFONT g_hFont = NULL;

static std::wstring g_rdpFolder;
static std::wstring g_defaultRdp;

struct RdpEntry {
    std::wstring filename;
    std::wstring fullPath;
};
static std::vector<RdpEntry> g_entries;

// ---------------------------------------------------------------------------
// Scan folder for *.rdp
// ---------------------------------------------------------------------------
static void ScanFolder()
{
    g_entries.clear();

    std::wstring pattern = g_rdpFolder + L"\\*.rdp";
    WIN32_FIND_DATAW fd = {};
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE) return;

    do {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
        RdpEntry e;
        e.filename = fd.cFileName;
        e.fullPath = g_rdpFolder + L"\\" + fd.cFileName;
        g_entries.push_back(e);
    } while (FindNextFileW(hFind, &fd));

    FindClose(hFind);

    std::sort(g_entries.begin(), g_entries.end(),
        [](const RdpEntry& a, const RdpEntry& b) {
            return _wcsicmp(a.filename.c_str(), b.filename.c_str()) < 0;
        });
}

// ---------------------------------------------------------------------------
// Fill listbox
// ---------------------------------------------------------------------------
static void PopulateList(HWND hList)
{
    SendMessageW(hList, LB_RESETCONTENT, 0, 0);

    for (auto& e : g_entries) {
        std::wstring label = e.filename;
        auto dot = label.rfind(L'.');
        if (dot != std::wstring::npos) label.resize(dot);
        SendMessageW(hList, LB_ADDSTRING, 0, (LPARAM)label.c_str());
    }

    if (!g_entries.empty())
        SendMessageW(hList, LB_SETCURSEL, 0, 0);
}

// ---------------------------------------------------------------------------
// Import RDP file
// ---------------------------------------------------------------------------
static bool ImportRdpFile(HWND hWnd, const std::wstring& srcPath)
{
    // 1. Извлекаем имя файла
    const wchar_t* fname = wcsrchr(srcPath.c_str(), L'\\');
    if (!fname) fname = wcsrchr(srcPath.c_str(), L'/');
    fname = fname ? fname + 1 : srcPath.c_str();

    // 2. Проверяем расширение .rdp
    const wchar_t* ext = wcsrchr(fname, L'.');
    if (!ext || _wcsicmp(ext, L".rdp") != 0) return false;

    std::wstring dest = g_rdpFolder + L"\\" + fname;

    // 3. Детекция самокопирования: канонизируем оба пути и сравниваем.
    //    Динамические буферы — консистентно с B5 (long paths >260 символов).
    DWORD lenSrc = GetFullPathNameW(srcPath.c_str(), 0, NULL, NULL);
    DWORD lenDst = GetFullPathNameW(dest.c_str(), 0, NULL, NULL);
    std::wstring canonSrc(lenSrc, L'\0'), canonDst(lenDst, L'\0');
    GetFullPathNameW(srcPath.c_str(), lenSrc, &canonSrc[0], NULL);
    GetFullPathNameW(dest.c_str(), lenDst, &canonDst[0], NULL);
    if (CompareStringOrdinal(canonSrc.c_str(), -1, canonDst.c_str(), -1, TRUE) == CSTR_EQUAL) {
        return false;  // файл уже в целевой папке — пропускаем молча
    }

    // 4. Если файл с таким именем уже существует — спрашиваем
    if (GetFileAttributesW(dest.c_str()) != INVALID_FILE_ATTRIBUTES) {
        wchar_t msg[512];
        swprintf_s(msg, L"Файл \"%s\" уже существует в RDP_Connect.\nПерезаписать?", fname);
        if (MessageBoxW(hWnd, msg, L"RDP Launcher - Import",
                        MB_YESNO | MB_ICONQUESTION) != IDYES)
            return false;
    }

    // 5. Копирование
    if (!CopyFileW(srcPath.c_str(), dest.c_str(), FALSE)) {
        DWORD err = GetLastError();
        wchar_t msg[512];
        swprintf_s(msg, L"Ошибка копирования \"%s\" (WinError %lu).", fname, err);
        MessageBoxW(hWnd, msg, L"RDP Launcher - Error", MB_ICONERROR);
        return false;
    }
    return true;
}

// ---------------------------------------------------------------------------
// Refresh list
// ---------------------------------------------------------------------------
static void RefreshList()
{
    ScanFolder();
    if (g_hList)   PopulateList(g_hList);
    if (g_hStatus) {
        std::wstring status = g_rdpFolder + L"  (" +
            std::to_wstring(g_entries.size()) + L" files)";
        SetWindowTextW(g_hStatus, status.c_str());
    }
}

// ---------------------------------------------------------------------------
// Connect: copy rdp -> Default.rdp, launch mstsc
// ---------------------------------------------------------------------------
static void Connect(HWND hWnd, int idx)
{
    if (idx < 0 || idx >= (int)g_entries.size()) return;

    const std::wstring& src = g_entries[idx].fullPath;

    // CopyFileW overwrites destination. If it fails with access denied
    // (e.g. Default.rdp was left hidden/read-only by mstsc), clear attributes and retry.
    // CopyFileW returns 0 on failure. GetLastError() is valid only if !copied.
    bool copied = CopyFileW(src.c_str(), g_defaultRdp.c_str(), FALSE) != 0;
    if (!copied && GetLastError() == ERROR_ACCESS_DENIED) {
        SetFileAttributesW(g_defaultRdp.c_str(), FILE_ATTRIBUTE_NORMAL);
        copied = CopyFileW(src.c_str(), g_defaultRdp.c_str(), FALSE) != 0;
    }
    if (!copied) {
        DWORD err = GetLastError();
        wchar_t msg[1024];
        swprintf_s(msg,
            L"Failed to copy file (WinError %lu).\n\nFrom:\n%s\n\nTo:\n%s",
            err, src.c_str(), g_defaultRdp.c_str());
        MessageBoxW(hWnd, msg, L"RDP Launcher - Error", MB_ICONERROR);
        return;
    }

    // Launch mstsc - it reads Default.rdp automatically
    SHELLEXECUTEINFOW sei = { sizeof(sei) };
    sei.fMask = SEE_MASK_DEFAULT;
    sei.lpVerb = L"open";
    sei.lpFile = L"mstsc.exe";
    sei.nShow = SW_SHOWNORMAL;

    if (!ShellExecuteExW(&sei)) {
        MessageBoxW(hWnd, L"Failed to launch mstsc.exe", L"RDP Launcher - Error", MB_ICONERROR);
        return;
    }

    // Hide window immediately to avoid feeling like the app is lagging
    ShowWindow(hWnd, SW_HIDE);

    DestroyWindow(hWnd);
}

// ---------------------------------------------------------------------------
// Apply font callback
// ---------------------------------------------------------------------------
static BOOL CALLBACK SetFontProc(HWND hw, LPARAM lp)
{
    SendMessageW(hw, WM_SETFONT, (WPARAM)lp, TRUE);
    return TRUE;
}

// ---------------------------------------------------------------------------
// Window procedure
// ---------------------------------------------------------------------------
static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        // Label
        CreateWindowExW(0, L"STATIC", L"Select connection:",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            16, 12, 340, 20,
            hWnd, (HMENU)IDC_LABEL, NULL, NULL);

        // Listbox
        g_hList = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT,
            16, 36, 340, 210,
            hWnd, (HMENU)IDC_LIST, NULL, NULL);

        // Status: show folder path + count
        g_hStatus = CreateWindowExW(0, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            16, 252, 340, 18,
            hWnd, (HMENU)IDC_STATUS, NULL, NULL);

        RefreshList();

        // Connect button
        CreateWindowExW(0, L"BUTTON", L"Connect",
            WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
            16, 276, 160, 32,
            hWnd, (HMENU)IDC_BTN_CONNECT, NULL, NULL);

        // Cancel button
        CreateWindowExW(0, L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            196, 276, 160, 32,
            hWnd, (HMENU)IDC_BTN_CANCEL, NULL, NULL);

        // Apply GUI font to all controls
        NONCLIENTMETRICSW ncm = { sizeof(ncm) };
        SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
        g_hFont = CreateFontIndirectW(&ncm.lfMessageFont);
        EnumChildWindows(hWnd, SetFontProc, (LPARAM)g_hFont);

        break;
    }

    case WM_COPYDATA:
    {
        COPYDATASTRUCT* cds = (COPYDATASTRUCT*)lParam;
        if (!cds || cds->dwData != 0x52445046) // 'RDPF' magic
            return FALSE;

        // Валидация: минимальный размер, кратность sizeof(wchar_t), нуль-терминатор
        if (cds->cbData < sizeof(wchar_t))
            return FALSE;
        if (cds->cbData % sizeof(wchar_t) != 0)
            return FALSE;
        const wchar_t* raw = (const wchar_t*)cds->lpData;
        size_t maxChars = cds->cbData / sizeof(wchar_t);
        if (raw[maxChars - 1] != L'\0')
            return FALSE;

        std::wstring path(raw);
        if (ImportRdpFile(hWnd, path))
            RefreshList();

        return TRUE;
    }

    case WM_DROPFILES:
    {
        HDROP hDrop = (HDROP)wParam;
        UINT count = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
        int imported = 0;
        for (UINT i = 0; i < count; ++i) {
            UINT len = DragQueryFileW(hDrop, i, NULL, 0);
            std::wstring path(len, L'\0');
            DragQueryFileW(hDrop, i, &path[0], len + 1);
            if (ImportRdpFile(hWnd, path))
                ++imported;
        }
        DragFinish(hDrop);
        if (imported > 0)
            RefreshList();
        break;
    }

    case WM_COMMAND:
    {
        int ctrl  = LOWORD(wParam);
        int notif = HIWORD(wParam);

        if (ctrl == IDC_BTN_CANCEL || ctrl == IDCANCEL) {
            DestroyWindow(hWnd);
        }
        else if (ctrl == IDC_BTN_CONNECT ||
                 (ctrl == IDC_LIST && notif == LBN_DBLCLK))
        {
            int idx = (int)SendMessageW(g_hList, LB_GETCURSEL, 0, 0);
            if (idx != LB_ERR && !g_entries.empty())
                Connect(hWnd, idx);
        }
        break;
    }

    case WM_DESTROY:
        if (g_hFont) {
            DeleteObject(g_hFont);
            g_hFont = NULL;
        }
        PostQuitMessage(0);
        break;
    }

    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ---------------------------------------------------------------------------
// wWinMain
// ---------------------------------------------------------------------------
int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int)
{
    HANDLE hMutex = CreateMutexW(NULL, FALSE, L"ArtSoft_RdpLauncher_SingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        // Find and activate existing window
        HWND hExisting = FindWindowW(WND_CLASS, NULL);
        if (hExisting) {
            SetForegroundWindow(hExisting);
            if (IsIconic(hExisting)) ShowWindow(hExisting, SW_RESTORE);

            int argc = 0;
            LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
            if (argv) {
                for (int i = 1; i < argc; ++i) {
                    COPYDATASTRUCT cds = {};
                    cds.dwData = 0x52445046; // 'RDPF' magic
                    cds.cbData = (DWORD)((wcslen(argv[i]) + 1) * sizeof(wchar_t));
                    cds.lpData = argv[i];
                    SendMessageW(hExisting, WM_COPYDATA, 0, (LPARAM)&cds);
                }
                LocalFree(argv);
            }
        }
        CloseHandle(hMutex);
        return 0;
    }

    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    // Resolve Documents path
    PWSTR docsPath = nullptr;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_Documents, 0, NULL, &docsPath);
    if (FAILED(hr)) {
        MessageBoxW(NULL, L"Cannot resolve Documents folder.", WND_TITLE, MB_ICONERROR);
        CloseHandle(hMutex);
        return 1;
    }

    g_rdpFolder  = std::wstring(docsPath) + L"\\" + RDP_SUBFOLDER;
    g_defaultRdp = std::wstring(docsPath) + L"\\Default.rdp";
    CoTaskMemFree(docsPath);

    // Create RDP_Connect folder if it does not exist
    CreateDirectoryW(g_rdpFolder.c_str(), NULL);

    g_hAppIcon = LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON));
    if (!g_hAppIcon) {
        // Fallback to system icon if resource not found
        g_hAppIcon = LoadIconW(NULL, IDI_APPLICATION);
    }

    // Register window class
    WNDCLASSEXW wc = {};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = WND_CLASS;
    wc.hIcon         = g_hAppIcon;
    wc.hIconSm       = g_hAppIcon;
    if (!RegisterClassExW(&wc)) return 1;

    // Create fixed-size dialog-like window
    HWND hWnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_ACCEPTFILES,
        WND_CLASS, WND_TITLE,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, WND_WIDTH, WND_HEIGHT,
        NULL, NULL, hInst, NULL);

    if (!hWnd) return 1;

    // Center on screen
    RECT rc = {}, rcScr = {};
    GetWindowRect(hWnd, &rc);
    GetWindowRect(GetDesktopWindow(), &rcScr);
    SetWindowPos(hWnd, NULL,
        (rcScr.right  - (rc.right  - rc.left)) / 2,
        (rcScr.bottom - (rc.bottom - rc.top))  / 2,
        0, 0, SWP_NOSIZE | SWP_NOZORDER);

    ShowWindow(hWnd, SW_SHOW);
    UpdateWindow(hWnd);

    // Обработка аргументов (drag на EXE)
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv) {
        int imported = 0;
        for (int i = 1; i < argc; ++i) {
            if (ImportRdpFile(hWnd, argv[i]))
                ++imported;
        }
        LocalFree(argv);
        if (imported > 0)
            RefreshList();
    }

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        if (!IsDialogMessageW(hWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    
    CloseHandle(hMutex);
    return 0;
}