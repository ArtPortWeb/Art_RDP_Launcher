/*
  RDP Launcher
  ------------
  Scans Documents\RDP_Connect for *.rdp files, shows a list,
  copies selected file to Default.rdp and launches mstsc.exe.

  Build (MSVC Developer Command Prompt):
    rc resources.rc
    cl /EHsc /O1 art_rdp_launcher.cpp resources.res /FeArt_RDP_Launcher.exe /link /subsystem:windows

  Build (MinGW):
    windres resources.rc -O coff -o resources.res
    g++ -o Art_RDP_Launcher.exe art_rdp_launcher.cpp resources.res -lshell32 -lcomctl32 -luser32 -lgdi32 -mwindows -std=c++17
*/

#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <algorithm>
#include "version.h"

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

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

// ---------------------------------------------------------------------------
// Globals
// ---------------------------------------------------------------------------
static HWND g_hList   = NULL;
static HWND g_hStatus = NULL;
static HICON g_hAppIcon = NULL;

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
// Connect: copy rdp -> Default.rdp, launch mstsc
// ---------------------------------------------------------------------------
static void Connect(HWND hWnd, int idx)
{
    if (idx < 0 || idx >= (int)g_entries.size()) return;

    const std::wstring& src = g_entries[idx].fullPath;

    BOOL ok = CopyFileW(src.c_str(), g_defaultRdp.c_str(), FALSE);
    if (!ok) {
        DWORD err = GetLastError();
        wchar_t msg[1024];
        wsprintfW(msg,
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

        PopulateList(g_hList);

        // Status: show folder path + count
        g_hStatus = CreateWindowExW(0, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            16, 252, 340, 18,
            hWnd, (HMENU)IDC_STATUS, NULL, NULL);

        std::wstring status = g_rdpFolder + L"  (" +
            std::to_wstring(g_entries.size()) + L" files)";
        SetWindowTextW(g_hStatus, status.c_str());

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
        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        EnumChildWindows(hWnd, SetFontProc, (LPARAM)hFont);

        break;
    }

    case WM_COMMAND:
    {
        int ctrl  = LOWORD(wParam);
        int notif = HIWORD(wParam);

        if (ctrl == IDC_BTN_CANCEL) {
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

    case WM_KEYDOWN:
        if (wParam == VK_ESCAPE) DestroyWindow(hWnd);
        break;

    case WM_DESTROY:
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
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    // Resolve Documents path
    wchar_t docs[MAX_PATH] = {};
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_PERSONAL, NULL, SHGFP_TYPE_CURRENT, docs))) {
        MessageBoxW(NULL, L"Cannot resolve Documents folder.", WND_TITLE, MB_ICONERROR);
        return 1;
    }

    g_rdpFolder  = std::wstring(docs) + L"\\" + RDP_SUBFOLDER;
    g_defaultRdp = std::wstring(docs) + L"\\Default.rdp";

    // Create RDP_Connect folder if it does not exist
    CreateDirectoryW(g_rdpFolder.c_str(), NULL);

    ScanFolder();

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
        WS_EX_DLGMODALFRAME,
        WND_CLASS, WND_TITLE,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 392, 358,
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

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return 0;
}