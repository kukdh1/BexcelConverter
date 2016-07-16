#include "Main.h"

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
DWORD WINAPI ConvertThread(LPVOID);
BOOL FileOpenDialog(WCHAR **pszPath);
BOOL BrowseFolder(HWND hParent, LPCWSTR szTitle, LPCWSTR szStartPath, WCHAR *szFolder, DWORD dwBufferLength);
BOOL ConvertString(std::wstring &in, std::string &out);

LPCWSTR lpszClass = L"Bexel converter";

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdParam, int nCmdShow) {
  HWND hWnd;
  WNDCLASS wndclass;
  MSG msg;
  INITCOMMONCONTROLSEX iccex;

  iccex.dwSize = sizeof(INITCOMMONCONTROLSEX);
  iccex.dwICC = ICC_WIN95_CLASSES | ICC_PROGRESS_CLASS;

  if (!InitCommonControlsEx(&iccex)) {
    return -1;
  }

  if (!SUCCEEDED(CoInitializeEx(NULL, COINIT_MULTITHREADED))) {
    return -2;
  }

  wndclass.cbClsExtra = 0;
  wndclass.cbWndExtra = 0;
  wndclass.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
  wndclass.hCursor = LoadCursor(NULL, IDC_ARROW);
  wndclass.hIcon = LoadIcon(NULL, IDI_APPLICATION);
  wndclass.hInstance = hInstance;
  wndclass.lpfnWndProc = WndProc;
  wndclass.lpszClassName = lpszClass;
  wndclass.lpszMenuName = NULL;
  wndclass.style = CS_VREDRAW | CS_HREDRAW;

  RegisterClass(&wndclass);

  RECT rtRect;
  int uiHeight = 0;
  int nStatusbarHeight = GetSystemMetrics(SM_CYMENU) + GetSystemMetrics(SM_CYBORDER) * 2;

  uiHeight = CONTROL_HEIGHT * 4 + CONTROL_MARGIN * 5 + nStatusbarHeight;
  
  SetRect(&rtRect, 0, 0, WINDOW_WIDTH, uiHeight);
  AdjustWindowRect(&rtRect, WS_CAPTION | WS_SYSMENU, FALSE);

  hWnd = CreateWindow(lpszClass, lpszClass, WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, rtRect.right - rtRect.left, rtRect.bottom - rtRect.top, NULL, NULL, hInstance, NULL);
  ShowWindow(hWnd, nCmdShow);

  while (GetMessage(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessage(&msg);
  }

  CoUninitialize();

  return msg.wParam;
}

HWND hEditSource;
HWND hEditDest;
HWND hButtonSource;
HWND hButtonDest;
HWND hProgress;
HWND hButtonStart;
HWND hButtonOption;
HWND hStatusBar;
HFONT hFont;

WCHAR *pszOpenFilePath = NULL;
WCHAR *pszSaveFolderPath = NULL;

LRESULT CALLBACK WndProc(HWND hWnd, UINT iMessage, WPARAM wParam, LPARAM lParam) {
  switch (iMessage) {
    case WM_CREATE:
      hFont = CreateFont(FONT_SIZE, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_ROMAN, FONT_FACE);

      hStatusBar = CreateWindow(STATUSCLASSNAME, NULL, WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hWnd, (HMENU)ID_STATUS_BAR, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      hEditSource = CreateWindowEx(WS_EX_CLIENTEDGE, WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, CONTROL_MARGIN, CONTROL_MARGIN, EDIT_WIDTH, CONTROL_HEIGHT, hWnd, (HMENU)ID_EDIT_SRC, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      hEditDest = CreateWindowEx(WS_EX_CLIENTEDGE, WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, CONTROL_MARGIN, CONTROL_MARGIN * 2 + CONTROL_HEIGHT, EDIT_WIDTH, CONTROL_HEIGHT, hWnd, (HMENU)ID_EDIT_DST, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      hButtonSource = CreateWindow(WC_BUTTON, STRING_OPEN, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, CONTROL_MARGIN * 2 + EDIT_WIDTH, CONTROL_MARGIN, BUTTON_WIDTH, CONTROL_HEIGHT, hWnd, (HMENU)ID_BUTTON_OPEN_SRC, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      hButtonDest = CreateWindow(WC_BUTTON, STRING_SAVE, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, CONTROL_MARGIN * 2 + EDIT_WIDTH, CONTROL_MARGIN * 2 + CONTROL_HEIGHT, BUTTON_WIDTH, CONTROL_HEIGHT, hWnd, (HMENU)ID_BUTTON_OPEN_DST, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      hProgress = CreateWindow(PROGRESS_CLASS, NULL, WS_CHILD | WS_VISIBLE | PBS_SMOOTH, CONTROL_MARGIN, CONTROL_MARGIN * 4 + CONTROL_HEIGHT * 3, PROGRESS_WIDTH, CONTROL_HEIGHT, hWnd, (HMENU)ID_PROGRESS, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      hButtonStart = CreateWindow(WC_BUTTON, STRING_START, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, CONTROL_MARGIN * 2 + EDIT_WIDTH, CONTROL_MARGIN * 3 + CONTROL_HEIGHT * 2, BUTTON_WIDTH, CONTROL_HEIGHT, hWnd, (HMENU)ID_BUTTON_START, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      hButtonOption = CreateWindow(WC_BUTTON, STRING_OPTION, WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, CONTROL_MARGIN, CONTROL_MARGIN * 3 + CONTROL_HEIGHT * 2, EDIT_WIDTH, CONTROL_HEIGHT, hWnd, (HMENU)ID_BUTTON_OPTION, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
      
      SendMessage(hEditSource, WM_SETFONT, (WPARAM)hFont, TRUE);
      SendMessage(hEditDest, WM_SETFONT, (WPARAM)hFont, TRUE);
      SendMessage(hButtonSource, WM_SETFONT, (WPARAM)hFont, TRUE);
      SendMessage(hButtonDest, WM_SETFONT, (WPARAM)hFont, TRUE);
      SendMessage(hButtonStart, WM_SETFONT, (WPARAM)hFont, TRUE);
      SendMessage(hButtonOption, WM_SETFONT, (WPARAM)hFont, TRUE);

      SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)STRING_PROGRESS_IDLE);

      return 0;
    case WM_COMMAND:
      switch (LOWORD(wParam)) {
        case ID_BUTTON_OPEN_SRC:
          {
            if (FileOpenDialog(&pszOpenFilePath)) {
              SendMessage(hEditSource, WM_SETTEXT, NULL, (LPARAM)pszOpenFilePath);
              SAFE_FREE(pszOpenFilePath);
            }
          }

          break;
        case ID_BUTTON_OPEN_DST:
          {
            pszSaveFolderPath = (WCHAR *)calloc(4096, sizeof(WCHAR));

            if (BrowseFolder(hWnd, L"Select folder to save", L"C:\\", pszSaveFolderPath, 4096)) {
              SendMessage(hEditDest, WM_SETTEXT, NULL, (LPARAM)pszSaveFolderPath);
            }

            SAFE_FREE(pszSaveFolderPath);
          }

          break;
        case ID_BUTTON_START:
          {
            BOOL bChecked = SendMessage(hButtonOption, BM_GETCHECK, NULL, NULL) == BST_CHECKED;
            HANDLE hThread = CreateThread(NULL, 0, ConvertThread, (LPVOID)bChecked, NULL, NULL);
            CloseHandle(hThread);
          }

          break;
      }

      return 0;
    case WM_CTLCOLORSTATIC:
      return (LRESULT)GetStockObject(WHITE_BRUSH);
    case WM_DESTROY:
      DeleteObject(hFont);
      SAFE_FREE(pszOpenFilePath);
      SAFE_FREE(pszSaveFolderPath);

      PostQuitMessage(0);

      return 0;
  }

  return DefWindowProc(hWnd, iMessage, wParam, lParam);
}

DWORD WINAPI ConvertThread(LPVOID arg) {
  BOOL bSplit = (BOOL)arg;

  EnableWindow(hEditSource, FALSE);
  EnableWindow(hEditDest, FALSE);
  EnableWindow(hButtonSource, FALSE);
  EnableWindow(hButtonDest, FALSE);
  EnableWindow(hButtonStart, FALSE);
  EnableWindow(hButtonOption, FALSE);

  // Get paths
  uint32_t uiLength;

  uiLength = SendMessage(hEditSource, WM_GETTEXTLENGTH, 0, 0);
  pszOpenFilePath = (WCHAR *)calloc(uiLength + 1, sizeof(WCHAR));
  SendMessage(hEditSource, WM_GETTEXT, uiLength + 1, (LPARAM)pszOpenFilePath);
  uiLength = SendMessage(hEditDest, WM_GETTEXTLENGTH, 0, 0);
  pszSaveFolderPath = (WCHAR *)calloc(uiLength + 1, sizeof(WCHAR));
  SendMessage(hEditDest, WM_GETTEXT, uiLength + 1, (LPARAM)pszSaveFolderPath);

  std::wstring wsOpenFile(pszOpenFilePath);
  std::wstring wsSaveFolder(pszSaveFolderPath);

  if (wsSaveFolder.back() != L'\\') {
    wsSaveFolder.push_back(L'\\');
  }

  SAFE_FREE(pszOpenFilePath);
  SAFE_FREE(pszSaveFolderPath);

  // Read bexcel
  kukdh1::BexcelFile file(wsOpenFile.c_str(), NULL, [&](uint32_t idx, uint32_t count, void *arg) {
    WCHAR szTemp[64];

    wsprintf(szTemp, STRING_PROGRESS_READING, bSplit ? 2 : 3, idx, count);
    SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)szTemp);

    if (idx == 0) {
      SendMessage(hProgress, PBM_SETRANGE32, 0, count);
    }
    SendMessage(hProgress, PBM_SETPOS, idx, NULL);
  });

  // Create excel file
  WCHAR szTemp[64];
  std::mutex m;
  uint32_t sheetidx = 0;
  uint32_t count = file.vTable.size();
  std::function<void(lxw_worksheet *, kukdh1::BexcelTable &)> work = [&](lxw_worksheet *ws, kukdh1::BexcelTable &table) {
    std::string temp;

    uint32_t rowidx = 0;
    for (auto row : table.vTableData) {
      uint32_t colidx = 0;

      for (auto col : row) {
        ConvertString(col, temp);
        worksheet_write_string(ws, rowidx, colidx, temp.c_str(), NULL);
        colidx++;
      }

      rowidx++;
    }

    worksheet_freeze_panes(ws, 1, 0);
  };

  wsprintf(szTemp, STRING_PROGRESS_PROCESSING, bSplit ? 2 : 3, sheetidx, count);
  SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)szTemp);
  SendMessage(hProgress, PBM_SETPOS, sheetidx, NULL);

  if (bSplit) {
    concurrency::parallel_for((size_t)0, file.vTable.size(), [&](size_t i) {
      lxw_workbook *wb;
      lxw_worksheet *ws;
      std::wstring filename(wsSaveFolder);

      filename.append(file.vTable.at(i).wsTableName);
      filename.append(L".xlsx");

      std::string temp;
      ConvertString(filename, temp);

      wb = workbook_new(temp.c_str());

      ConvertString(file.vTable.at(i).wsTableName, temp);
      ws = workbook_add_worksheet(wb, temp.c_str());

      work(ws, file.vTable.at(i));

      workbook_close(wb);

      {
        std::lock_guard<std::mutex> guard(m);
        WCHAR szTemp[32];

        sheetidx++;

        wsprintf(szTemp, STRING_PROGRESS_PROCESSING, 2, sheetidx, count);
        SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)szTemp);
        SendMessage(hProgress, PBM_SETPOS, sheetidx, NULL);
      }
    });
  }
  else {
    lxw_workbook *wb;
    std::wstring filename(wsSaveFolder);

    filename.append(L"datasheets.xlsx");

    std::string temp;
    ConvertString(filename, temp);

    wb = workbook_new(temp.c_str());

    std::vector<lxw_worksheet *> vSheets;
    for (auto iter : file.vTable) {
      std::string temp;

      ConvertString(iter.wsTableName, temp);
      vSheets.push_back(workbook_add_worksheet(wb, temp.c_str()));
    }

    concurrency::parallel_for((size_t)0, vSheets.size(), [&](size_t i) {
      work(vSheets.at(i), file.vTable.at(i));

      {
        std::lock_guard<std::mutex> guard(m);
        WCHAR szTemp[32];

        sheetidx++;

        wsprintf(szTemp, STRING_PROGRESS_PROCESSING, 3, sheetidx, count);
        SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)szTemp);
        SendMessage(hProgress, PBM_SETPOS, sheetidx, NULL);
      }
    });

    SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)STRING_PROGRESS_WRITING);

    workbook_close(wb);
  }

  SendMessage(hStatusBar, SB_SETTEXT, 0, (LPARAM)STRING_PROGRESS_IDLE);
  SendMessage(hProgress, PBM_SETPOS, 0, NULL);

  EnableWindow(hEditSource, TRUE);
  EnableWindow(hEditDest, TRUE);
  EnableWindow(hButtonSource, TRUE);
  EnableWindow(hButtonDest, TRUE);
  EnableWindow(hButtonStart, TRUE);
  EnableWindow(hButtonOption, TRUE);

  return 0;
}

BOOL FileOpenDialog(WCHAR **pszPath) {
  HRESULT hr;
  BOOL bOpened = FALSE;
  IFileOpenDialog *pDlg;

  hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_IFileOpenDialog, reinterpret_cast<void **>(&pDlg));
  if (SUCCEEDED(hr)) {
    COMDLG_FILTERSPEC cfFilter[] = { { L"Bexcel files", L"*.bexcel" } };

    pDlg->SetFileTypes(1, cfFilter);
    hr = pDlg->Show(NULL);

    if (SUCCEEDED(hr)) {
      IShellItem *pItem;
      hr = pDlg->GetResult(&pItem);

      if (SUCCEEDED(hr)) {
        PWSTR pszTemp;
        hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszTemp);

        if (SUCCEEDED(hr)) {
          *pszPath = (WCHAR *)calloc(wcslen(pszTemp) + 1, sizeof(WCHAR));
          wcscpy_s(*pszPath, wcslen(pszTemp) + 1, pszTemp);
          CoTaskMemFree(pszTemp);
          bOpened = TRUE;
        }

        pItem->Release();
      }
    }

    pDlg->Release();
  }

  return bOpened;
}

int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData) {
  switch (uMsg) {
  case BFFM_INITIALIZED:
    if (lpData != NULL) {
      SendMessage(hwnd, BFFM_SETSELECTION, TRUE, (LPARAM)lpData);
    }
    break;
  }

  return 0;
}

BOOL BrowseFolder(HWND hParent, LPCWSTR szTitle, LPCWSTR szStartPath, WCHAR *szFolder, DWORD dwBufferLength) {
  LPMALLOC pMalloc;
  LPITEMIDLIST pidl;
  BROWSEINFO bi;

  bi.hwndOwner = hParent;
  bi.pidlRoot = NULL;
  bi.pszDisplayName = NULL;
  bi.lpszTitle = szTitle;
  bi.ulFlags = 0;
  bi.lpfn = BrowseCallbackProc;
  bi.lParam = (LPARAM)szStartPath;

  pidl = SHBrowseForFolder(&bi);

  if (pidl == NULL) {
    return FALSE;
  }

  if (SHGetMalloc(&pMalloc) != NOERROR) {
    return FALSE;
  }

  if (!SHGetPathFromIDListEx(pidl, szFolder, dwBufferLength, GPFIDL_DEFAULT)) {
    return FALSE;
  }

  pMalloc->Free(pidl);
  pMalloc->Release();

  return TRUE;
}

BOOL ConvertString(std::wstring &in, std::string &out) {
  uint32_t newlen;

  newlen = WideCharToMultiByte(CP_UTF8, NULL, in.c_str(), in.size(), NULL, NULL, NULL, NULL);
  out.resize(newlen);
  WideCharToMultiByte(CP_UTF8, NULL, in.c_str(), in.size(), (char *)out.c_str(), out.size(), NULL, NULL);

  return TRUE;
}
