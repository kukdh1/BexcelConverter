#pragma once

#ifndef _MAIN_H_
#define _MAIN_H_

#include <Windows.h>
#include <CommCtrl.h>
#include <ShObjIdl.h>
#include <ShlObj.h>
#include <ppl.h>
#include <fstream>
#include <string>
#include <functional>
#include <mutex>
#include <thread>

#include <xlsxwriter.h>

#include "Bexcel.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#ifdef _DEBUG
#pragma comment(lib, "libxlsxwriterd.lib")
#pragma comment(lib, "zlibstaticd.lib")
#else
#pragma comment(lib, "libxlsxwriter.lib")
#pragma comment(lib, "zlibstatic.lib")
#endif

#define WINDOW_WIDTH    400

#define CONTROL_MARGIN  10
#define CONTROL_HEIGHT  24
#define EDIT_WIDTH      300
#define BUTTON_WIDTH    (WINDOW_WIDTH - EDIT_WIDTH - CONTROL_MARGIN * 3)
#define PROGRESS_WIDTH  (WINDOW_WIDTH - CONTROL_MARGIN * 2)

#define ID_BUTTON_OPEN_SRC  0
#define ID_BUTTON_OPEN_DST  1
#define ID_BUTTON_START     2
#define ID_EDIT_SRC         10
#define ID_EDIT_DST         11
#define ID_PROGRESS         12
#define ID_STATUS_BAR       20
#define ID_BUTTON_OPTION    30

#define FONT_SIZE       17
#define FONT_FACE       L"Segoe UI"

#define STRING_PROGRESS_IDLE        L"Idle"
#define STRING_PROGRESS_READING     L"(1/%d) Reading %d/%d..."
#define STRING_PROGRESS_PROCESSING  L"(2/%d) Processing %d/%d..."
#define STRING_PROGRESS_WRITING     L"(3/3) Saving..."
#define STRING_OPEN                 L"Open"
#define STRING_SAVE                 L"Save"
#define STRING_START                L"Start"
#define STRING_OPTION               L"Save sheets as individual file (fast)"

#define SAFE_FREE(ptr)      { if (ptr) { free(ptr); (ptr) = NULL; } }

#endif
