#pragma once

#ifndef _BEXCEL_H_
#define _BEXCEL_H_

#include <iostream>
#include <string>
#include <vector>
#include <functional>
#include <Windows.h>

namespace kukdh1 {
  struct BexcelTable {
      std::wstring wsTableName;
      std::vector<std::vector<std::wstring>> vTableData;

      BexcelTable(HANDLE hFile);
  };

  struct BexcelFile {
      std::vector<BexcelTable> vTable;

      BexcelFile(const wchar_t *wpszFilePath, void *lpParam, std::function<void (uint32_t, uint32_t, void *)> fProgress);
  };
}

#endif
