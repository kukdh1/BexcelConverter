#include "Bexcel.h"

namespace kukdh1 {
  BexcelTable::BexcelTable(HANDLE hFile) {
    DWORD dwRead;
    uint64_t uiLength;

    ReadFile(hFile, &uiLength, 8, &dwRead, NULL);
    wsTableName.resize(uiLength);
    ReadFile(hFile, (wchar_t *)wsTableName.c_str(), uiLength * 2, &dwRead, NULL);
    SetFilePointer(hFile, 8, 0, FILE_CURRENT);   // skip 0xFFFFFFFF and row index

    uint32_t uiColCount;
    std::vector<std::wstring> vRowTemp;

    ReadFile(hFile, &uiColCount, 4, &dwRead, NULL);
    vRowTemp.resize(uiColCount);

    for (uint32_t i = 0; i < uiColCount; i++) {
      ReadFile(hFile, &uiLength, 8, &dwRead, NULL);
      vRowTemp.at(i).resize(uiLength);
      ReadFile(hFile, (wchar_t *)vRowTemp.at(i).c_str(), uiLength * 2, &dwRead, NULL);
    }

    uint32_t uiRowCount;

    ReadFile(hFile, &uiRowCount, 4, &dwRead, NULL);
    vTableData.resize(uiRowCount + 1);
    vTableData.at(0) = vRowTemp;

    for (uint32_t i = 0; i < uiRowCount; i++) {
      SetFilePointer(hFile, 12, 0, FILE_CURRENT);   // skip row indexes and col count

      for (uint32_t j = 0; j < uiColCount; j++) {
        ReadFile(hFile, &uiLength, 8, &dwRead, NULL);
        vRowTemp.at(j).resize(uiLength);
        ReadFile(hFile, (wchar_t *)vRowTemp.at(j).c_str(), uiLength * 2, &dwRead, NULL);
      }

      vTableData.at(i + 1) = vRowTemp;
    }

    /* skip column names and row names */
    ReadFile(hFile, &uiColCount, 4, &dwRead, NULL);

    for (uint32_t i = 0; i < uiColCount; i++) {
      ReadFile(hFile, &uiLength, 8, &dwRead, NULL);
      SetFilePointer(hFile, uiLength * 2 + 4, 0, FILE_CURRENT);
    }

    ReadFile(hFile, &uiRowCount, 4, &dwRead, NULL);

    for (uint32_t i = 0; i < uiRowCount; i++) {
      ReadFile(hFile, &uiLength, 8, &dwRead, NULL);
      SetFilePointer(hFile, uiLength * 2 + 4, 0, FILE_CURRENT);
    }
  }

  BexcelFile::BexcelFile(const wchar_t *wpszFilePath, void *lpParam, std::function<void(uint32_t, uint32_t, void *)> fProgress) {
    HANDLE hFile;
    DWORD dwRead;

    hFile = CreateFile(wpszFilePath, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile != INVALID_HANDLE_VALUE) {
      uint32_t uiTableCount;

      ReadFile(hFile, &uiTableCount, 4, &dwRead, NULL);
      
      /* Skip name tables */
      for (uint32_t i = 0; i < uiTableCount; i++) {
        uint64_t uiRecordLength;

        ReadFile(hFile, &uiRecordLength, 8, &dwRead, NULL);
        SetFilePointer(hFile, uiRecordLength * 2 + 4, 0, FILE_CURRENT);
      }

      ReadFile(hFile, &uiTableCount, 4, &dwRead, NULL);

      for (uint32_t i = 0; i < uiTableCount; i++) {
        fProgress(i, uiTableCount, lpParam);
        vTable.push_back(BexcelTable(hFile));
      }
      fProgress(uiTableCount, uiTableCount, lpParam);

      CloseHandle(hFile);
    }
  }
}
