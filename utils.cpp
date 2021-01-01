#include "utils.h"

#include <iostream>

void PrintHR(HRESULT hr)
{
    LPSTR lpMsgBuf = nullptr;
    
    DWORD ret = 0;
    HMODULE modules[] = {GetModuleHandle(TEXT("imapi2.dll")), GetModuleHandle(TEXT("imapi2fs.dll")), nullptr};

    for (const auto& hModule : modules)
    {
        ret = FormatMessage(
            FORMAT_MESSAGE_ALLOCATE_BUFFER |
            FORMAT_MESSAGE_FROM_HMODULE,
            hModule,
            hr,
            0, //MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            lpMsgBuf,
            0, nullptr);

        if (ret != 0)
        {
            printf("  Returned %08x: %s\n", hr, lpMsgBuf);
            break;
        }
    }

    LocalFree(lpMsgBuf);
}