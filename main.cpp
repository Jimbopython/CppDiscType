#include <imapi2.h>
#include <ntverp.h>
#include <iostream>
#include <atlbase.h>
#include <atlcom.h>

#define DISC_INDEX 0

#define ReleaseAndNull(x)       \
{                               \
    if ((x) != NULL)            \
    {                           \
        (x)->Release();         \
        (x) = NULL;             \
    }                           \
}

#define SafeArrayDestroyDataAndNull(x) \
{                                      \
    if ((x) != NULL)                   \
    {                                  \
        SafeArrayDestroyData(x);       \
        (x) = NULL;                    \
    }                                  \
}

// using a simple array due to consecutive zero-based values in this enum
const CHAR* g_MediaTypeStrings[] = {
    "IMAPI_MEDIA_TYPE_UNKNOWN",
    "IMAPI_MEDIA_TYPE_CDROM",
    "IMAPI_MEDIA_TYPE_CDR",
    "IMAPI_MEDIA_TYPE_CDRW",
    "IMAPI_MEDIA_TYPE_DVDROM",
    "IMAPI_MEDIA_TYPE_DVDRAM",
    "IMAPI_MEDIA_TYPE_DVDPLUSR",
    "IMAPI_MEDIA_TYPE_DVDPLUSRW",
    "IMAPI_MEDIA_TYPE_DVDPLUSR_DUALLAYER",
    "IMAPI_MEDIA_TYPE_DVDDASHR",
    "IMAPI_MEDIA_TYPE_DVDDASHRW",
    "IMAPI_MEDIA_TYPE_DVDDASHR_DUALLAYER",
    "IMAPI_MEDIA_TYPE_DISK",
    "IMAPI_MEDIA_TYPE_DVDPLUSRW_DUALLAYER",
    "IMAPI_MEDIA_TYPE_HDDVDROM",
    "IMAPI_MEDIA_TYPE_HDDVDR",
    "IMAPI_MEDIA_TYPE_HDDVDRAM",
    "IMAPI_MEDIA_TYPE_BDROM",
    "IMAPI_MEDIA_TYPE_BDR",
    "IMAPI_MEDIA_TYPE_BDRE",
    "IMAPI_MEDIA_TYPE_MAX"
};


static void PrintHR(HRESULT hr)
{
    LPVOID lpMsgBuf;
    DWORD ret;

    ret = FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_HMODULE,
        GetModuleHandle(TEXT("imapi2.dll")),
        hr,
        0, //MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPTSTR)&lpMsgBuf,
        0, NULL);

    if (ret != 0)
    {
        printf("  Returned %08x: %s\n", hr, lpMsgBuf);
    }

    if (ret == 0)
    {
        ret = FormatMessage(
            FORMAT_MESSAGE_ALLOCATE_BUFFER |
            FORMAT_MESSAGE_FROM_HMODULE,
            GetModuleHandle(TEXT("imapi2fs.dll")),
            hr,
            0, //MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPTSTR)&lpMsgBuf,
            0, NULL);

        if (ret != 0)
        {
            printf("  Returned %08x: %s\n", hr, lpMsgBuf);
        }
    }

    if (ret == 0)
    {
        ret = FormatMessage(
            FORMAT_MESSAGE_ALLOCATE_BUFFER |
            FORMAT_MESSAGE_FROM_SYSTEM,
            NULL,
            hr,
            0, //MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
            (LPTSTR)&lpMsgBuf,
            0, NULL);

        if (ret != 0)
        {
            printf("  Returned %08x: %s\n", hr, lpMsgBuf);
        }
        else
        {
            printf("  Returned %08x (no description)\n\n", hr);
        }
    }

    LocalFree(lpMsgBuf);
}

static void FreeSysStringAndNull(BSTR& t)
{
    ::SysFreeString(t);
    t = NULL;
    return;
}


// Get a disc recorder given a disc index
static HRESULT GetDiscRecorder(__in ULONG index, __out IDiscRecorder2** recorder)
{
    HRESULT hr = S_OK;
    IDiscMaster2* tmpDiscMaster = NULL;
    BSTR tmpUniqueId = NULL;
    IDiscRecorder2* tmpRecorder = NULL;

    *recorder = NULL;

    // create the disc master object
    if (SUCCEEDED(hr))
    {
        hr = CoCreateInstance(CLSID_MsftDiscMaster2,
            NULL, CLSCTX_ALL,
            IID_PPV_ARGS(&tmpDiscMaster)
        );
        if (FAILED(hr))
        {
            printf("Failed CoCreateInstance\n");
            PrintHR(hr);
        }
    }

    // get the unique id string
    if (SUCCEEDED(hr))
    {
        hr = tmpDiscMaster->get_Item(index, &tmpUniqueId);
        if (FAILED(hr))
        {
            printf("Failed tmpDiscMaster->get_Item\n");
            PrintHR(hr);
        }
    }

    // Create a new IDiscRecorder2
    if (SUCCEEDED(hr))
    {
        hr = CoCreateInstance(CLSID_MsftDiscRecorder2,
            NULL, CLSCTX_ALL,
            IID_PPV_ARGS(&tmpRecorder)
        );
        if (FAILED(hr))
        {
            printf("Failed CoCreateInstance\n");
            PrintHR(hr);
        }
    }
    // Initialize it with the provided BSTR
    if (SUCCEEDED(hr))
    {
        hr = tmpRecorder->InitializeDiscRecorder(tmpUniqueId);
        if (FAILED(hr))
        {
            printf("Failed to init disc recorder\n");
            PrintHR(hr);
        }
    }
    // copy to caller or release recorder
    if (SUCCEEDED(hr))
    {
        *recorder = tmpRecorder;
    }
    else
    {
        ReleaseAndNull(tmpRecorder);
    }
    // all other cleanup
    ReleaseAndNull(tmpDiscMaster);
    FreeSysStringAndNull(tmpUniqueId);
    return hr;
}

static HRESULT ListAllRecorders()
{
    HRESULT hr = S_OK;
    LONG          count = 0;
    IDiscMaster2* tmpDiscMaster = NULL;

    // create a disc master object
    if (SUCCEEDED(hr))
    {
        hr = CoCreateInstance(CLSID_MsftDiscMaster2,
            NULL, CLSCTX_ALL,
            IID_PPV_ARGS(&tmpDiscMaster)
        );
        if (FAILED(hr))
        {
            printf("Failed CoCreateInstance\n");
            PrintHR(hr);
        }
    }
    // Get number of recorders
    if (SUCCEEDED(hr))
    {
        hr = tmpDiscMaster->get_Count(&count);

        if (FAILED(hr))
        {
            printf("Failed to get count\n");
            PrintHR(hr);
        }
    }

    if (count > 0)
    {

        IDiscRecorder2* discRecorder = NULL;

        hr = GetDiscRecorder(DISC_INDEX, &discRecorder);

        if (SUCCEEDED(hr))
        {

            BSTR discId = NULL;
            BSTR venId = NULL;

            // Get the device strings
            if (SUCCEEDED(hr)) { hr = discRecorder->get_VendorId(&venId); }
            if (SUCCEEDED(hr)) { hr = discRecorder->get_ProductId(&discId); }
            if (FAILED(hr))
            {
                printf("Failed to get ID's\n");
                PrintHR(hr);
            }

            if (SUCCEEDED(hr))
            {
                printf("Recorder %d: %ws %ws", DISC_INDEX, venId, discId);
            }
            // Get the mount point
            if (SUCCEEDED(hr))
            {
                SAFEARRAY* mountPoints = NULL;
                hr = discRecorder->get_VolumePathNames(&mountPoints);
                if (FAILED(hr))
                {
                    printf("Unable to get mount points, failed\n");
                    PrintHR(hr);
                }
                else if (mountPoints->rgsabound[0].cElements == 0)
                {
                    printf(" (*** NO MOUNT POINTS ***)");
                }
                else
                {
                    VARIANT* tmp = (VARIANT*)(mountPoints->pvData);
                    printf(" (");
                    for (ULONG j = 0; j < mountPoints->rgsabound[0].cElements; j++)
                    {
                        printf(" %ws ", tmp[j].bstrVal);
                    }
                    printf(")");
                }
                SafeArrayDestroyDataAndNull(mountPoints);
            }
            // Get the media type
            if (SUCCEEDED(hr))
            {
                IDiscFormat2Data* dataWriter = NULL;

                // create a DiscFormat2Data object
                if (SUCCEEDED(hr))
                {
                    hr = CoCreateInstance(CLSID_MsftDiscFormat2Data,
                        NULL, CLSCTX_ALL,
                        IID_PPV_ARGS(&dataWriter)
                    );
                    if (FAILED(hr))
                    {
                        printf("Failed CoCreateInstance on dataWriter\n");
                        PrintHR(hr);
                    }
                }

                if (SUCCEEDED(hr))
                {
                    hr = dataWriter->put_Recorder(discRecorder);
                    if (FAILED(hr))
                    {
                        printf("Failed dataWriter->put_Recorder\n");
                        PrintHR(hr);
                    }
                }
                // get the current media in the recorder
                if (SUCCEEDED(hr))
                {
                    IMAPI_MEDIA_PHYSICAL_TYPE mediaType = IMAPI_MEDIA_TYPE_UNKNOWN;
                    hr = dataWriter->get_CurrentPhysicalMediaType(&mediaType);
                    if (SUCCEEDED(hr))
                    {
                        printf(" (%s)", g_MediaTypeStrings[mediaType]);
                    }
                }
                ReleaseAndNull(dataWriter);
            }

            printf("\n");
            FreeSysStringAndNull(venId);
            FreeSysStringAndNull(discId);

        }
        else
        {
            printf("Failed to get drive %d\n", DISC_INDEX);
        }

        ReleaseAndNull(discRecorder);
    }

    return hr;
}

int __cdecl wmain(int argc, WCHAR* argv[])
{
    HRESULT coInitHr = S_OK;
    HRESULT hr = S_OK;

    coInitHr = CoInitialize(nullptr);

    if (CAtlBaseModule::m_bInitFailed)
    {
        printf("AtlBaseInit failed...\n");
        coInitHr = E_FAIL;
    }
    else
    {
        coInitHr = S_OK;
    }

    if (SUCCEEDED(coInitHr))
    {
        hr = ListAllRecorders();
        CoUninitialize();
    }

    if (SUCCEEDED(hr))
    {
        return 0;
    }
    else
    {
        PrintHR(hr);
        return 1;
    }
}
