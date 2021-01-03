#include <imapi2.h>
#include <ntverp.h>
#include <iostream>

#include "HResultException.h"


#define DISC_INDEX 0

#define ReleaseAndNull(x)       \
{                               \
    if ((x) != nullptr)            \
    {                           \
        (x)->Release();         \
        (x) = nullptr;             \
    }                           \
}

#define SafeArrayDestroyDataAndNull(x) \
{                                      \
    if ((x) != nullptr)                   \
    {                                  \
        SafeArrayDestroyData(x);       \
        (x) = nullptr;                    \
    }                                  \
}

#define CHECK_RESULT(result, msg)            \
{                                            \
    if (FAILED(result))                      \
    {                                        \
        throw HResultException(msg, result); \
    }                                        \
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

static void FreeSysStringAndNull(BSTR& t)
{
    ::SysFreeString(t);
    t = nullptr;
    return;
}


// Get a disc recorder given a disc index
static void GetDiscRecorder(__in ULONG index, __out IDiscRecorder2** recorder)
{
    IDiscMaster2* tmpDiscMaster = nullptr;
    BSTR tmpUniqueId = nullptr;
    IDiscRecorder2* tmpRecorder = nullptr;

    *recorder = nullptr;

    try {

        // create the disc master object
        CHECK_RESULT(CoCreateInstance(CLSID_MsftDiscMaster2, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&tmpDiscMaster)),
        "Failed CoCreateInstance\n");

        // get the unique id string
        CHECK_RESULT(tmpDiscMaster->get_Item(index, &tmpUniqueId),
        "Failed tmpDiscMaster->get_Item\n");

        // Create a new IDiscRecorder2
        CHECK_RESULT(CoCreateInstance(CLSID_MsftDiscRecorder2, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&tmpRecorder)),
        "Failed CoCreateInstance\n");

        // Initialize it with the provided BSTR
        CHECK_RESULT(tmpRecorder->InitializeDiscRecorder(tmpUniqueId),
        "Failed to init disc recorder\n");

        // copy to caller or release recorder
        *recorder = tmpRecorder;
    }

    catch(const HResultException&)
    {
        ReleaseAndNull(tmpRecorder);

        ReleaseAndNull(tmpDiscMaster);
        FreeSysStringAndNull(tmpUniqueId);
        throw;
    }

    // all other cleanup
    ReleaseAndNull(tmpDiscMaster);
    FreeSysStringAndNull(tmpUniqueId);
}

static int ListAllRecorders()
{
    int ret = -1;
    LONG count = 0;
    IDiscMaster2* tmpDiscMaster = nullptr;
    IDiscFormat2Data* dataWriter = nullptr;
    BSTR discId = nullptr;
    BSTR venId = nullptr;
    IDiscRecorder2* discRecorder = nullptr;
    SAFEARRAY* mountPoints = nullptr;

    // create a disc master object
    CHECK_RESULT(CoCreateInstance(CLSID_MsftDiscMaster2, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&tmpDiscMaster)),
    "Failed CoCreateInstance\n");

    try
    {
        // Get number of recorders
        CHECK_RESULT(tmpDiscMaster->get_Count(&count), "Failed to get count\n");

        if (count > 0)
        {

            GetDiscRecorder(DISC_INDEX, &discRecorder);

            // Get the device strings
            try
            {
                CHECK_RESULT(discRecorder->get_VendorId(&venId), "Failed to get Vendor ID.");
                CHECK_RESULT(discRecorder->get_ProductId(&venId), "Failed to get Product ID.");

                printf("Recorder %d: %ws %ws", DISC_INDEX, venId, discId);
            }
            catch(const HResultException& e)
            {
                std::cerr << e.what() << '\n';
            }

            CHECK_RESULT(discRecorder->get_VolumePathNames(&mountPoints), "Unable to get mount points, failed\n");

            if (mountPoints->rgsabound[0].cElements == 0)
            {
                throw std::runtime_error("(*** NO MOUNT POINTS ***)");
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

            // create a DiscFormat2Data object
            CHECK_RESULT(CoCreateInstance(CLSID_MsftDiscFormat2Data,
                nullptr, CLSCTX_ALL,
                IID_PPV_ARGS(&dataWriter)), "Failed CoCreateInstance on dataWriter\n");

            CHECK_RESULT(dataWriter->put_Recorder(discRecorder), "Failed dataWriter->put_Recorder\n");

            // get the current media in the recorder
            IMAPI_MEDIA_PHYSICAL_TYPE mediaType = IMAPI_MEDIA_TYPE_UNKNOWN;
            HRESULT hr = dataWriter->get_CurrentPhysicalMediaType(&mediaType);
            if (SUCCEEDED(hr))
            {
                printf(" (%s)", g_MediaTypeStrings[mediaType]);
                ret = mediaType;
            }
            else
            {
                throw HResultException("Error getting media type", hr);
            }

            printf("\n");
        }

        ReleaseAndNull(dataWriter);

        FreeSysStringAndNull(venId);
        FreeSysStringAndNull(discId);

        ReleaseAndNull(discRecorder);
    }
    catch (const HResultException&)
    {
        ReleaseAndNull(dataWriter);

        FreeSysStringAndNull(venId);
        FreeSysStringAndNull(discId);

        ReleaseAndNull(discRecorder);
        throw;
    }

    return ret;
}

int main(int argc, WCHAR* argv[])
{
    HRESULT coInitHr = S_OK;
    int ret = -1;

    coInitHr = CoInitialize(nullptr);

    try
    {
        CHECK_RESULT(coInitHr, "CoInitialize failed.\n");
        try
        {
            ret = ListAllRecorders();
        }
        catch(const HResultException& e)
        {
            std::cout << e.what() << std::endl;
            ret = -2;
        }

        CoUninitialize();
    }
    catch(const HResultException& e)
    {
        std::cout << e.what() << std::endl;
        ret = -3;
    }

    return ret;
}
