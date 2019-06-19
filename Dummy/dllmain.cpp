// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "unknwn.h"

extern "C++"
{
    MIDL_INTERFACE("00000000-0000-0000-C000-000000000047")
        IDummy //: public IUnknown
    {
        virtual HRESULT STDMETHODCALLTYPE QueryInterface(
            REFIID riid, void** ppvObject) = 0;

        virtual ULONG STDMETHODCALLTYPE AddRef() = 0;

        virtual ULONG STDMETHODCALLTYPE Release() = 0;

        // trivial method to troubleshoot x86 calling convention issue
        virtual void STDMETHODCALLTYPE Void() = 0;
    };
}

struct dummy : public IDummy
{ 
    dummy() : _ref_count(1)
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(
        REFIID /*riid*/,
        void ** /*ppvObject*/) override
    {
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++_ref_count;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        return --_ref_count;
    }

    void STDMETHODCALLTYPE Void() override
    {
    }

    int _ref_count{ 0 };
};

extern "C"
    __declspec(dllexport) void* CreateDummyUnknown()
{
    return new dummy;
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

