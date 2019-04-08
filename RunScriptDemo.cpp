// RunScriptDemo.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include "SimpleScriptSite.h"
#include <iostream>
#include <stdio.h>

#include <initguid.h>

// { 16d51579-a30b-4c8b-a276-0ff4dc41e755 }
DEFINE_GUID(CLSID_JSCRIPT9, 0x16d51579, 0xa30b, 0x4c8b, 0xa2, 0x76, 0x0f, 0xf4, 0xdc, 0x41, 0xe7, 0x55);
//{1b7cd997 - e5ff - 4932 - a7a6 - 2a9e636da385}
DEFINE_GUID(CLSID_CHAKRA, 0x1b7cd997, 0xe5ff, 0x4932, 0xa7, 0xa6, 0x2a, 0x9e, 0x63, 0x6d, 0xa3, 0x85);


void testExpression(const wchar_t* prefix, IActiveScriptParse* pScriptParse, LPCOLESTR script)
{
    wprintf(L"%s: %s: ", prefix, script);
    HRESULT hr = S_OK;
    CComVariant result;
    EXCEPINFO ei = { };
    hr = pScriptParse->ParseScriptText(script, NULL, NULL, NULL, 0, 0, SCRIPTTEXT_ISEXPRESSION, &result, &ei);
    CComVariant resultBSTR;
    VariantChangeType(&resultBSTR, &result, 0, VT_BSTR);
    wprintf(L"%s\n", V_BSTR(&resultBSTR));
}

void testScript(const wchar_t* prefix, IActiveScriptParse* pScriptParse, LPCOLESTR script)
{
    wprintf(L"%s: %s\n", prefix, script);
    HRESULT hr = S_OK;
    CComVariant result;
    EXCEPINFO ei = { };
    hr = pScriptParse->ParseScriptText(script, NULL, NULL, NULL, 0, 0, 0, &result, &ei);
}


int _tmain(int argc, _TCHAR* argv[])
{
    HRESULT hr = S_OK;
    hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

    // Initialize
    CSimpleScriptSite* pScriptSite = new CSimpleScriptSite();
    CComPtr<IActiveScript> spJScript;
    CComPtr<IActiveScriptParse> spJScriptParse;
    //hr = spJScript.CoCreateInstance(OLESTR("JScript"));
    //hr = spJScript.CoCreateInstance(CLSID_CHAKRA);
    hr = spJScript.CoCreateInstance(CLSID_JSCRIPT9);
    hr = spJScript->SetScriptSite(pScriptSite);
    hr = spJScript->QueryInterface(&spJScriptParse);
    hr = spJScriptParse->InitNew();
    CComPtr<IActiveScript> spVBScript;
    CComPtr<IActiveScriptParse> spVBScriptParse;
    hr = spVBScript.CoCreateInstance(OLESTR("VBScript"));
    hr = spVBScript->SetScriptSite(pScriptSite);
    hr = spVBScript->QueryInterface(&spVBScriptParse);
    hr = spVBScriptParse->InitNew();

    // Run some scripts
    
    testExpression(L"JScript", spJScriptParse, OLESTR("(new Date).getTime()"));
    testExpression(L"JScript", spJScriptParse, OLESTR("ScriptEngineMajorVersion()"));
    testExpression(L"JScript", spJScriptParse, OLESTR("let x=1"));
    testExpression(L"JScript", spJScriptParse, OLESTR("x"));
    testExpression(L"VBScript", spVBScriptParse, OLESTR("Now"));
    testScript(L"VBScript", spVBScriptParse, OLESTR("MsgBox \"Hello World! The current time is: \" & Now"));
    // Outputs:
    // JScript: (new Date).getTime() : 1554685325378
    // VScript: Now : 8 / 04 / 2019 11 : 02 : 05 AM
    // VBScript: MsgBox "Hello World! The current time is: " & Now

    // Cleanup
    spVBScriptParse = NULL;
    spVBScript = NULL;
    spJScriptParse = NULL;
    spJScript = NULL;
    pScriptSite->Release();
    pScriptSite = NULL;

    ::CoUninitialize();
    return 0;
}
