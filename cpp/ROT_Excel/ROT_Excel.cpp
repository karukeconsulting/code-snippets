// ROT_Excel.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include <iostream>

#include <atlbase.h>
#include <comdef.h>
#include <vector>
#include <memory>
#include <functional>

// {00020813-0000-0000-C000-000000000046}
static const GUID xl_typelib_guid = { 0x00020813, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };


void throw_native_if_fail(HRESULT hr)
{
    if (FAILED(hr))
    {
        throw _com_error(hr);
    }
}

void olestr_imalloc_free(LPOLESTR ptr)
{
    if (ptr)
    {
        CComPtr<IMalloc> im;
        throw_native_if_fail(CoGetMalloc(1, &im));
        im->Free(ptr);
    }
}

std::wstring GetGUIDStr(GUID _uuid)
{
    std::vector<OLECHAR> guidstr(100, 0);
    int _numchar{ 0 };
    _numchar = StringFromGUID2(_uuid, const_cast<OLECHAR*>(guidstr.data()), guidstr.size());
    if (_numchar > 0) _numchar -= 1;
    return std::wstring(guidstr.begin(), guidstr.begin() + _numchar);
}

std::wstring GetTypeKind(TYPEKIND tpKind)
{
    switch (tpKind)
    {
    case TYPEKIND::TKIND_ALIAS:
        return L"Alias";
    case TYPEKIND::TKIND_COCLASS:
        return L"CoClass";
    case TYPEKIND::TKIND_DISPATCH:
        return L"Dispatch";
    case TYPEKIND::TKIND_ENUM:
        return L"Enum";
    case TYPEKIND::TKIND_INTERFACE:
        return L"Interface";
    case TYPEKIND::TKIND_MAX:
        return L"Max";
    case TYPEKIND::TKIND_MODULE:
        return L"Module";
    case TYPEKIND::TKIND_RECORD:
        return L"Record";
    case TYPEKIND::TKIND_UNION:
        return L"Union";
    }
}

std::wstring GetFuncKind(FUNCKIND funckind)
{
    switch (funckind)
    {
    case FUNCKIND::FUNC_DISPATCH:
        return L"Dispatch";
    case FUNCKIND::FUNC_NONVIRTUAL:
        return L"NonVirtual";
    case FUNCKIND::FUNC_PUREVIRTUAL:
        return L"PureVirtual";
    case FUNCKIND::FUNC_STATIC:
        return L"Static";
    case FUNCKIND::FUNC_VIRTUAL:
        return L"Virtual";
    }
}

std::wstring GetInvokeKind(INVOKEKIND invkind)
{
    switch (invkind)
    {
    case INVOKEKIND::INVOKE_FUNC:
        return L"Func";
    case INVOKEKIND::INVOKE_PROPERTYGET:
        return L"PropertyGet";
    case INVOKEKIND::INVOKE_PROPERTYPUT:
        return L"PropertyPut";
    case INVOKEKIND::INVOKE_PROPERTYPUTREF:
        return L"PropertyPutRef";
    }
}

std::wstring GetCallConv(CALLCONV callconv)
{
    switch (callconv)
    {
    case CALLCONV::CC_CDECL:
        return L"_cdecl";
    case CALLCONV::CC_FASTCALL:
        return L"_fastcall";
    case CALLCONV::CC_FPFASTCALL:
        return L"_fpfastcall";
    case CALLCONV::CC_MACPASCAL:
        return L"_macpascal";
    case CALLCONV::CC_MAX:
        return L"_max";
    case CALLCONV::CC_MPWCDECL:
        return L"_mpwcdecl";
    case CALLCONV::CC_MPWPASCAL:
        return L"_mpwpascal";
    case CALLCONV::CC_PASCAL:
        return L"_pascal";
    case CALLCONV::CC_STDCALL:
        return L"_stdcall";
    case CALLCONV::CC_SYSCALL:
        return L"_syscall";
    }
}

std::wstring GetVarKind(VARKIND varkind)
{
    switch (varkind)
    {
    case VARKIND::VAR_CONST:
        return L"Const";
    case VARKIND::VAR_DISPATCH:
        return L"Dispatch";
    case VARKIND::VAR_PERINSTANCE:
        return L"PerInstance";
    case VARKIND::VAR_STATIC:
        return L"Static";
    }
}

std::wstring GetVarType(VARTYPE vt)
{
    switch (vt)
    {
    case VARENUM::VT_EMPTY:
        return L"Empty";
    case VARENUM::VT_NULL:
        return L"Null";
    case VARENUM::VT_I2:
        return L"I2";
    case VARENUM::VT_I4:
        return L"I4";
    case VARENUM::VT_R4:
        return L"R4";
    case VARENUM::VT_R8:
        return L"R8";
    case VARENUM::VT_CY:
        return L"CY";
    case VARENUM::VT_DATE:
        return L"Date";
    case VARENUM::VT_BSTR:
        return L"BSTR";
    case VARENUM::VT_DISPATCH:
        return L"Dispatch";
    case VARENUM::VT_ERROR:
        return L"Error";
    case VARENUM::VT_BOOL:
        return L"Bool";
    case VARENUM::VT_VARIANT:
        return L"Variant";
    case VARENUM::VT_UNKNOWN:
        return L"Unknown";
    case VARENUM::VT_DECIMAL:
        return L"Decimal";
    case VARENUM::VT_I1:
        return L"I1";
    case VARENUM::VT_UI1:
        return L"UI1";
    case VARENUM::VT_UI2:
        return L"UI2";
    case VARENUM::VT_UI4:
        return L"UI4";
    case VARENUM::VT_I8:
        return L"I8";
    case VARENUM::VT_UI8:
        return L"UI8";
    case VARENUM::VT_INT:
        return L"Int";
    case VARENUM::VT_UINT:
        return L"UInt";
    case VARENUM::VT_VOID:
        return L"Void";
    case VARENUM::VT_HRESULT:
        return L"HRESULT";
    case VARENUM::VT_PTR:
        return L"Ptr";
    case VARENUM::VT_SAFEARRAY:
        return L"SAFEARRAY";
    case VARENUM::VT_CARRAY:
        return L"CARRAY";
    case VARENUM::VT_USERDEFINED:
        return L"UserDefined";
    case VARENUM::VT_LPSTR:
        return L"LPSTR";
    case VARENUM::VT_LPWSTR:
        return L"LPWSTR";
    case VARENUM::VT_RECORD:
        return L"RECORD";
    case VARENUM::VT_INT_PTR:
        return L"Int_Ptr";
    case VARENUM::VT_UINT_PTR:
        return L"UInt_Pre";
    case VARENUM::VT_FILETIME:
        return L"Filetime";
    case VARENUM::VT_BLOB:
        return L"Blob";
    case VARENUM::VT_STREAM:
        return L"Stream";
    case VARENUM::VT_STORAGE:
        return L"Storage";
    case VARENUM::VT_STREAMED_OBJECT:
        return L"Streamed_Object";
    case VARENUM::VT_STORED_OBJECT:
        return L"Stored_Object";
    case VARENUM::VT_BLOB_OBJECT:
        return L"Blob_Object";
    case VARENUM::VT_CF:
        return L"CF";
    case VARENUM::VT_CLSID:
        return L"CLSID";
    case VARENUM::VT_VERSIONED_STREAM:
        return L"Versioned_Stream";
    case VARENUM::VT_BSTR_BLOB:
        return L"BSTR_Blob";
    case VARENUM::VT_VECTOR:
        return L"Vector";
    case VARENUM::VT_ARRAY:
        return L"Array";
    case VARENUM::VT_BYREF:
        return L"ByRef";
    case VARENUM::VT_RESERVED:
        return L"Reserved";
    case VARENUM::VT_ILLEGAL:
        return L"Illegal";
    }
}

std::wstring GetTypeFlagList(TYPEFLAGS tpFlag)
{
    std::vector<std::string> _flags;

    if ((TYPEFLAGS::TYPEFLAG_FAPPOBJECT & tpFlag) == TYPEFLAGS::TYPEFLAG_FAPPOBJECT)
        _flags.push_back("appobject");
    if ((TYPEFLAGS::TYPEFLAG_FCANCREATE & tpFlag) == TYPEFLAGS::TYPEFLAG_FCANCREATE)
        _flags.push_back("cancreate");
    if ((TYPEFLAGS::TYPEFLAG_FLICENSED & tpFlag) == TYPEFLAGS::TYPEFLAG_FLICENSED)
        _flags.push_back("licensed");
    if ((TYPEFLAGS::TYPEFLAG_FPREDECLID & tpFlag) == TYPEFLAGS::TYPEFLAG_FPREDECLID)
        _flags.push_back("predeclid");
    if ((TYPEFLAGS::TYPEFLAG_FHIDDEN & tpFlag) == TYPEFLAGS::TYPEFLAG_FHIDDEN)
        _flags.push_back("hidden");
    if ((TYPEFLAGS::TYPEFLAG_FCONTROL & tpFlag) == TYPEFLAGS::TYPEFLAG_FCONTROL)
        _flags.push_back("control");
    if ((TYPEFLAGS::TYPEFLAG_FDUAL & tpFlag) == TYPEFLAGS::TYPEFLAG_FDUAL)
        _flags.push_back("dual");
    if ((TYPEFLAGS::TYPEFLAG_FNONEXTENSIBLE & tpFlag) == TYPEFLAGS::TYPEFLAG_FNONEXTENSIBLE)
        _flags.push_back("nonextensible");
    if ((TYPEFLAGS::TYPEFLAG_FOLEAUTOMATION & tpFlag) == TYPEFLAGS::TYPEFLAG_FOLEAUTOMATION)
        _flags.push_back("oleautomation");
    if ((TYPEFLAGS::TYPEFLAG_FRESTRICTED & tpFlag) == TYPEFLAGS::TYPEFLAG_FRESTRICTED)
        _flags.push_back("restricted");
    if ((TYPEFLAGS::TYPEFLAG_FAGGREGATABLE & tpFlag) == TYPEFLAGS::TYPEFLAG_FAGGREGATABLE)
        _flags.push_back("aggregatable");
    if ((TYPEFLAGS::TYPEFLAG_FREPLACEABLE & tpFlag) == TYPEFLAGS::TYPEFLAG_FREPLACEABLE)
        _flags.push_back("replaceable");
    if ((TYPEFLAGS::TYPEFLAG_FDISPATCHABLE & tpFlag) == TYPEFLAGS::TYPEFLAG_FDISPATCHABLE)
        _flags.push_back("dispatchable");
    if ((TYPEFLAGS::TYPEFLAG_FREVERSEBIND & tpFlag) == TYPEFLAGS::TYPEFLAG_FREVERSEBIND)
        _flags.push_back("reversebind");
    if ((TYPEFLAGS::TYPEFLAG_FPROXY & tpFlag) == TYPEFLAGS::TYPEFLAG_FPROXY)
        _flags.push_back("proxy");

    return L"";
}

void printErrorMessage(HRESULT hr)
{
    _com_error err(hr);
    std::wcout << "Error *-*-*-*-*-*-*-*" << std::endl;
    std::wcout << "Message: " << err.ErrorMessage() << std::endl;
    _bstr_t _desc(err.Description());
    std::wcout << "Description was extracted" << std::endl;
    if (_desc.GetBSTR() && _desc.length() > 0)
    {
        std::wcout << "Description is as follows: " << (LPCWSTR)_desc << std::endl;
    }
    else
    {
        std::wcout << "Description information was empty" << std::endl;
    }
}

static auto fn_funcdesc_dtor = [](ITypeInfo* i) { return [&](FUNCDESC* p) { i->ReleaseFuncDesc(p); }; };
static auto fn_tattr_dtor = [](ITypeInfo* i) { return [&](TYPEATTR* p) { i->ReleaseTypeAttr(p); }; };
static auto fn_vardesc_dtor = [](ITypeInfo* i) { return [&](VARDESC* p) { i->ReleaseVarDesc(p); }; };
static auto fn_tlibattr_dtor = [](ITypeLib* i) { return [&](TLIBATTR* p) { i->ReleaseTLibAttr(p); }; };

void ProcessTypeLibrary(ITypeLib* tpLib)
{
    unsigned int tiCount = tpLib->GetTypeInfoCount();
    std::wcout << "TypeInfo count is " << tiCount << std::endl;

    for (int i = 0; i < tiCount; i++)
    {
        CComPtr<ITypeInfo> pTypeInfo;

        if (SUCCEEDED(tpLib->GetTypeInfo(i, &pTypeInfo)))
        {
            TYPEATTR* _tpAttr;
            HREFTYPE refType;

            if (SUCCEEDED(pTypeInfo->GetTypeAttr(&_tpAttr)))
            {
                auto fn_1 = [&](TYPEATTR* p) { pTypeInfo->ReleaseTypeAttr(p); };
                std::unique_ptr<TYPEATTR, decltype(fn_1)> tpAttr{ _tpAttr, fn_1 };

                std::wcout << "TYPEKIND: " << GetTypeKind(tpAttr->typekind);
                std::wcout << "; GUID: " << GetGUIDStr(tpAttr->guid);
                std::wcout << "; cbSizeInstance: " << tpAttr->cbSizeInstance;
                std::wcout << "; cFuncs: " << tpAttr->cFuncs;
                std::wcout << "; cVars: " << tpAttr->cVars;
                std::wcout << "; cImplTypes: " << tpAttr->cImplTypes;
                std::wcout << "; cbSizeVft: " << tpAttr->cbSizeVft;
                std::wcout << std::endl;

                for (int funcIdx = 0; funcIdx < tpAttr->cFuncs; funcIdx++)
                {
                    FUNCDESC* _pFuncDesc;
                    if (SUCCEEDED(pTypeInfo->GetFuncDesc(funcIdx, &_pFuncDesc)))
                    {
                        auto fn_2 = [&](FUNCDESC* p) { pTypeInfo->ReleaseFuncDesc(p); };
                        std::unique_ptr<FUNCDESC, decltype(fn_2)> pFuncDesc{ _pFuncDesc, fn_2 };
                        BSTR _bstrName;

                        std::wcout << funcIdx << ": " << std::endl;
                        std::wcout << "  " << GetCallConv(pFuncDesc->callconv);

                        if (SUCCEEDED(pTypeInfo->GetDocumentation(pFuncDesc->memid, &_bstrName, nullptr, nullptr, nullptr)))
                        {
                            _bstr_t bstrName(_bstrName);
                            std::wcout << " " << (LPCWSTR)bstrName;
                            std::wcout << "(";

                            std::vector<BSTR> paramNames(20);

                            if (pFuncDesc->cParams > 0)
                            {
                                paramNames.assign(paramNames.size(), 0);

                                unsigned int pcNames;
                                throw_native_if_fail(pTypeInfo->GetNames(pFuncDesc->memid, paramNames.data(), paramNames.size(), &pcNames));

                                for (int iParamNum = 0; iParamNum < pFuncDesc->cParams; iParamNum++)
                                {
                                    _bstr_t paramName(paramNames[iParamNum + 1]);
                                    ELEMDESC _elem = pFuncDesc->lprgelemdescParam[iParamNum];

                                    if (iParamNum == 0)
                                    {
                                        std::wcout << GetVarType(_elem.tdesc.vt);
                                    }
                                    else
                                    {
                                        std::wcout << ", " << GetVarType(_elem.tdesc.vt);
                                    }
                                    std::wcout << " " << paramName;
                                }
                            }
                            std::wcout << ")";

                        }
                        std::wcout << "; funckind: " << GetFuncKind(pFuncDesc->funckind);
                        std::wcout << "; invokekind: " << GetInvokeKind(pFuncDesc->invkind);
                        std::wcout << "; oVft: " << pFuncDesc->oVft;
                        std::wcout << std::endl;
                    }
                }

                for (int varIdx = 0; varIdx < tpAttr->cVars; varIdx++)
                {
                    VARDESC* _pVarDesc;
                    if (SUCCEEDED(pTypeInfo->GetVarDesc(i, &_pVarDesc)))
                    {
                        auto fn_3 = [&](VARDESC* p) { pTypeInfo->ReleaseVarDesc(p); };
                        std::unique_ptr<VARDESC, decltype(fn_3)> pVarDesc{ _pVarDesc, fn_3 };
                        BSTR _bstrName;

                        std::wcout << "  -- memid: " << pVarDesc->memid;
                        if (SUCCEEDED(pTypeInfo->GetDocumentation(pVarDesc->memid, &_bstrName, nullptr, nullptr, nullptr)))
                        {
                            _bstr_t bstrName(_bstrName);
                            std::wcout << "; name: " << (LPCWSTR)bstrName;
                        }
                        std::wcout << "; varkind: " << GetVarKind(pVarDesc->varkind);
                        std::wcout << std::endl;
                    }
                }

                if (tpAttr->typekind == TYPEKIND::TKIND_DISPATCH)
                {
                    if (SUCCEEDED(pTypeInfo->GetRefTypeOfImplType(-1, &refType)))
                    {
                        CComPtr<ITypeInfo> pTypeInfo2;
                        pTypeInfo->GetRefTypeInfo(refType, &pTypeInfo2);

                        TYPEATTR* _tpAttr2;
                        if (SUCCEEDED(pTypeInfo2->GetTypeAttr(&_tpAttr2)))
                        {
                            auto fn_4 = [&](TYPEATTR* p) { pTypeInfo2->ReleaseTypeAttr(p); };
                            std::unique_ptr<TYPEATTR, decltype(fn_4)> tpAttr2{ _tpAttr2, fn_4 };

                            std::wcout << ">> TYPEKIND: " << GetTypeKind(tpAttr2->typekind);
                            std::wcout << "; GUID: " << GetGUIDStr(tpAttr2->guid);
                            std::wcout << "; cbSizeInstance: " << tpAttr2->cbSizeInstance;
                            std::wcout << "; cFuncs: " << tpAttr2->cFuncs;
                            std::wcout << "; cVars: " << tpAttr2->cVars;
                            std::wcout << "; cImplTypes: " << tpAttr2->cImplTypes;
                            std::wcout << "; cbSizeVft: " << tpAttr2->cbSizeVft;
                        }
                    }
                }
                std::wcout << std::endl;
            }
        }
    }
}

void ProcessMoniker(IBindCtx* ctx, IMoniker* mnk, IRunningObjectTable* objTbl)
{
    HRESULT hr;

    if (ctx && mnk && objTbl)
    {
        IUnknownPtr pIUnknown, pIUnknown2;
        IDispatchPtr pDispatch;

        throw_native_if_fail(objTbl->GetObject(mnk, &pIUnknown));

        // Redundant, but querying IUnknown on itself should always work
        bool has_IUnknown = true;
        hr = pIUnknown->QueryInterface(__uuidof(IUnknown), reinterpret_cast<void**>(&pIUnknown2));
        if (hr == E_NOINTERFACE)
        {
            has_IUnknown = false;
        }

        bool has_IDispatch = true;
        hr = pIUnknown->QueryInterface(__uuidof(IDispatch), reinterpret_cast<void**>(&pDispatch));
        if (hr == E_NOINTERFACE)
        {
            has_IDispatch = false;
        }

        //std::wcout << "Interfaces: IUnknown = " << (has_IUnknown ? "true" : "false") << "; IDispatch = " \
        //    << (has_IDispatch ? "true" : "false") << std::endl;

        if (!has_IDispatch)
        {
            return;
        }

        /* Type Information */
        CComPtr<ITypeInfo> pTypeInfo;
        throw_native_if_fail(pDispatch->GetTypeInfo(0, 0, &pTypeInfo));

        /* Containing Type Library information */
        CComPtr<ITypeLib> pTypeLibrary;
        unsigned int libIndex = 0;
        throw_native_if_fail(pTypeInfo->GetContainingTypeLib(&pTypeLibrary, &libIndex));

        TLIBATTR* _tLibAttr = nullptr;
        throw_native_if_fail(pTypeLibrary->GetLibAttr(&_tLibAttr));

        auto fn_1 = [&](TLIBATTR* p) { pTypeLibrary->ReleaseTLibAttr(p); };
        std::unique_ptr<TLIBATTR, decltype(fn_1)> pTypeLibraryAttr{ _tLibAttr, fn_1 };

        // We only keep processing once we've determined that we're dealing with Excel
        if (!IsEqualGUID(xl_typelib_guid, pTypeLibraryAttr->guid))
        {
            return;
        }

        LPOLESTR moniker_name;
        throw_native_if_fail(mnk->GetDisplayName(ctx, mnk, &moniker_name));
        std::unique_ptr < OLECHAR, decltype(&olestr_imalloc_free) > name(moniker_name, &olestr_imalloc_free);

        std::wstring sep(80, '$');
        std::wcout << sep.c_str() << std::endl;

        BSTR typeName;
        throw_native_if_fail(pTypeInfo->GetDocumentation(-1, &typeName, nullptr, nullptr, nullptr));
        _bstr_t bstrTypeName(typeName);

        TYPEATTR* _tAttr = nullptr;
        throw_native_if_fail(pTypeInfo->GetTypeAttr(&_tAttr));

        auto fn_2 = [&](TYPEATTR* p) { pTypeInfo->ReleaseTypeAttr(p); };
        std::unique_ptr<TYPEATTR, decltype(fn_2)> pTypeAttr{ _tAttr, fn_2 };

        std::wcout << "Type name: " << (LPCWSTR)bstrTypeName << "; GUID: " << GetGUIDStr(pTypeAttr->guid) << std::endl;

        ProcessTypeLibrary(pTypeLibrary.p);

        std::wstring wsName((LPCWSTR)bstrTypeName);
        if (wsName.find(L"Workbook") != std::wstring::npos)
        {
            CComPtr<IDispatch> spDisp(pDispatch);
            CComVariant fullName;
            hr = spDisp.GetPropertyByName(L"FullName", &fullName);
            if FAILED(hr)
            {
                if (hr == DISP_E_UNKNOWNNAME)
                {
                    std::wcout << L"Could not find property Workbook::FullName" << std::endl;
                }
            }
            else
            {
                std::wcout << L"Workbook full name is " << (LPCWSTR)fullName.bstrVal << std::endl;
            }

            CComVariant appVar;
            CComPtr<IDispatch> app;
            hr = spDisp.GetPropertyByName(L"Application", &appVar);
            app.Attach(appVar.pdispVal);
            CComVariant version;
            hr = app.GetPropertyByName(L"Version", &version);
            std::wcout << L"Application version is " << (LPCWSTR)version.bstrVal << std::endl;
        }
    }
}

int main()
{
    try
    {
        IBindCtxPtr ctx;
        IRunningObjectTablePtr table;
        IEnumMonikerPtr enumMnk;

        CLSID xl_app_clsid;
        CLSID xl_workbook_clsid;

        std::wstring xl_typelib_clsid_string(L"{00020813-0000-0000-C000-000000000046}");
        std::wstring xl_app_clsid_string(L"{00024500-0000-0000-C000-000000000046}");
        std::wstring xl_workbook_clsid_string(L"{00024500-0000-0000-C000-000000000046}");

        throw_native_if_fail(CLSIDFromString(xl_app_clsid_string.c_str(), &xl_app_clsid));
        throw_native_if_fail(CLSIDFromString(xl_workbook_clsid_string.c_str(), &xl_workbook_clsid));

        throw_native_if_fail(::CoInitialize(NULL));

        throw_native_if_fail(CreateBindCtx(NULL, &ctx));
        throw_native_if_fail(ctx->GetRunningObjectTable(&table));
        throw_native_if_fail(table->EnumRunning(&enumMnk));

        ULONG _mnk_count = 10;
        std::vector<IMonikerPtr> mnkArray(_mnk_count);
        ULONG pceltFetched(0);
        IMoniker** _monikers = &mnkArray[0];
        int mnkSoFar = 0;

        enumMnk->Next(_mnk_count, _monikers, &pceltFetched);

        while (pceltFetched > 0)
        {
            std::wcout << L"Retrieved " << pceltFetched << L" items and listing them." << std::endl;

            for (IMoniker* mnk : mnkArray)
            {
                if (mnk)
                {
                    //std::wcout << "Processing moniker " << ++mnkSoFar << std::endl;
                    ProcessMoniker((IBindCtx*)ctx, mnk, (IRunningObjectTable*)table);
                }
            }
            mnkArray.clear();
            mnkArray.resize(_mnk_count);
            pceltFetched = 0;

            enumMnk->Next(_mnk_count, _monikers, &pceltFetched);
        }

        ::CoUninitialize();
    }
    catch (_com_error & e)
    {
        std::wcout << "Error Message: " << e.ErrorMessage() << std::endl;
    }
}
