// AppDLL.cpp : Определяет экспортированные функции для приложения DLL.
//

#include "stdafx.h"
#include "string.h"
#include "windows.h"
#include <iostream>
#include <ole2.h>
#include <io.h>
#include <comdef.h>

HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp,
	LPOLESTR ptName, int cArgs...)
{
	// Begin variable-argument list...
	va_list marker;
	va_start(marker, cArgs); //формируем список аргументов, передаваемых в cArgs

	if (!pDisp) {
		MessageBox(NULL, TEXT("NULL IDispatch passed to AutoWrap()"),
			TEXT("Error"), 0x10010);
		_exit(0);
	}
	// pDisp – это объект, на котoром выполняется метод, напр. document
	// Variables used...
	DISPPARAMS dp = { NULL, NULL, 0, 0 }; //структура определяет параметры,
										  // передаваемые в метод
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	char buf[200];
	char szName[200];

	// Convert down to ANSI
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// Get DISPID for name passed...
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT,
		&dispID);
	if (FAILED(hr)) {
		sprintf(buf,
			"IDispatch::GetIDsOfNames(\"%s\") failed w/err0x%08lx",
			szName, hr);
		MessageBox(NULL, TEXT("Click"), TEXT("AutoWrap()"), 0x10010);
		_exit(0);
		return hr;
	}

	// Allocate memory for arguments...
	VARIANT *pArgs = new VARIANT[cArgs + 1];

	// Extract arguments...
	for (int i = 0; i < cArgs; i++) {
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts!
	if (autoType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call!......................
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType,
		&dp, pvResult, NULL, NULL);
	//dispID = DISPATH_PROPERTYPUT

	if (FAILED(hr)) {
		sprintf(buf,
			"IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx",
			szName, dispID, hr);
		MessageBox(NULL, TEXT("Click Me"), TEXT("AutoWrap()"), 0x10010);
		_exit(0);
		return hr;
	}
	// End variable-argument section...
	va_end(marker);

	delete[] pArgs;

	return hr;

}



extern "C++" __declspec(dllexport) void OpenProgramm(wchar_t* docName)
{
	// Get CLSID for Word.Application...
	//получаем в clsid ссылку на тип word.application 
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Word.Application", &clsid);
	if (FAILED(hr)) {
		::MessageBox(NULL, TEXT("CLSIDFromProgID() failed"), TEXT("Error"),
			0x10010);
		return;
	}

	// Start Word and get IDispatch...
	//Если Word зарегистрирован в системе, то следующая команда создает объект Word Dispatch *pWordApp

	IDispatch *pWordApp;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER,
		IID_IDispatch, (void **)&pWordApp);
	if (FAILED(hr)) {
		::MessageBox(NULL, TEXT("Word not registered properly"),
			TEXT("Error"), 0x10010);
		return;
	}if (FAILED(hr)) {
		::MessageBox(NULL, TEXT("Word not registered properly"),
			TEXT("Error"), 0x10010);
		return;
	}
	_bstr_t b(docName);
	const char* c = b;
	_finddata_t data;
	long nFind = _findfirst(c, &data);
	if (nFind != -1)
	{
		_findclose(nFind);

		// Make Word visible
		//формиуем и выполняем команды на созданном компоненте
		{
			VARIANT x;
			x.vt = VT_I4;
			x.lVal = 1;
			//DISPATCH_PROPERTYGET означает, что значение параметра COM-объекта считывается
			AutoWrap(DISPATCH_PROPERTYPUT, NULL, pWordApp, L"Visible", 1,
				x);//выдает метод по invoke и передает параметы в метод
		}

		// Get Documents collection
		IDispatch *pDocs;
		{
			VARIANT result;
			VariantInit(&result);
			AutoWrap(DISPATCH_PROPERTYGET, &result, pWordApp, L"Documents",
				0);

			pDocs = result.pdispVal;
		}

		// Call Documents.Open() to open D:\laba.doc
		IDispatch *pDoc;
		{
			VARIANT result;
			VariantInit(&result);
			VARIANT x;
			x.vt = VT_BSTR;
			x.bstrVal = ::SysAllocString(docName);

			AutoWrap(DISPATCH_METHOD, &result, pDocs, L"Open", 1, x);
			pDoc = result.pdispVal;
			SysFreeString(x.bstrVal);
		}

		::MessageBox(NULL,
			TEXT("Click Me"),
			TEXT("word doc"), 0x10000);
		pDoc->Release();
		pDocs->Release();
		pWordApp->Release();

	}
}


