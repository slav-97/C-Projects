// ActiveXCPP3.cpp: реализация CActiveXCPP3App и регистрация библиотеки DLL.

#include "stdafx.h"
#include "ActiveXCPP3.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CActiveXCPP3App theApp;

const GUID CDECL _tlid = {0x8971dbf5,0x18f3,0x433e,{0x88,0x5a,0xa0,0x7c,0x15,0x88,0x53,0x12}};
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;



// CActiveXCPP3App::InitInstance — инициализация библиотеки DLL

BOOL CActiveXCPP3App::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO: Добавьте здесь свой код инициализации модуля.
	}

	return bInit;
}



// CActiveXCPP3App::ExitInstance — завершение библиотеки DLL

int CActiveXCPP3App::ExitInstance()
{
	// TODO: Добавьте здесь свой код завершения работы модуля.

	return COleControlModule::ExitInstance();
}



// DllRegisterServer - добавляет записи в системный реестр

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}



// DllUnregisterServer - удаляет записи из системного реестра

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
