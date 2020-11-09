#pragma once

// ActiveXCPP3.h: основной файл заголовка для ActiveXCPP3.DLL

#if !defined( __AFXCTL_H__ )
#error "включить afxctl.h до включения этого файла"
#endif

#include "resource.h"       // основные символы


// CActiveXCPP3App: реализацию см. в файле ActiveXCPP3.cpp.

class CActiveXCPP3App : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

