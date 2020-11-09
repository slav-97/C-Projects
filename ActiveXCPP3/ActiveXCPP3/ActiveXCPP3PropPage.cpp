// ActiveXCPP3PropPage.cpp: реализация класса страницы свойств CActiveXCPP3PropPage.

#include "stdafx.h"
#include "ActiveXCPP3.h"
#include "ActiveXCPP3PropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CActiveXCPP3PropPage, COlePropertyPage)

// Схема сообщений

BEGIN_MESSAGE_MAP(CActiveXCPP3PropPage, COlePropertyPage)
END_MESSAGE_MAP()

// Инициализировать фабрика класса и guid

IMPLEMENT_OLECREATE_EX(CActiveXCPP3PropPage, "MFCACTIVEXCONT.ActiveXCPP3PropPage.1",
	0x4288ea3f,0x6072,0x4f3b,0x90,0xb6,0x41,0xb6,0x4d,0xde,0xf2,0x36)

// CActiveXCPP3PropPage::CActiveXCPP3PropPageFactory::UpdateRegistry -
// Добавляет или удаляет записи системного реестра для CActiveXCPP3PropPage

BOOL CActiveXCPP3PropPage::CActiveXCPP3PropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_ACTIVEXCPP3_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, nullptr);
}

// CActiveXCPP3PropPage::CActiveXCPP3PropPage — конструктор

CActiveXCPP3PropPage::CActiveXCPP3PropPage() :
	COlePropertyPage(IDD, IDS_ACTIVEXCPP3_PPG_CAPTION)
{
}

// CActiveXCPP3PropPage::DoDataExchange — перемещение данных между страницей и свойствами

void CActiveXCPP3PropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}

// Обработчики сообщений CActiveXCPP3PropPage
