// ActiveXCPP3Ctrl.cpp: реализация класса элемента управления ActiveX CActiveXCPP3Ctrl.

#include "stdafx.h"
#include "ActiveXCPP3.h"
#include "ActiveXCPP3Ctrl.h"
#include "ActiveXCPP3PropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CActiveXCPP3Ctrl, COleControl)

// Схема сообщений

BEGIN_MESSAGE_MAP(CActiveXCPP3Ctrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// Схема подготовки к отправке

BEGIN_DISPATCH_MAP(CActiveXCPP3Ctrl, COleControl)
	DISP_FUNCTION_ID(CActiveXCPP3Ctrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)

	DISP_FUNCTION_ID(CActiveXCPP3Ctrl, "Ellipse", DISPID_TEST_METHOD, Ellipse, VT_R8, VTS_R8 VTS_R8)
END_DISPATCH_MAP()

// Схема событий

BEGIN_EVENT_MAP(CActiveXCPP3Ctrl, COleControl)
END_EVENT_MAP()

// Страницы свойств

// TODO: при необходимости добавьте дополнительные страницы свойств.  Не забудьте увеличить значение счетчика.
BEGIN_PROPPAGEIDS(CActiveXCPP3Ctrl, 1)
	PROPPAGEID(CActiveXCPP3PropPage::guid)
END_PROPPAGEIDS(CActiveXCPP3Ctrl)

// Инициализировать фабрику класса и guid

IMPLEMENT_OLECREATE_EX(CActiveXCPP3Ctrl, "MFCACTIVEXCONTRO.ActiveXCPP3Ctrl.1",
	0xe8fad06c,0xd6bd,0x4e2a,0xb9,0xdc,0xea,0x11,0xf5,0x8f,0x8a,0xaa)

// Введите ИД и версию библиотеки

IMPLEMENT_OLETYPELIB(CActiveXCPP3Ctrl, _tlid, _wVerMajor, _wVerMinor)

// ИД интерфейса

const IID IID_DActiveXCPP3 = {0xc8bdc564,0x0a0e,0x4d2b,{0xac,0xcc,0xe7,0x3d,0x2d,0x91,0x06,0x56}};
const IID IID_DActiveXCPP3Events = {0x6b497838,0xa8b4,0x4940,{0x9e,0x15,0x2c,0x23,0xe3,0x85,0x29,0x67}};

// Сведения о типах элементов управления

static const DWORD _dwActiveXCPP3OleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CActiveXCPP3Ctrl, IDS_ACTIVEXCPP3, _dwActiveXCPP3OleMisc)

// CActiveXCPP3Ctrl::CActiveXCPP3CtrlFactory::UpdateRegistry -
// Добавляет или удаляет записи системного реестра для CActiveXCPP3Ctrl

BOOL CActiveXCPP3Ctrl::CActiveXCPP3CtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: убедитесь, что элементы управления следуют правилам модели изолированных потоков.
	// Дополнительные сведения см. в MFC TechNote 64.
	// Если элемент управления не соответствует правилам модели изоляции, то
	// необходимо модифицировать приведенный ниже код, изменив значение 6-го параметра с
	// afxRegApartmentThreading на 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_ACTIVEXCPP3,
			IDB_ACTIVEXCPP3,
			afxRegApartmentThreading,
			_dwActiveXCPP3OleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


// CActiveXCPP3Ctrl::CActiveXCPP3Ctrl — конструктор

CActiveXCPP3Ctrl::CActiveXCPP3Ctrl()
{
	InitializeIIDs(&IID_DActiveXCPP3, &IID_DActiveXCPP3Events);
	// TODO: Инициализируйте здесь данные экземпляра элемента управления.
}

// CActiveXCPP3Ctrl::~CActiveXCPP3Ctrl — деструктор

CActiveXCPP3Ctrl::~CActiveXCPP3Ctrl()
{
	// TODO: Выполните здесь очистку данных экземпляра элемента управления.
}

// CActiveXCPP3Ctrl::OnDraw — функция рисования

void CActiveXCPP3Ctrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;

	// TODO: Замените следующий код собственным кодом рисования.
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}

// CActiveXCPP3Ctrl::DoPropExchange — поддержка сохраняемости

void CActiveXCPP3Ctrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Вызывать функции PX_ для каждого постоянного настраиваемого свойства.
}


// CActiveXCPP3Ctrl::OnResetState — сброс элемента управления к состоянию по умолчанию

void CActiveXCPP3Ctrl::OnResetState()
{
	COleControl::OnResetState();  // Сбрасывает значения по умолчанию, найденные в DoPropExchange

	// TODO: Сбросьте здесь состояние любого другого элемента управления.
}


// CActiveXCPP3Ctrl::AboutBox — отображение окна "О программе" для пользователя

void CActiveXCPP3Ctrl::AboutBox()
{
	CDialogEx dlgAbout(IDD_ABOUTBOX_ACTIVEXCPP3);
	dlgAbout.DoModal();
}

DOUBLE CActiveXCPP3Ctrl::Ellipse(DOUBLE param1, DOUBLE param2)

{

	double q = 3.14 * param1 * param2;


	return q;

}
// Обработчики сообщений CActiveXCPP3Ctrl
