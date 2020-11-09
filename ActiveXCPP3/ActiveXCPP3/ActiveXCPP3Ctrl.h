#pragma once
#define DISPID_TEST_METHOD (1025317) //ДОБАВЛЕНО МНОЙ
// ActiveXCPP3Ctrl.h: объявление класса элемента управления ActiveX CActiveXCPP3Ctrl.


// CActiveXCPP3Ctrl: реализацию см. в ActiveXCPP3Ctrl.cpp.

class CActiveXCPP3Ctrl : public COleControl
{
	DECLARE_DYNCREATE(CActiveXCPP3Ctrl)

// Конструктор
public:
	CActiveXCPP3Ctrl();

// Переопределение
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// Реализация
protected:
	~CActiveXCPP3Ctrl();

	DECLARE_OLECREATE_EX(CActiveXCPP3Ctrl)    // Фабрика класса и guid
	DECLARE_OLETYPELIB(CActiveXCPP3Ctrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CActiveXCPP3Ctrl)     // ИД страницы свойств
	DECLARE_OLECTLTYPE(CActiveXCPP3Ctrl)		// Введите имя и промежуточное состояние

// Схемы сообщений
	DECLARE_MESSAGE_MAP()

// Схемы подготовки к отправке
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

	afx_msg double Ellipse(double param1, double param2);//МНОЙ

// Схемы событий
	DECLARE_EVENT_MAP()

// Подготовка к отправке и ИД событий
public: //Эти строки может создавать мастер ClassView
	enum {
		eventidEventHandlerTest = 1L
	};
protected:
	void EventHandlerTest(LPCTSTR name, LPCTSTR surname, LONG age)
	{
		FireEvent(eventidEventHandlerTest, EVENT_PARAM(VTS_BSTR VTS_BSTR VTS_I4), name, surname, age);
	} //ДОБАВЛЕНО МНОЙ
};

