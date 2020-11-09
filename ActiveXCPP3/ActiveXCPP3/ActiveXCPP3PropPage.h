#pragma once

// ActiveXCPP3PropPage.h: объявление класса страницы свойств CActiveXCPP3PropPage.


// CActiveXCPP3PropPage: реализацию см. в ActiveXCPP3PropPage.cpp.

class CActiveXCPP3PropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CActiveXCPP3PropPage)
	DECLARE_OLECREATE_EX(CActiveXCPP3PropPage)

// Конструктор
public:
	CActiveXCPP3PropPage();

// Данные диалогового окна
	enum { IDD = IDD_PROPPAGE_ACTIVEXCPP3 };

// Реализация
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // поддержка DDX/DDV

// Схемы сообщений
protected:
	DECLARE_MESSAGE_MAP()
};

