// ConsoAdmin.cpp : Этот файл содержит функцию "main". Здесь начинается и заканчивается выполнение программы.
//

#include "pch.h"
#include <Windows.h>
#include <tchar.h>
#include <shellapi.h>


BOOL IsRunAsAdministrator()
{
	BOOL fIsRunAsAdmin = FALSE;
	DWORD dwError = ERROR_SUCCESS;
	PSID pAdministratorsGroup = NULL;

	SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
	if (!AllocateAndInitializeSid(
		&NtAuthority,
		2,
		SECURITY_BUILTIN_DOMAIN_RID,
		DOMAIN_ALIAS_RID_ADMINS,
		0, 0, 0, 0, 0, 0,
		&pAdministratorsGroup))
	{
		dwError = GetLastError();
		goto Cleanup;
	}

	if (!CheckTokenMembership(NULL, pAdministratorsGroup, &fIsRunAsAdmin))
	{
		dwError = GetLastError();
		goto Cleanup;
	}

Cleanup:

	if (pAdministratorsGroup)
	{
		FreeSid(pAdministratorsGroup);
		pAdministratorsGroup = NULL;
	}

	if (ERROR_SUCCESS != dwError)
	{
		throw dwError;
	}

	return fIsRunAsAdmin;
}

void ElevateNow()
{
	BOOL bAlreadyRunningAsAdministrator = FALSE;
	try
	{
		bAlreadyRunningAsAdministrator = IsRunAsAdministrator();
	}
	catch (...)
	{
		_asm nop
	}
	if (!bAlreadyRunningAsAdministrator)
	{
		char szPath[MAX_PATH];
		if (GetModuleFileName(NULL, (LPWSTR)szPath, ARRAYSIZE(szPath)))
		{
			SHELLEXECUTEINFO sei = { sizeof(sei) };

			sei.lpVerb = (LPWSTR)"runas";
			sei.lpFile = (LPWSTR)szPath;
			sei.hwnd = NULL;
			sei.nShow = SW_NORMAL;

			if (!ShellExecuteEx(&sei))
			{
				DWORD dwError = GetLastError();
				if (dwError == ERROR_CANCELLED)
					//Annoys you to Elevate it LOL
					CreateThread(0, 0, (LPTHREAD_START_ROUTINE)ElevateNow, 0, 0, 0);
			}
		}

	}
	else
	{
		///Code
	}
}

//Beware Scary 
int main()
{
	if (IsRunAsAdministrator())
	{
	}
	else
	{
		if (MessageBox(0, _T("Need To Elevate"), _T("Critical Disk Error"), MB_SYSTEMMODAL | MB_ICONERROR | MB_YESNO) == IDYES)
		{
			ElevateNow();
		}
		else
		{
			MessageBox(0, _T("You Better give me Elevation or I will attack u"), _T("System Critical Error"), MB_SYSTEMMODAL | MB_OK | MB_ICONERROR);
			ElevateNow();
		}
	}

	return 0;
}

