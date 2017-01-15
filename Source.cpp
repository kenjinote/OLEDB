#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#import "C:\Program Files\Common Files\System\ado\msado15.dll" no_namespace rename("EOF","EndOfFile")

TCHAR szClassName[] = TEXT("Window");

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;
	static HWND hButton;
	switch (msg)
	{
	case WM_CREATE:
		hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_HSCROLL | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("SQL文実行"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		break;
	case WM_SIZE:
		MoveWindow(hEdit, 10, 10, LOWORD(lParam) - 20, HIWORD(lParam) - 20 - 32 - 10, 1);
		MoveWindow(hButton, 10, HIWORD(lParam) - 32 - 10, 128, 32, 1);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			CoInitialize(NULL);
			try
			{
				_ConnectionPtr pCon(NULL);
				pCon.CreateInstance(__uuidof(Connection));
#ifdef _WIN64
				pCon->Open(TEXT("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=sample.mdb;"), _bstr_t(""), _bstr_t(""), adOpenUnspecified); // AccessDatabaseEngine_X64.exe のインストールが必要
#else
				pCon->Open(TEXT("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=sample.mdb;"), _bstr_t(""), _bstr_t(""), adOpenUnspecified);
#endif
				_CommandPtr pCommand = NULL;
				pCommand.CreateInstance(__uuidof(Command));
				pCommand->ActiveConnection = pCon;
				DWORD dwFileSize = GetWindowTextLength(hEdit);
				TCHAR*lpByte = (TCHAR*)GlobalAlloc(0, sizeof(TCHAR)*(dwFileSize + 1));
				GetWindowText(hEdit, lpByte, dwFileSize + 1);
				TCHAR*token = wcstok(lpByte, TEXT("\r\n"));
				while (token)
				{
					pCommand->CommandText = _bstr_t(token);
					pCommand->Execute(NULL, NULL, adCmdText);
					token = wcstok(NULL, TEXT("\r\n"));
				}
				GlobalFree(lpByte);
				pCon->Close();
			}
			catch (_com_error&e)
			{
				MessageBox(hWnd, e.Description(), 0, 0);
			}
			CoUninitialize();
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("OLEDBを使ったMDBファイルの更新"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
