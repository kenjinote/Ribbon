#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "comctl32")

#include <windows.h>
#include "Ribbon.h"

TCHAR szClassName[] = TEXT("Window");

IUIFramework   *g_pFramework = NULL;
IUIApplication *g_pApplication = NULL;
HWND hWnd;

BOOL InitializeFramework(HWND hWnd)
{
	HRESULT hr = CoCreateInstance(
		CLSID_UIRibbonFramework,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&g_pFramework));
	if (FAILED(hr))
	{
		return FALSE;
	}
	hr = CApplication::CreateInstance(&g_pApplication);
	if (FAILED(hr))
	{
		return FALSE;
	}
	hr = g_pFramework->Initialize(hWnd, g_pApplication);
	if (FAILED(hr))
	{
		return FALSE;
	}
	hr = g_pFramework->LoadUI(
		GetModuleHandle(NULL),
		TEXT("APPLICATION_RIBBON"));
	if (FAILED(hr))
	{
		return FALSE;
	}
	return TRUE;
}

void DestroyFramework()
{
	if (g_pFramework)
	{
		g_pFramework->Destroy();
		g_pFramework->Release();
		g_pFramework = NULL;
	}
	if (g_pApplication)
	{
		g_pApplication->Release();
		g_pApplication = NULL;
	}
}

HRESULT CApplication::CreateInstance(IUIApplication **ppApplication)
{
	if (!ppApplication)
	{
		return E_POINTER;
	}
	*ppApplication = NULL;
	HRESULT hr = S_OK;
	CApplication* pApplication = new CApplication();
	if (pApplication != NULL)
	{
		*ppApplication = static_cast<IUIApplication *>(pApplication);
	}
	else
	{
		hr = E_OUTOFMEMORY;
	}
	return hr;
}

STDMETHODIMP_(ULONG) CApplication::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CApplication::Release()
{
	LONG cRef = InterlockedDecrement(&m_cRef);
	if (cRef == 0) {
		delete this;
	}
	return cRef;
}

STDMETHODIMP CApplication::QueryInterface(REFIID iid, void** ppv)
{
	if (iid == __uuidof(IUnknown))
	{
		*ppv = static_cast<IUnknown*>(this);
	}
	else if (iid == __uuidof(IUIApplication))
	{
		*ppv = static_cast<IUIApplication*>(this);
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP CApplication::OnCreateUICommand(
	UINT nCmdID,
	UI_COMMANDTYPE typeID,
	IUICommandHandler** ppCommandHandler)
{
	if (NULL == m_pCommandHandler)
	{
		HRESULT hr = CCommandHandler::CreateInstance(&m_pCommandHandler);
		if (FAILED(hr))
		{
			return hr;
		}
	}
	return m_pCommandHandler->QueryInterface(IID_PPV_ARGS(ppCommandHandler));
}

STDMETHODIMP CApplication::OnViewChanged(
	UINT viewId,
	UI_VIEWTYPE typeId,
	IUnknown* pView,
	UI_VIEWVERB verb,
	INT uReasonCode)
{
	HRESULT hr = E_NOTIMPL;
	if (UI_VIEWTYPE_RIBBON == typeId)
	{
		switch (verb)
		{
		case UI_VIEWVERB_CREATE:
			hr = S_OK;
			break;
		case UI_VIEWVERB_SIZE:
			{
				IUIRibbon* pRibbon = NULL;
				UINT uRibbonHeight;
				hr = pView->QueryInterface(IID_PPV_ARGS(&pRibbon));
				if (SUCCEEDED(hr))
				{
					hr = pRibbon->GetHeight(&uRibbonHeight);
					pRibbon->Release();
				}
			}
			break;
		case UI_VIEWVERB_DESTROY:
			hr = S_OK;
			break;
		}
	}
	return hr;
}

STDMETHODIMP CApplication::OnDestroyUICommand(
	UINT32 nCmdID,
	UI_COMMANDTYPE typeID,
	IUICommandHandler* commandHandler)
{
	return E_NOTIMPL;
}

HRESULT CCommandHandler::CreateInstance(IUICommandHandler **ppCommandHandler)
{
	if (!ppCommandHandler)
	{
		return E_POINTER;
	}
	*ppCommandHandler = NULL;
	HRESULT hr = S_OK;
	CCommandHandler* pCommandHandler = new CCommandHandler();
	if (pCommandHandler != NULL)
	{
		*ppCommandHandler = static_cast<IUICommandHandler *>(pCommandHandler);
	}
	else
	{
		hr = E_OUTOFMEMORY;
	}
	return hr;
}

STDMETHODIMP_(ULONG) CCommandHandler::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CCommandHandler::Release()
{
	LONG cRef = InterlockedDecrement(&m_cRef);
	if (cRef == 0)
	{
		delete this;
	}
	return cRef;
}

STDMETHODIMP CCommandHandler::QueryInterface(REFIID iid, void** ppv)
{
	if (iid == __uuidof(IUnknown))
	{
		*ppv = static_cast<IUnknown*>(this);
	}
	else if (iid == __uuidof(IUICommandHandler))
	{
		*ppv = static_cast<IUICommandHandler*>(this);
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	AddRef();
	return S_OK;
}

STDMETHODIMP CCommandHandler::UpdateProperty(UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* ppropvarCurrentValue, PROPVARIANT* ppropvarNewValue)
{
	return E_NOTIMPL;
}

STDMETHODIMP CCommandHandler::Execute(UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* ppropvarValue, IUISimplePropertySet* pCommandExecutionProperties)
{
	if (nCmdID == 2)
	{
		PostMessage(hWnd, WM_CLOSE, 0, 0);
	}
	return S_OK;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hButton;
	switch (msg)
	{
	case WM_CREATE:
		if (FAILED(CoInitialize(NULL)))
		{
			return -1;
		}
		if (!InitializeFramework(hWnd))
		{
			return -1;
		}
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("変換"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		break;
	case WM_SIZE:
		MoveWindow(hButton, 10, 300, 256, 32, TRUE);
		break;
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hdc = BeginPaint(hWnd, &ps);
			EndPaint(hWnd, &ps);
		}
		break;
	case WM_DESTROY:
		DestroyFramework();
		CoUninitialize();
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
	hWnd = CreateWindow(
		szClassName,
		TEXT("Window"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
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
