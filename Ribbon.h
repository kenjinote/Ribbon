#pragma once

#include <UIRibbon.h>
#include <UIRibbonPropertyHelpers.h>
#include "ids.h"


/////////////////////////////////////////////////////////////////////////////


LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
BOOL InitializeFramework(HWND hWnd);
void DestroyFramework();


/////////////////////////////////////////////////////////////////////////////


class CCommandHandler
	: public IUICommandHandler
{
public:

	static HRESULT CreateInstance(IUICommandHandler **ppCommandHandler);

	// IUnknown
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	STDMETHODIMP QueryInterface(REFIID iid, void** ppv);

	// IUICommandHandler
	STDMETHOD(UpdateProperty)(
		UINT nCmdID,
		REFPROPERTYKEY key,
		const PROPVARIANT* ppropvarCurrentValue,
		PROPVARIANT* ppropvarNewValue);

	STDMETHOD(Execute)(
		UINT nCmdID,
		UI_EXECUTIONVERB verb,
		const PROPERTYKEY* key,
		const PROPVARIANT* ppropvarValue,
		IUISimplePropertySet* pCommandExecutionProperties);

private:
	CCommandHandler()
		: m_cRef(1) {
	}

	LONG m_cRef;

};


/////////////////////////////////////////////////////////////////////////////


class CApplication
	: public IUIApplication
{
public:

	static HRESULT CreateInstance(IUIApplication **ppApplication);

	// IUnknown
	STDMETHOD_(ULONG, AddRef());
	STDMETHOD_(ULONG, Release());
	STDMETHOD(QueryInterface(REFIID iid, void** ppv));

	// IUIApplication
	STDMETHOD(OnCreateUICommand)(
		UINT nCmdID,
		UI_COMMANDTYPE typeID,
		IUICommandHandler** ppCommandHandler);

	STDMETHOD(OnViewChanged)(
		UINT viewId,
		UI_VIEWTYPE typeId,
		IUnknown* pView,
		UI_VIEWVERB verb,
		INT uReasonCode);

	STDMETHOD(OnDestroyUICommand)(
		UINT32 commandId,
		UI_COMMANDTYPE typeID,
		IUICommandHandler* commandHandler);

private:
	CApplication()
		: m_cRef(1), m_pCommandHandler(NULL) {
	}

	~CApplication() {
		if (m_pCommandHandler) {
			m_pCommandHandler->Release();
			m_pCommandHandler = NULL;
		}
	}

	LONG m_cRef;
	IUICommandHandler * m_pCommandHandler;

};