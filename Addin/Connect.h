// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols

using namespace Office;
using namespace AddInDesignerObjects;

// CConnect
class ATL_NO_VTABLE CConnect : 
	public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CConnect, &CLSID_Connect>
	, public IDispatchImpl<ICallbackInterface, &__uuidof(ICallbackInterface), &LIBID_AddinLib, 1, 0>
	, public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), 1, 0>
	, public IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), 12, 0>
{
public:
	CConnect()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
DECLARE_NOT_AGGREGATABLE(CConnect)

BEGIN_COM_MAP(CConnect)
	COM_INTERFACE_ENTRY2(IDispatch, ICallbackInterface)
	COM_INTERFACE_ENTRY(IDTExtensibility2)
	COM_INTERFACE_ENTRY(IRibbonExtensibility)
	COM_INTERFACE_ENTRY(ICallbackInterface)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

	//IRibbonExtensibility implementation:
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonString);

	STDMETHOD(ButtonClicked)(IDispatch * RibbonControl);
	STDMETHOD(IsRibbonButtonVisible)(IDispatch * RibbonControl, VARIANT_BOOL* pResult);
	STDMETHOD(OnRibbonComboBoxChange)(IDispatch * RibbonControl, BSTR *pbstrText);

	STDMETHOD(GetRibbonComboBoxItemCount)(IDispatch*pControl, long *count);
	STDMETHOD(GetRibbonComboBoxText)(IDispatch*pControl, BSTR *pbstrText);
	STDMETHOD(GetRibbonComboBoxItemLabel)(IDispatch*pControl, long cIndex, BSTR *pbstrLabel);

	STDMETHOD(OnRibbonLoad)(IDispatch* disp);

	CComPtr<IDispatch> m_pApplication;
	CComPtr<IDispatch> m_pAddInInstance;
	CComPtr<IDispatch> m_pRibbon;

	typedef CSimpleMap<CString, CString> StringMap;
	StringMap m_ribbon_map;

	void LoadHistory();
	void SaveHistory();
	void AddToHistory(const CString& key);


	CSimpleArray<CString> m_history;

	CString m_text;
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
