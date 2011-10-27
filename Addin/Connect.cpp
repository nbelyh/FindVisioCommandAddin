// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn_i.h"
#include "Connect.h"

// When run, the Add-in wizard prepared the registry for the Add-in.
// At a later time, if the Add-in becomes unavailable for reasons such as:
//   1) You moved this project to a computer other than which is was originally created on.
//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
//   3) Registry corruption.
// you will need to re-register the Add-in by building the MyAddin1Setup project, 
// right click the project in the Solution Explorer, then choose install.

CString LoadTextFile(UINT resource_id)
{
	HMODULE hResources = AfxGetResourceHandle();

	HRSRC rc = ::FindResource(
		hResources, MAKEINTRESOURCE(resource_id), L"TEXTFILE");

	LPSTR content = static_cast<LPSTR>(
		::LockResource(::LoadResource(hResources, rc)));

	DWORD content_length = 
		::SizeofResource(hResources, rc);

	return CString(content, content_length);
}

// CConnect
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** custom)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
	pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pAddInInstance);

	CString string_map = LoadTextFile(IDR_RIBBON_MAP);

	LPCWSTR begin = string_map;
	LPCWSTR end = string_map;

	while (*end)
	{
		begin = end;
		while (*end && *end != '\t') ++end;

		CString control_id = CString(begin, end - begin);

		begin = end;
		while (*end && *end != '\n') ++end;

		CString control_name = CString(begin, end - begin);

		if (!control_id.IsEmpty() && !control_name.IsEmpty())
			m_ribbon_map.Add(control_id, control_name);

		++end;
	}

	LoadHistory();

	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	SaveHistory();

	m_pApplication = NULL;
	m_pAddInInstance = NULL;
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return S_OK;
}

STDMETHODIMP CConnect::GetCustomUI(BSTR RibbonID, BSTR * RibbonString)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	*RibbonString = LoadTextFile(IDR_RIBBON).AllocSysString();

	return S_OK;
}

STDMETHODIMP CConnect::ButtonClicked (IDispatch * RibbonControl)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	MessageBoxW(NULL,L"The button was clicked!", L"Message from ExampleATLAddIn", MB_OK | MB_ICONINFORMATION); 
	return S_OK;
}

STDMETHODIMP CConnect::IsRibbonButtonVisible(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	if (RibbonControl == NULL) return E_POINTER;
	if (pResult == NULL) return E_POINTER;

	if (m_text.IsEmpty())
		return S_OK;

	IRibbonControlPtr control;
	RibbonControl->QueryInterface(__uuidof(IRibbonControl), (void**)&control);

	BSTR bstr_control_id = NULL;
	if (FAILED(control->get_Id(&bstr_control_id)))
		return S_OK;

	CString control_id = bstr_control_id;
	SysFreeString(bstr_control_id);

	int idx = m_ribbon_map.FindKey(control_id);
	if (idx < 0)
		return S_OK;

	CString name = m_ribbon_map.GetValueAt(idx);

	if (StrStrIW(name, m_text))
		*pResult = VARIANT_TRUE;

	return S_OK;
}

#define MAX_HISTORY	5

void CConnect::AddToHistory(const CString& key)
{
	for (int i = m_history.GetSize()-1; i >= 0; --i)
	{
		if (!StrCmpI(m_history[i], m_text))
		{
			for (int j = i; j < m_history.GetSize()-1; ++j)
				m_history[j] = m_history[j+1];
			m_history[m_history.GetSize()-1] = m_text;
			return;
		}
	}

	if (m_history.GetSize() == MAX_HISTORY)
	{
		for (int i = 0; i < m_history.GetSize()-1; ++i)
			m_history[i] = m_history[i+1];
		m_history[m_history.GetSize()-1] = m_text;
		return;
	}

	m_history.Add(m_text);
}

#define REG_KEY	L"Software\\UnmanagedVisio\\FindVisioCommandAddin"
#define REG_VAL	L"History"
#define REG_SEP	L"|"

void CConnect::LoadHistory ()
{
	CRegKey key;
	if (key.Open(HKEY_CURRENT_USER, REG_KEY, KEY_READ) != 0)
		return;

	CString val;
	ULONG val_len = 1024;
	key.QueryStringValue(REG_VAL, val.GetBuffer(val_len), &val_len);
	val.ReleaseBuffer(val_len);

	int fragment_begin = 0;
	while (fragment_begin < val.GetLength())
	{
		int fragment_end = val.Find(REG_SEP, fragment_begin);
		if (fragment_end < 0)
			fragment_end = val.GetLength();

		CString tok = 
			val.Mid(fragment_begin, fragment_end - fragment_begin);

		m_history.Add(tok);

		fragment_begin = fragment_end + wcslen(REG_SEP);
	}

	key.Close();
}

void CConnect::SaveHistory ()
{
	CRegKey key;
	if (key.Create(HKEY_CURRENT_USER, REG_KEY) != 0)
		return;

	CString val;

	int i = 0; 
	while (i < m_history.GetSize())
	{
		val += m_history[i];

		if (i < m_history.GetSize()-1)
			val += REG_SEP;

		++i;
	}


	key.SetStringValue(REG_VAL, val);
	key.Close();
}

STDMETHODIMP CConnect::OnRibbonComboBoxChange(IDispatch * RibbonControl, BSTR *pbstrText)
{
	if (RibbonControl == NULL) return E_POINTER;
	if (pbstrText == NULL) return E_POINTER;

	IRibbonUIPtr ui;
	m_pRibbon->QueryInterface(__uuidof(IRibbonUI), (void**)&ui);

	m_text = *pbstrText;

	AddToHistory(m_text);
	ui->Invalidate();

	return S_OK;
}

STDMETHODIMP CConnect::GetRibbonComboBoxText(IDispatch*pControl, BSTR *pbstrText)
{
	*pbstrText = m_text.AllocSysString();
	return S_OK;
}

STDMETHODIMP CConnect::OnRibbonLoad(IDispatch* disp)
{
	m_pRibbon = disp;
	return S_OK;
}


STDMETHODIMP CConnect::GetRibbonComboBoxItemCount(IDispatch*pControl, long *count)
{
	*count = m_history.GetSize();
	return S_OK;
}

STDMETHODIMP CConnect::GetRibbonComboBoxItemLabel(IDispatch*pControl, long cIndex, BSTR *pbstrLabel)
{
	*pbstrLabel = m_history[m_history.GetSize()-cIndex-1].AllocSysString();
	return S_OK;
}
