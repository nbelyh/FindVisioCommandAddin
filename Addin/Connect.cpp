// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn_i.h"
#include "Connect.h"

namespace {

	#define DEFAULT_LANGUAGE 1033

	int GetAppLanguage(CComPtr<IDispatch> app)
	{
		CComDispatchDriver disp = app;

		variant_t v_language_settings;
		if (FAILED(disp.GetPropertyByName(L"LanguageSettings", &v_language_settings)))
			return DEFAULT_LANGUAGE;

		if (V_VT(&v_language_settings) != VT_DISPATCH)
			return DEFAULT_LANGUAGE;

		LanguageSettingsPtr language_settings;
		if (FAILED(V_DISPATCH(&v_language_settings)->QueryInterface(__uuidof(LanguageSettings), (void**)&language_settings)))
			return DEFAULT_LANGUAGE;

		int app_language = 0;
		if (FAILED(language_settings->get_LanguageID(msoLanguageIDUI, &app_language)))
			return DEFAULT_LANGUAGE;

		switch (app_language)
		{
		case 1033:
		case 1031:
		case 1049:
		case 1041:
			return app_language;
		}

		return DEFAULT_LANGUAGE;
	}

	struct LanguageLock
	{
		int old_lcid;
		int old_langid;

		LanguageLock(int app_language)
		{
			HMODULE hKernel32 = GetModuleHandle(L"Kernel32.dll");

			typedef LANGID (WINAPI *FP_SetThreadUILanguage)(LANGID LangId);
			FP_SetThreadUILanguage pSetThreadUILanguage = (FP_SetThreadUILanguage)GetProcAddress(hKernel32, "SetThreadUILanguage");

			typedef LANGID (WINAPI *FP_GetThreadUILanguage)();
			FP_GetThreadUILanguage pGetThreadUILanguage = (FP_GetThreadUILanguage)GetProcAddress(hKernel32, "GetThreadUILanguage");

			old_lcid = GetThreadLocale();
			SetThreadLocale(app_language);

			old_langid = 0;
			if (pSetThreadUILanguage && pGetThreadUILanguage)
			{
				old_langid = pGetThreadUILanguage();
				pSetThreadUILanguage(app_language);
			}
		}

		~LanguageLock()
		{
			SetThreadLocale(old_lcid);

			HMODULE hKernel32 = GetModuleHandle(L"Kernel32.dll");

			typedef LANGID (WINAPI *FP_SetThreadUILanguage)(LANGID LangId);
			FP_SetThreadUILanguage pSetThreadUILanguage = (FP_SetThreadUILanguage)GetProcAddress(hKernel32, "SetThreadUILanguage");

			typedef LANGID (WINAPI *FP_GetThreadUILanguage)();
			FP_GetThreadUILanguage pGetThreadUILanguage = (FP_GetThreadUILanguage)GetProcAddress(hKernel32, "GetThreadUILanguage");

			if (pSetThreadUILanguage && pGetThreadUILanguage)
				pSetThreadUILanguage(old_langid);
		}
	};

	CString LoadTextFile(UINT resource_id)
	{
		HMODULE hResources = AfxGetResourceHandle();

		HRSRC rc = ::FindResource(
			hResources, MAKEINTRESOURCE(resource_id), L"TEXTFILE");

		LPWSTR content = static_cast<LPWSTR>(
			::LockResource(::LoadResource(hResources, rc)));

		DWORD content_length = 
			::SizeofResource(hResources, rc);

		return CString(content, content_length / 2);
	}

	int GetLanguageIdFromControlId(const CString& control_id)
	{
		if (control_id == L"RibbonLanguageEN")
			return 1033;
		if (control_id == L"RibbonLanguageRU")
			return 1049;
		if (control_id == L"RibbonLanguageDE")
			return 1031;
		if (control_id == L"RibbonLanguageJP")
			return 1041;

		return 1033;
	}
}

#define REG_KEY	L"Software\\UnmanagedVisio\\FindVisioCommandAddin"
#define REG_VAL_HISTORY	L"History"
#define REG_VAL_LANGUAGES	L"Languages"
#define REG_SEP	L"|"

void CVisioConnect::LoadHistory ()
{
	CRegKey key;
	if (key.Open(HKEY_CURRENT_USER, REG_KEY, KEY_READ) != 0)
		return;

	CString val_history;
	ULONG val_history_len = 1024;
	key.QueryStringValue(REG_VAL_HISTORY, val_history.GetBuffer(val_history_len), &val_history_len);
	val_history.ReleaseBuffer(val_history_len);

	int fragment_begin = 0;
	while (fragment_begin < val_history.GetLength())
	{
		int fragment_end = val_history.Find(REG_SEP, fragment_begin);
		if (fragment_end < 0)
			fragment_end = val_history.GetLength();

		CString tok = 
			val_history.Mid(fragment_begin, fragment_end - fragment_begin);

		m_history.Add(tok);

		fragment_begin = fragment_end + wcslen(REG_SEP);
	}

	DWORD language = 0;
	if (0 == key.QueryDWORDValue(REG_VAL_LANGUAGES, language))
		m_language = language;

	key.Close();
}

void CVisioConnect::SaveHistory ()
{
	CRegKey key;
	if (key.Create(HKEY_CURRENT_USER, REG_KEY) != 0)
		return;

	CString val_history;

	int i = 0; 
	while (i < m_history.GetSize())
	{
		val_history += m_history[i];

		if (i < m_history.GetSize()-1)
			val_history += REG_SEP;

		++i;
	}


	key.SetStringValue(REG_VAL_HISTORY, val_history);

	key.SetDWORDValue(REG_VAL_LANGUAGES, m_language);

	key.Close();
}

// CConnect
STDMETHODIMP CVisioConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** custom)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pApplication);
	pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_pAddInInstance);

	m_language = GetAppLanguage(m_pApplication);

	LoadHistory();
	LoadStringMap();

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	SaveHistory();

	m_pApplication = NULL;
	m_pAddInInstance = NULL;
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return S_OK;
}

STDMETHODIMP CVisioConnect::GetCustomUI(BSTR RibbonID, BSTR * RibbonString)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	*RibbonString = LoadTextFile(IDR_RIBBON).AllocSysString();

	return S_OK;
}

namespace {

	CString GetControlId(IDispatch* pControl)
	{
		IRibbonControlPtr control;
		pControl->QueryInterface(__uuidof(IRibbonControl), (void**)&control);

		CComBSTR bstr_control_id;
		if (FAILED(control->get_Id(&bstr_control_id)))
			return S_OK;

		return static_cast<LPCWSTR>(bstr_control_id);
	}
}

STDMETHODIMP CVisioConnect::OnRibbonButtonClicked (IDispatch * pControl)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString control_id = GetControlId(pControl);

	if (control_id == L"FindButton")
		OnFind();

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnLanguageCheck(IDispatch *pControl, VARIANT_BOOL *pvarfPressed)
{
	CString control_id = GetControlId(pControl);

	if (*pvarfPressed)
	{
		m_language = GetLanguageIdFromControlId(control_id);

		LoadStringMap();
		InvalidateRibbon();
	}

	return S_OK;
}

STDMETHODIMP CVisioConnect::IsLanguageChecked(IDispatch * pControl, VARIANT_BOOL* pResult)
{
	CString control_id = GetControlId(pControl);

	*pResult = (m_language == GetLanguageIdFromControlId(control_id))
		? VARIANT_TRUE : VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CVisioConnect::IsRibbonButtonVisible(IDispatch * pControl, VARIANT_BOOL* pResult)
{
	if (pControl == NULL) return E_POINTER;
	if (pResult == NULL) return E_POINTER;

	if (m_text.IsEmpty())
		return S_OK;

	CString control_id = GetControlId(pControl);

	int idx = m_ribbon_map.FindKey(control_id);
	if (idx < 0)
		return S_OK;

	CString name = m_ribbon_map.GetValueAt(idx);

	if (StrStrIW(name, m_text))
		*pResult = VARIANT_TRUE;

	return S_OK;
}

#define MAX_HISTORY	10

void CVisioConnect::AddToHistory()
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

STDMETHODIMP CVisioConnect::OnRibbonComboBoxChange(IDispatch * pControl, BSTR *pbstrText)
{
	if (pControl == NULL) return E_POINTER;
	if (pbstrText == NULL) return E_POINTER;

	m_text = *pbstrText;

	OnFind();

	return S_OK;
}

STDMETHODIMP CVisioConnect::GetRibbonComboBoxText(IDispatch*pControl, BSTR *pbstrText)
{
	*pbstrText = m_text.AllocSysString();
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnRibbonLoad(IDispatch* disp)
{
	m_pRibbon = disp;
	return S_OK;
}


STDMETHODIMP CVisioConnect::GetRibbonComboBoxItemCount(IDispatch*pControl, long *count)
{
	*count = m_history.GetSize();
	return S_OK;
}


STDMETHODIMP CVisioConnect::GetRibbonLabel(IDispatch* pControl, BSTR *pbstrLabel)
{
	CString control_id = GetControlId(pControl);

	CString result;

	LanguageLock lock(GetAppLanguage(m_pApplication));

	if (false)
		;
	else if (control_id == L"RibbonTab")
		result.LoadString(IDS_RibbonTab);
	else if (control_id == L"RibbonGroupFind")
		result.LoadString(IDS_RibbonGroupFind);
	else if (control_id == L"RibbonComboBox")
		result.LoadString(IDS_RibbonComboBox);
	else if (control_id == L"RibbonLanguageMenu")
		result.LoadString(IDS_RibbonLanguageMenu);
	else if (control_id == L"RibbonLanguageEN")
		result.LoadString(IDS_RibbonLanguageEN);
	else if (control_id == L"RibbonLanguageDE")
		result.LoadString(IDS_RibbonLanguageDE);
	else if (control_id == L"RibbonLanguageRU")
		result.LoadString(IDS_RibbonLanguageRU);
	else if (control_id == L"RibbonLanguageJP")
		result.LoadString(IDS_RibbonLanguageJP);
	else if (control_id == L"RibbonSearchResults")
		result.LoadString(IDS_RibbonSearchResults);
	
	*pbstrLabel = result.AllocSysString();
	return S_OK;
}

STDMETHODIMP CVisioConnect::GetRibbonComboBoxItemLabel(IDispatch*pControl, long cIndex, BSTR *pbstrLabel)
{
	*pbstrLabel = m_history[m_history.GetSize()-cIndex-1].AllocSysString();
	return S_OK;
}

#include <GdiPlus.h>
#pragma comment(lib,"gdiplus.lib")

HRESULT CustomUiGetPng(LPCWSTR resource_id, IPictureDisp ** result_image, IPictureDisp** result_mask)
{
	HMODULE hModule = AfxGetResourceHandle();
	HRESULT hr = S_OK;

	HRSRC hResource = FindResource (hModule, resource_id, L"PNG");

	if (!hResource)
		return HRESULT_FROM_WIN32(GetLastError());

	DWORD dwImageSize = SizeofResource (hModule, hResource);

	const void* pResourceData = LockResource (LoadResource(hModule, hResource));
	if (!pResourceData)
		return HRESULT_FROM_WIN32(GetLastError());

	using namespace Gdiplus;

	PICTDESC pd = {0};
	pd.cbSizeofstruct = sizeof (PICTDESC);
	pd.picType = PICTYPE_BITMAP;

	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;

	gdiplusStartupInput.DebugEventCallback = NULL;
	gdiplusStartupInput.SuppressBackgroundThread = FALSE;
	gdiplusStartupInput.SuppressExternalCodecs = FALSE;
	gdiplusStartupInput.GdiplusVersion = 1;
	GdiplusStartup (&gdiplusToken, &gdiplusStartupInput, NULL);

	HGLOBAL hBuffer = GlobalAlloc (GMEM_MOVEABLE, dwImageSize);
	if (hBuffer)
	{
		void* pBuffer = GlobalLock (hBuffer);
		if (pBuffer)
		{
			CopyMemory (pBuffer, pResourceData, dwImageSize);

			IStream* pStream = NULL;
			if (::CreateStreamOnHGlobal (hBuffer, FALSE, &pStream) == S_OK)
			{
				Bitmap bitmap(pStream);

				if (result_mask == NULL) // direct support for picture
				{
					bitmap.GetHBITMAP (0, &pd.bmp.hbitmap);
					hr = OleCreatePictureIndirect (&pd, IID_IDispatch, TRUE, (LPVOID*)result_image);
				}
				else	// old version - 2003/2007 - split into picture + mask
				{
					UINT w = bitmap.GetWidth();
					UINT h = bitmap.GetHeight();
					Rect r(0, 0, w, h);

					Bitmap bitmap_rgb(w, h, PixelFormat24bppRGB);
					Bitmap bitmap_msk(w, h, PixelFormat24bppRGB);

					BitmapData argb_bits;
					bitmap.LockBits(&r, ImageLockModeRead, PixelFormat32bppARGB, &argb_bits);

					BitmapData rgb_bits;
					bitmap_rgb.LockBits(&r, ImageLockModeWrite, PixelFormat24bppRGB, &rgb_bits);

					BitmapData msk_bits;
					bitmap_msk.LockBits(&r, ImageLockModeWrite, PixelFormat24bppRGB, &msk_bits);

					for (UINT y = 0; y < h; ++y)
					{
						for (UINT x = 0; x < w; ++x)
						{
							BYTE* idx_argb = 
								static_cast<BYTE*>(argb_bits.Scan0) + (x + y*w) * 4;

							BYTE* idx_rgb = 
								static_cast<BYTE*>(static_cast<BYTE*>(rgb_bits.Scan0) + (x + y*w) * 3);

							BYTE* idx_msk =
								static_cast<BYTE*>(static_cast<BYTE*>(msk_bits.Scan0) + (x + y*w) * 3);


							idx_rgb[0] = idx_argb[0];
							idx_rgb[1] = idx_argb[1];
							idx_rgb[2] = idx_argb[2];

							byte t = (idx_argb[3] < 128) ? 0xFF : 0x00;

							idx_msk[0] = t;
							idx_msk[1] = t;
							idx_msk[2] = t;
						}
					}

					bitmap.UnlockBits(&argb_bits);

					bitmap_rgb.UnlockBits(&rgb_bits);
					bitmap_msk.UnlockBits(&msk_bits);

					bitmap_rgb.GetHBITMAP (0, &pd.bmp.hbitmap);
					hr = OleCreatePictureIndirect (&pd, IID_IDispatch, TRUE, (LPVOID*)result_image);

					bitmap_msk.GetHBITMAP (0, &pd.bmp.hbitmap);
					hr = OleCreatePictureIndirect (&pd, IID_IDispatch, TRUE, (LPVOID*)result_mask);
				}
				pStream->Release();
			}
			GlobalUnlock (pBuffer);
		}
		GlobalFree (hBuffer);
	}

	GdiplusShutdown (gdiplusToken);
	return hr;
}

STDMETHODIMP CVisioConnect::OnRibbonLoadImage(BSTR *bstrID, IPictureDisp ** ppdispImage)
{
	return CustomUiGetPng(*bstrID, ppdispImage, NULL);
}

void CVisioConnect::OnFind()
{
	AddToHistory();
	InvalidateRibbon();
}

void CVisioConnect::InvalidateRibbon()
{
	IRibbonUIPtr ui;
	m_pRibbon->QueryInterface(__uuidof(IRibbonUI), (void**)&ui);
	ui->Invalidate();
}

void CVisioConnect::LoadStringMap()
{
	m_ribbon_map.RemoveAll();
	LanguageLock lock(m_language);

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
}
