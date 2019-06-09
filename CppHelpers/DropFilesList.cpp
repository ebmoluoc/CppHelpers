#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

#pragma warning(suppress: 26495) // stgm_ doesn't need to be initialized.

	DropFilesList::DropFilesList() : pList_{ nullptr }
	{
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	DropFilesList::~DropFilesList()
	{
		Unload();
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	// https://docs.microsoft.com/en-us/windows/desktop/shell/clipboard#cf_hdrop
	bool DropFilesList::Load(LPDATAOBJECT pDataObject)
	{
		Unload();

		FORMATETC fetc{ CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		auto succeeded{ SUCCEEDED(pDataObject->GetData(&fetc, &stgm_)) };
		if (succeeded)
		{
			auto pDropFiles{ static_cast<DROPFILES*>(stgm_.hGlobal) };
			pList_ = reinterpret_cast<LPCWSTR>((BYTE*)pDropFiles + pDropFiles->pFiles);
		}

		return succeeded;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	void DropFilesList::Unload()
	{
		if (pList_ != nullptr)
		{
			pList_ = nullptr;
			ReleaseStgMedium(&stgm_);
		}
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	bool DropFilesList::IsMultiItems()
	{
		if (pList_ != nullptr)
		{
			auto p{ pList_ };
			p += wcslen(p) + 1;

			return *p != 0;
		}

		return false;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::vector<std::wstring> DropFilesList::GetItems()
	{
		std::vector<std::wstring> items{};

		if (pList_ != nullptr)
		{
			auto p{ pList_ };
			while (*p)
			{
				items.push_back(std::wstring{ p });
				p += wcslen(p) + 1;
			}
		}

		return items;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
