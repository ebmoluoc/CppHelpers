#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	bool CopyToClipboard(HWND hWndNewOwner, LPCVOID pData, SIZE_T nBytes, UINT uFormat)
	{
		auto succeeded{ false };

		auto hGlobal{ GlobalAlloc(GMEM_MOVEABLE, nBytes) };
		if (hGlobal != nullptr)
		{
			auto pMem{ GlobalLock(hGlobal) };
			if (pMem != nullptr)
			{
				memcpy(pMem, pData, nBytes);
				GlobalUnlock(hGlobal);

				if (OpenClipboard(hWndNewOwner))
				{
					if (EmptyClipboard())
					{
						succeeded = SetClipboardData(uFormat, hGlobal) != nullptr;
					}

					CloseClipboard();
				}
			}

			if (!succeeded)
			{
				GlobalFree(hGlobal);
			}
		}

		return succeeded;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////////

	bool CopyToClipboard(HWND hWndNewOwner, const std::wstring& str)
	{
		auto nBytes{ (str.length() + 1) * sizeof(WCHAR) };
		return CopyToClipboard(hWndNewOwner, str.c_str(), nBytes, CF_UNICODETEXT);
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
