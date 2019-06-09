#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	HICON LoadIconResource(HMODULE hModule, WORD idIcon, int width, int height)
	{
		return static_cast<HICON>(LoadImageW(hModule, reinterpret_cast<LPCWSTR>(idIcon), IMAGE_ICON, width, height, LR_DEFAULTCOLOR));
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::wstring LoadStringResource(HMODULE hModule, WORD idString)
	{
		WCHAR buffer[4096];
		if (LoadStringW(hModule, idString, buffer, _countof(buffer)) > 0)
		{
			return std::wstring{ buffer };
		}

		return std::wstring{};
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
