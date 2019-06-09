#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	bool RegKeyExists(HKEY hKey, LPCWSTR pSubKey)
	{
		HKEY hKeyResult;
		auto exists{ RegOpenKeyExW(hKey, pSubKey, 0, KEY_READ, &hKeyResult) == ERROR_SUCCESS };
		if (exists)
		{
			RegCloseKey(hKeyResult);
		}

		return exists;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	bool RegValueExists(HKEY hKey, LPCWSTR pSubKey, LPCWSTR pValName)
	{
		return RegGetValueW(hKey, pSubKey, pValName, RRF_RT_ANY, nullptr, nullptr, nullptr) == ERROR_SUCCESS;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	bool SetRegValue(HKEY hKey, LPCWSTR pSubKey, LPCWSTR pValName, LPCWSTR pValData, DWORD dwType)
	{
		size_t cbData;
		switch (dwType)
		{
		case REG_SZ:
		case REG_EXPAND_SZ:
			cbData = (wcslen(pValData) + 1) * sizeof(WCHAR);
			break;
		case REG_MULTI_SZ:
			cbData = MultiStringSize(pValData);
			break;
		default:
			return false;
		}

		if (cbData < ULONG_MAX)
		{
			return RegSetKeyValueW(hKey, pSubKey, pValName, dwType, pValData, static_cast<DWORD>(cbData)) == ERROR_SUCCESS;
		}

		return false;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
