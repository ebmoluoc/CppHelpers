#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

static const WCHAR BACKSLASH{ '\\' };

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	std::wstring GetFilePath(const KNOWNFOLDERID& knownFolder, const std::wstring& subdirName, const std::wstring& fileName)
	{
		std::wstring filePath{};

		LPWSTR path;
		if (SUCCEEDED(SHGetKnownFolderPath(knownFolder, KF_FLAG_DEFAULT, nullptr, &path)))
		{
			filePath = path;
			CoTaskMemFree(path);

			if (!subdirName.empty())
			{
				filePath.insert(filePath.length(), 1, BACKSLASH).append(subdirName);
			}

			filePath.insert(filePath.length(), 1, BACKSLASH).append(fileName);
		}

		return filePath;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::wstring RenamePath(const std::wstring& fullPath, const std::wstring& newName)
	{
		auto pos{ fullPath.find_last_of(BACKSLASH) + 1 };
		return fullPath.substr(0, pos).append(newName);
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
