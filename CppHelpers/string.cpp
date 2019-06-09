#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

static const WCHAR BACKSLASH{ '\\' };

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	std::wstring EscapeBackslash(std::wstring str)
	{
		auto index{ str.length() };

		while (--index != str.npos)
		{
			if (str[index] == BACKSLASH)
			{
				str.insert(index, 1, BACKSLASH);
			}
		}

		return str;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::wstring JoinStrings(const std::vector<std::wstring>& strings, const std::wstring& separator, size_t avgLength)
	{
		if (!strings.empty())
		{
			auto it{ strings.begin() };
			auto end{ strings.end() };
			auto buffer{ *it };

			buffer.reserve(strings.size() * avgLength);

			while (++it != end)
			{
				buffer.append(separator).append(*it);
			}

			return buffer;
		}

		return std::wstring{};
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	size_t MultiStringCount(LPCWSTR str)
	{
		size_t count{ 0 };

		while (*str)
		{
			++count;
			str += wcslen(str) + 1;
		}

		return count;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	size_t MultiStringSize(LPCWSTR str)
	{
		auto p{ str };
		while (*p)
		{
			p += wcslen(p) + 1;
		}

		return ((p - str) + 1) * sizeof(WCHAR);
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::wstring TrimString(std::wstring str, WCHAR chr)
	{
		return str.erase(str.find_last_not_of(chr) + 1).erase(0, str.find_first_not_of(chr));
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::wstring TrimStringBack(std::wstring str, WCHAR chr)
	{
		return str.erase(str.find_last_not_of(chr) + 1);
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::wstring TrimStringFront(std::wstring str, WCHAR chr)
	{
		return str.erase(0, str.find_first_not_of(chr));
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
