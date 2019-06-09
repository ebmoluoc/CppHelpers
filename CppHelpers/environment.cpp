#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

static const WCHAR BACKSLASH{ '\\' };
static const WCHAR QUOTE{ '\"' };

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	// https://docs.microsoft.com/en-us/cpp/cpp/parsing-cpp-command-line-arguments
	std::wstring EscapeArgument(std::wstring argument)
	{
		if (!argument.empty())
		{
			if (argument.find_first_of(L" \"\t") == argument.npos)
			{
				return argument;
			}

			argument.reserve(argument.length() * 2);
			auto index{ argument.length() };
			auto quote{ true };

			while (--index != argument.npos)
			{
				if (quote && argument[index] == BACKSLASH)
				{
					argument.insert(index, 1, BACKSLASH);
				}
				else if (argument[index] == QUOTE)
				{
					argument.insert(index, 1, BACKSLASH);
					quote = true;
				}
				else if (quote == true)
				{
					quote = false;
				}
			}
		}

		return std::wstring{ QUOTE + argument + QUOTE };
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
