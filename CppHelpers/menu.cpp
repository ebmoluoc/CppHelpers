#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	bool AddMenuItem(HMENU hMenu, UINT nItemPos, LPCWSTR pMenuText, UINT idCmd, HMENU hSubMenu, HBITMAP hBmpItem)
	{
		MENUITEMINFOW mii{ sizeof(MENUITEMINFOW) };

		if (pMenuText != nullptr)
		{
			mii.fMask = MIIM_STRING | MIIM_ID;
			mii.dwTypeData = const_cast<LPWSTR>(pMenuText);
			mii.wID = idCmd;

			if (hSubMenu != nullptr)
			{
				mii.fMask |= MIIM_SUBMENU;
				mii.hSubMenu = hSubMenu;
			}

			if (hBmpItem != nullptr)
			{
				mii.fMask |= MIIM_BITMAP;
				mii.hbmpItem = hBmpItem;
			}
		}
		else
		{
			mii.fMask = MIIM_FTYPE;
			mii.fType = MFT_SEPARATOR;
		}

		return InsertMenuItemW(hMenu, nItemPos, TRUE, &mii) != FALSE;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	int GetMenuItemPosition(HMENU hMenu, UINT commandId, bool searchSubmenus, HMENU& hMenuFound)
	{
		auto handles{ std::queue<HMENU>({ hMenu }) };

		do
		{
			hMenuFound = handles.front();
			handles.pop();
			auto count{ GetMenuItemCount(hMenuFound) };

			for (int index = 0; index < count; ++index)
			{
				if (GetMenuItemID(hMenuFound, index) == commandId)
				{
					return index;
				}

				if (searchSubmenus)
				{
					auto submenu{ GetSubMenu(hMenuFound, index) };
					if (submenu != nullptr)
					{
						handles.push(submenu);
					}
				}
			}

		} while (!handles.empty());

		hMenuFound = nullptr;
		return -1;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
