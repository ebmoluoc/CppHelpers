#pragma once

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// CppHelpers 1.4
// Windows 7 and above
//
///////////////////////////////////////////////////////////////////////////////////////////////////

#include <string>
#include <vector>
#include <chrono>
#include <shlobj.h>

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	/////////////////////////////////////// clipboard /////////////////////////////////////////////

	// hWndNewOwner : Window associated with the open clipboard or nullptr to be associated with the current process.
	// pData : Data to be copied to the clipboard.
	// nBytes : Size in bytes of the data to be copied to the clipboard.
	// uFormat : Format of clipboard data.
	// Returns : True if successful.
	bool CopyToClipboard(HWND hWndNewOwner, LPCVOID pData, SIZE_T nBytes, UINT uFormat);

	// hWndNewOwner : Window associated with the open clipboard or nullptr to be associated with the current process.
	// str : String to be copied to the clipboard (CF_UNICODETEXT).
	// Returns : True if successful.
	bool CopyToClipboard(HWND hWndNewOwner, const std::wstring& str);

	/////////////////////////////////////// CodeTiming ////////////////////////////////////////////

	class CodeTiming
	{
	public:
		void Start() { start_ = std::chrono::steady_clock::now(); }
		void Stop() { stop_ = std::chrono::steady_clock::now(); }
		template <typename T>
		std::wstring Result() { return std::to_wstring(std::chrono::duration_cast<T>(stop_ - start_).count()); }
	private:
		std::chrono::time_point<std::chrono::steady_clock> start_;
		std::chrono::time_point<std::chrono::steady_clock> stop_;
	};

	/////////////////////////////////////// com ///////////////////////////////////////////////////

	// guid : String representation of the GUID "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}".
	// Returns : GUID created from the string if successful or GUID_NULL otherwise.
	GUID CreateGUID(LPCWSTR guid);

	/////////////////////////////////////// DropFilesList /////////////////////////////////////////

	class DropFilesList
	{
	public:
		DropFilesList();
		~DropFilesList();

		// pDataObject : Pointer to a data object interface (IDataObject).
		// Returns true if the drop files list from the data object has been loaded successfully.
		bool Load(LPDATAOBJECT pDataObject);

		// Free the resources.
		void Unload();

		// Returns true if the list has more than 1 item or false otherise (including no data object loaded).
		bool IsMultiItems();

		// Returns all the items in the list or an empty container if there is no data object loaded.
		std::vector<std::wstring> GetItems();

	protected:
		// Pointer to a sequence of null-terminated strings, terminated by a null character "c:\\temp1.txt\0c:\\temp2.txt\0\0".
		LPCWSTR pList_;

	private:
		STGMEDIUM stgm_;
	};

	/////////////////////////////////////// environment ///////////////////////////////////////////

	// argument : Argument to be escaped.
	// Returns : String containing the escaped argument.
	std::wstring EscapeArgument(std::wstring argument);

	/////////////////////////////////////// image /////////////////////////////////////////////////

	// hIcon : Handle of the icon.
	// Returns : Handle of the newly created bitmap if successful or nullptr otherwise. Call DeleteObject on this handle.
	HBITMAP BitmapFromIcon(HICON hIcon);

	// hModule : Handle of the module that contains the icon resource.
	// idIcon : Resource identifier of the icon to be loaded.
	// width : Width, in pixels, of the icon. Can be zero to use the actual icon width.
	// height : Height, in pixels, of the icon. Can be zero to use the actual icon height.
	// Returns : Handle of the newly created bitmap if successful or nullptr otherwise. Call DeleteObject on this handle.
	HBITMAP BitmapFromIconResource(HMODULE hModule, WORD idIcon, int width, int height);

	/////////////////////////////////////// menu //////////////////////////////////////////////////

	// hMenu : Handle of the menu in which the new item is added.
	// nItemPos : Position index of the menu item before which to add the new item.
	// pMenuText : Menu item text or nullptr to add a separator (subsequent parameters are useless).
	// idCmd : Application-defined value that identifies the menu item.
	// hSubMenu : Handle of the submenu associated with the menu item.
	// hBmpItem : Handle of the bitmap to be displayed.
	// Returns : True if successful.
	bool AddMenuItem(HMENU hMenu, UINT nItemPos, LPCWSTR pMenuText, UINT idCmd, HMENU hSubMenu, HBITMAP hBmpItem);

	// hMenu : Handle of the menu in which to search for the command id.
	// commandId : Command id to search for.
	// searchSubmenus : If the item is not found on the specified menu, search submenus if there is any.
	// hMenuFound : Handle of the menu or submenu where the command id was found.
	// Returns : Position index if the command id was found or -1 otherwise.
	int GetMenuItemPosition(HMENU hMenu, UINT commandId, bool searchSubmenus, HMENU& hMenuFound);

	/////////////////////////////////////// path //////////////////////////////////////////////////

	// knownFolder : Directory path given by a KNOWNFOLDERID.
	// subdirName : Subdirectory of knownFolder (no leading or trailing backslash - can be empty).
	// fileName : File name to get a path in knownFolder\subdirName.
	// Returns : File path if successful or an empty string otherwise.
	std::wstring GetFilePath(const KNOWNFOLDERID& knownFolder, const std::wstring& subdirName, const std::wstring& fileName);

	// fullPath : Full path of a file or directory.
	// newName : New name of the last path item (after the last backslash).
	// Returns : String containing the renamed path (doesn't rename the actual item on disk if it exists).
	std::wstring RenamePath(const std::wstring& fullPath, const std::wstring& newName);

	/////////////////////////////////////// registry //////////////////////////////////////////////

	// hKey : Registry key handle or a predefined key like HKEY_CLASSES_ROOT or HKEY_LOCAL_MACHINE.
	// pSubKey : Registry subkey relative to hKey.
	// Returns : True if the key exists.
	bool RegKeyExists(HKEY hKey, LPCWSTR pSubKey);

	// hKey : Registry key handle or a predefined key like HKEY_CLASSES_ROOT or HKEY_LOCAL_MACHINE.
	// pSubKey : Registry subkey relative to hKey.
	// pValName : Name of the value.
	// Returns : True if the value exits.
	bool RegValueExists(HKEY hKey, LPCWSTR pSubKey, LPCWSTR pValName);

	// hKey : Registry key handle or a predefined key like HKEY_CLASSES_ROOT or HKEY_LOCAL_MACHINE.
	// pSubKey : Registry subkey relative to hKey.
	// pValName : Name of the value to be set.
	// pValData : String value to be set.
	// dwType : Type of registry data (REG_SZ or REG_EXPAND_SZ or REG_MULTI_SZ).
	// Returns : True if successful.
	bool SetRegValue(HKEY hKey, LPCWSTR pSubKey, LPCWSTR pValName, LPCWSTR pValData, DWORD dwType);

	/////////////////////////////////////// resource //////////////////////////////////////////////

	// hModule : Handle of the module that contains the icon resource.
	// idIcon : Resource identifier of the icon to be loaded.
	// width : Width, in pixels, of the icon. Can be zero to use the actual icon width.
	// height : Height, in pixels, of the icon. Can be zero to use the actual icon height.
	// Returns : Handle of the icon if successful or nullptr otherwise. Call DestroyIcon on this handle.
	HICON LoadIconResource(HMODULE hModule, WORD idIcon, int width, int height);

	// hModule : Handle of the module that contains the string resource.
	// idString : Resource identifier of the string to be loaded.
	// Returns : String if successful or an empty string otherwise.
	std::wstring LoadStringResource(HMODULE hModule, WORD idString);

	/////////////////////////////////////// string ////////////////////////////////////////////////

	// str : String to be processed.
	// Returns : String where all the backslashes are escaped (doubled).
	std::wstring EscapeBackslash(std::wstring str);

	// strings : Strings to be joined.
	// separator : String used as separator.
	// avgLength : Approximation of the average length of a string in the container (optimization purpose).
	// Returns : String made of all the strings joined together and separated by the specified separator.
	std::wstring JoinStrings(const std::vector<std::wstring>& strings, const std::wstring& separator, size_t avgLength);

	// str : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : Count of null-terminated strings in the sequence.
	size_t MultiStringCount(LPCWSTR str);

	// str : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : Size in bytes of the sequence including the null terminator.
	size_t MultiStringSize(LPCWSTR str);

	// str : String to be trimmed.
	// chr : Leading and trailing characters to be removed.
	// Returns : Trimmed string.
	std::wstring TrimString(std::wstring str, WCHAR chr);

	// str : String to be trimmed.
	// chr : Trailing characters to be removed.
	// Returns : Trimmed string.
	std::wstring TrimStringBack(std::wstring str, WCHAR chr);

	// str : String to be trimmed.
	// chr : Leading characters to be removed.
	// Returns : Trimmed string.
	std::wstring TrimStringFront(std::wstring str, WCHAR chr);

	/////////////////////////////////////// synchronization ///////////////////////////////////////

	// name : Name of the mutex.
	// Returns : True if the mutex exists.
	bool MutexExists(LPCWSTR name);

	/////////////////////////////////////// ? /////////////////////////////////////////////////////

}
