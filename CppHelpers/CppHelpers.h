///////////////////////////////////////////////////////////////////////////////////////////////////
// 
// CppHelpers 1.5
// Windows 7 and above
//
// MIT License
//
// Copyright(c) 2019 Philippe Coulombe
// https://github.com/ebmoluoc/CppHelpers
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this softwareand associated documentation files(the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and /or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions :
//
// The above copyright noticeand this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//
///////////////////////////////////////////////////////////////////////////////////////////////////

#pragma once

#include <string>
#include <vector>
#include <chrono>
#include <memory>
#include <shlobj.h>
#include <setupapi.h>

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      clipboard
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// hWndNewOwner : Window associated with the open clipboard or nullptr to be associated with the current process.
	// pData : Data to be copied to the clipboard.
	// nBytes : Size in bytes of the data to be copied to the clipboard.
	// uFormat : Format of clipboard data.
	// Returns : True if successful.
	bool CopyToClipboard(HWND hWndNewOwner, LPCVOID pData, SIZE_T nBytes, UINT uFormat);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// hWndNewOwner : Window associated with the open clipboard or nullptr to be associated with the current process.
	// str : String to be copied to the clipboard (CF_UNICODETEXT).
	// Returns : True if successful.
	bool CopyToClipboard(HWND hWndNewOwner, const std::wstring& str);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      CodeTiming
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// Simple class for code timing.
	class CodeTiming
	{
	public:
		void Start() { start_ = std::chrono::steady_clock::now(); }
		void Stop() { stop_ = std::chrono::steady_clock::now(); }
		template <typename T>
		std::wstring Result() const { return std::to_wstring(std::chrono::duration_cast<T>(stop_ - start_).count()); }
	private:
		std::chrono::time_point<std::chrono::steady_clock> start_;
		std::chrono::time_point<std::chrono::steady_clock> stop_;
	};

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      com
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// pGuidString : String representation of the GUID "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}".
	// Returns : GUID created from the string if successful or GUID_NULL otherwise.
	GUID CreateGUID(LPCWSTR pGuidString);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// Class for reading a DROPFILES structure.
	class DropFilesList
	{
	public:
		DropFilesList();
		~DropFilesList();
		// pDataObject : Pointer to a data object interface (IDataObject).
		// Returns : True if the DROPFILES structure has been loaded successfully.
		bool Load(LPDATAOBJECT pDataObject);
		// Free the resources (doesn't need to be called before Load).
		void Unload();
		// Returns : True if the list has more than 1 item or false otherise (including no data loaded).
		bool IsMultiItems() const;
		// Returns : First item in the list or an empty string if there is no data loaded.
		std::wstring GetFirstItem() const;
		// Returns : All the items in the list or an empty container if there is no data loaded.
		std::vector<std::wstring> GetItems() const;
	protected:
		// Pointer to a sequence of null-terminated strings, terminated by a null character "c:\\temp1.txt\0c:\\temp2.txt\0\0".
		LPCWSTR pList_;
		// If fWide_ is not true, pList_ must be cast as LPCSTR to be usable.
		bool fWide_;
	private:
		STGMEDIUM stgm_;
	};

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      device
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// Class to work with device information set.
	class DeviceInformationSet
	{
	public:
		DeviceInformationSet();
		~DeviceInformationSet();
		// classGuid : GUID of a device setup class or a device interface class (can be null).
		// enumerator : GUID or symbolic name to filter devices returned (can be null).
		// hwndParent : Handle to a top-level window to be used for a user interface (can be null).
		// flags : DIGCF_* value to filter devices returned (can be 0).
		// Returns : True if the device information set has been loaded successfully.
		bool Load(LPCGUID classGuid, LPCWSTR enumerator, HWND hwndParent, DWORD flags);
		// Free the resources (doesn't need to be called before Load).
		void Unload();
		// pDeviceName : Device name.
		// Returns : Device path if successful or an empty string otherwise.
		std::wstring GetDevicePath(LPCWSTR pDeviceName) const;
		// Returns : All the device paths from the device information set or an empty container otherwise.
		std::vector<std::wstring> GetDevicePaths() const;
	private:
		HDEVINFO hDevInfo_;
		std::unique_ptr<GUID> classGuid_;
		std::unique_ptr<WCHAR[]> enumerator_;
		HWND hwndParent_;
		DWORD flags_;
	};

	///////////////////////////////////////////////////////////////////////////////////////////////

	// path : Path to be processed.
	// Returns : 0 if 8dot3name is enabled, 1 if 8dot3name is disabled or -1 in case of errors. Administrative privileges required.
	int GetShortPathCreationValue(LPCWSTR path);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// path : Path to be processed.
	// trailingBackslash : True to keep the trailing backslash or false otherwise.
	// Returns: Volume GUID path if successful or an empty string otherwise.
	std::wstring GetVolumeGuidPath(LPCWSTR path, bool trailingBackslash);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// Class for reading a STORAGE_DEVICE_DESCRIPTOR structure.
	class StorageDeviceDescriptor
	{
	public:
		StorageDeviceDescriptor();
		// pDeviceName : Device name.
		// Returns : True if the STORAGE_DEVICE_DESCRIPTOR structure has been loaded successfully.
		bool Load(LPCWSTR pDeviceName);
		// Free the resources (doesn't need to be called before Load).
		void Unload();
		// The SCSI-2 device type.
		BYTE DeviceType() const;
		// The SCSI-2 device type modifier (if any) - this may be zero.
		BYTE DeviceTypeModifier() const;
		// Flag indicating whether the device's media (if any) is removable.
		BOOLEAN RemovableMedia() const;
		// Flag indicating whether the device can support mulitple outstanding commands.
		// The actual synchronization in this case is the responsibility of the port driver.
		BOOLEAN CommandQueueing() const;
		// Contains the bus type of the device.
		STORAGE_BUS_TYPE BusType() const;
		// Device's vendor id.
		std::wstring VendorId() const;
		// Device's product id.
		std::wstring ProductId() const;
		// Device's product revision.
		std::wstring ProductRevision() const;
		// Device's serial number.
		std::wstring SerialNumber() const;
		// Bus specific property data.
		std::vector<BYTE> RawDeviceProperties() const;
	private:
		PSTORAGE_DEVICE_DESCRIPTOR pDescriptor_;
		std::unique_ptr<BYTE[]> descriptorBuffer_;
		std::wstring GetStringData(DWORD offset) const;
	};

	///////////////////////////////////////////////////////////////////////////////////////////////

	// Class for reading a STORAGE_DEVICE_NUMBER structure.
	class StorageDeviceNumber
	{
	public:
		StorageDeviceNumber();
		// pDeviceName : Device name.
		// Returns : True if the STORAGE_DEVICE_NUMBER structure has been loaded successfully.
		bool Load(LPCWSTR pDeviceName);
		// Free the resources (doesn't need to be called before Load).
		void Unload();
		// FILE_DEVICE_XXX type for the device.
		DEVICE_TYPE DeviceType() const;
		// Number of the device.
		DWORD DeviceNumber() const;
		// Partition number of the device if it is partitionable or -1 otherwise.
		DWORD PartitionNumber() const;
	private:
		STORAGE_DEVICE_NUMBER sdn_;
	};

	///////////////////////////////////////////////////////////////////////////////////////////////

	// Class for reading a VOLUME_DISK_EXTENTS / DISK_EXTENT structure.
	class VolumeDiskExtents
	{
	public:
		VolumeDiskExtents();
		// pDeviceName : Device name.
		// Returns : True if the VOLUME_DISK_EXTENTS structure has been loaded successfully.
		bool Load(LPCWSTR pDeviceName);
		// Free the resources (doesn't need to be called before Load).
		void Unload();
		// Returns : All the DISK_EXTENT structures from VOLUME_DISK_EXTENTS or an empty container otherwise.
		std::vector<DISK_EXTENT> GetDiskExtents() const;
	private:
		PVOLUME_DISK_EXTENTS pVolDiskExt_;
		std::unique_ptr<BYTE[]> volDiskExtBuffer_;
		void CreateVolDiskExtBuffer(DWORD size);
	};

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      environment
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// argument : Argument to be escaped.
	// Returns : String containing the escaped argument.
	std::wstring EscapeArgument(std::wstring argument);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      image
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// hIcon : Handle of the icon.
	// Returns : Handle of the newly created bitmap if successful or nullptr otherwise. Call DeleteObject on this handle.
	HBITMAP BitmapFromIcon(HICON hIcon);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// hModule : Handle of the module that contains the icon resource.
	// idIcon : Resource identifier of the icon to be loaded.
	// width : Width, in pixels, of the icon. Can be zero to use the actual icon width.
	// height : Height, in pixels, of the icon. Can be zero to use the actual icon height.
	// Returns : Handle of the newly created bitmap if successful or nullptr otherwise. Call DeleteObject on this handle.
	HBITMAP BitmapFromIconResource(HMODULE hModule, WORD idIcon, int width, int height);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      menu
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// hMenu : Handle of the menu in which the new item is added.
	// nItemPos : Position index of the menu item before which to add the new item.
	// pMenuText : Menu item text or nullptr to add a separator (subsequent parameters are useless).
	// idCmd : Application-defined value that identifies the menu item.
	// hSubMenu : Handle of the submenu associated with the menu item.
	// hBmpItem : Handle of the bitmap to be displayed.
	// Returns : True if successful.
	bool AddMenuItem(HMENU hMenu, UINT nItemPos, LPCWSTR pMenuText, UINT idCmd, HMENU hSubMenu, HBITMAP hBmpItem);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// hMenu : Handle of the menu in which to search for the command id.
	// commandId : Command id to search for.
	// searchSubmenus : If the item is not found on the specified menu, search submenus if there is any.
	// hMenuFound : Handle of the menu or submenu where the command id was found.
	// Returns : Position index if the command id was found or -1 otherwise.
	int GetMenuItemPosition(HMENU hMenu, UINT commandId, bool searchSubmenus, HMENU& hMenuFound);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      path
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// knownFolder : Directory path given by a KNOWNFOLDERID.
	// subdirName : Subdirectory of knownFolder (no leading or trailing backslash - can be empty).
	// fileName : File name to get a path in knownFolder\subdirName.
	// Returns : File path if successful or an empty string otherwise.
	std::wstring GetFilePath(const KNOWNFOLDERID& knownFolder, const std::wstring& subdirName, const std::wstring& fileName);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// fullPath : Full path of a file or directory.
	// newName : New name of the last path item (after the last backslash).
	// Returns : String containing the renamed path (doesn't rename the actual item on disk if it exists).
	std::wstring RenamePath(const std::wstring& fullPath, const std::wstring& newName);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      registry
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// hKey : Registry key handle or a predefined key like HKEY_CLASSES_ROOT or HKEY_LOCAL_MACHINE.
	// pSubKey : Registry subkey relative to hKey.
	// Returns : True if the key exists.
	bool RegKeyExists(HKEY hKey, LPCWSTR pSubKey);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// hKey : Registry key handle or a predefined key like HKEY_CLASSES_ROOT or HKEY_LOCAL_MACHINE.
	// pSubKey : Registry subkey relative to hKey.
	// pValName : Name of the value.
	// Returns : True if the value exits.
	bool RegValueExists(HKEY hKey, LPCWSTR pSubKey, LPCWSTR pValName);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// hKey : Registry key handle or a predefined key like HKEY_CLASSES_ROOT or HKEY_LOCAL_MACHINE.
	// pSubKey : Registry subkey relative to hKey.
	// pValName : Name of the value to be set.
	// pValData : String value to be set.
	// dwType : Type of registry data (REG_SZ or REG_EXPAND_SZ or REG_MULTI_SZ).
	// Returns : True if successful.
	bool SetRegValue(HKEY hKey, LPCWSTR pSubKey, LPCWSTR pValName, LPCWSTR pValData, DWORD dwType);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      resource
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// hModule : Handle of the module that contains the icon resource.
	// idIcon : Resource identifier of the icon to be loaded.
	// width : Width, in pixels, of the icon. Can be zero to use the actual icon width.
	// height : Height, in pixels, of the icon. Can be zero to use the actual icon height.
	// Returns : Handle of the icon if successful or nullptr otherwise. Call DestroyIcon on this handle.
	HICON LoadIconResource(HMODULE hModule, WORD idIcon, int width, int height);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// hModule : Handle of the module that contains the string resource.
	// idString : Resource identifier of the string to be loaded.
	// Returns : String if successful or an empty string otherwise.
	std::wstring LoadStringResource(HMODULE hModule, WORD idString);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      string
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// str : String to be processed.
	// Returns : String where all the backslashes are escaped (doubled).
	std::wstring EscapeBackslash(std::wstring str);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// strings : Strings to be joined.
	// separator : String used as separator.
	// avgLength : Approximation of the average length of a string in the container (optimization purpose).
	// Returns : String made of all the strings joined together and separated by the specified separator.
	std::wstring JoinStrings(const std::vector<std::wstring>& strings, const std::wstring& separator, size_t avgLength);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : True if the sequence has more than one item.
	bool IsMultiSzItems(LPCSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : True if the sequence has more than one item.
	bool IsMultiSzItems(LPCWSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : Count of null-terminated strings in the sequence.
	size_t GetMultiSzCount(LPCSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : Count of null-terminated strings in the sequence.
	size_t GetMultiSzCount(LPCWSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : Size in bytes of the sequence including the null terminator or 0 if str is null.
	size_t GetMultiSzSize(LPCSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : Size in bytes of the sequence including the null terminator or 0 if str is null.
	size_t GetMultiSzSize(LPCWSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : All the items in the sequence or an empty container if str is null or points to a null terminator.
	std::vector<std::string> GetMultiSzItems(LPCSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : All the items in the sequence (converted to wide character string) or an empty container if str is null or points to a null terminator.
	std::vector<std::wstring> GetMultiSzItemsWide(LPCSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pMultiSz : Pointer to a null-terminated sequence of null-terminated strings "c:\\temp1.txt\0c:\\temp2.txt\0\0".
	// Returns : All the items in the sequence or an empty container if str is null or points to a null terminator.
	std::vector<std::wstring> GetMultiSzItems(LPCWSTR pMultiSz);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// str : String to be trimmed.
	// chr : Leading and trailing characters to be removed.
	// Returns : Trimmed string.
	std::wstring TrimString(std::wstring str, WCHAR chr);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// str : String to be trimmed.
	// chr : Trailing characters to be removed.
	// Returns : Trimmed string.
	std::wstring TrimStringBack(std::wstring str, WCHAR chr);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// str : String to be trimmed.
	// chr : Leading characters to be removed.
	// Returns : Trimmed string.
	std::wstring TrimStringFront(std::wstring str, WCHAR chr);

	///////////////////////////////////////////////////////////////////////////////////////////////

	// pStr : String to be processed.
	// Returns : Wide character string.
	std::wstring WStrFromStr(LPCSTR pStr);

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      synchronization
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	// pName : Name of the mutex.
	// Returns : True if the mutex exists.
	bool MutexExists(LPCWSTR pName);

	///////////////////////////////////////////////////////////////////////////////////////////////

}
