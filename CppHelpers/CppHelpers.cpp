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

#define NTDDI_VERSION	0x06010000
#define WINVER			0x0601
#define _WIN32_WINDOWS	0x0601
#define _WIN32_WINNT	0x0601
#define _WIN32_IE		0x0A00

///////////////////////////////////////////////////////////////////////////////////////////////////

#include <queue>
#include <wincodec.h>
#include "CppHelpers.h"

#pragma comment (lib, "setupapi.lib")

///////////////////////////////////////////////////////////////////////////////////////////////////

typedef DWORD ARGB;

static const WCHAR BACKSLASH{ '\\' };
static const WCHAR QUOTE{ '\"' };

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      clipboard
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

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
	//
	//                                      CodeTiming
	//
	///////////////////////////////////////////////////////////////////////////////////////////////



	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      com
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	GUID CreateGUID(LPCWSTR pGuidString)
	{
		IID iid;
		if (SUCCEEDED(IIDFromString(pGuidString, &iid)))
		{
			return iid;
		}

		return GUID_NULL;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

#pragma warning(suppress: 26495) // stgm_ doesn't need to be initialized.

	DropFilesList::DropFilesList() : pList_{ nullptr }
	{
	}

	DropFilesList::~DropFilesList()
	{
		Unload();
	}

	// https://docs.microsoft.com/en-us/windows/desktop/shell/clipboard#cf_hdrop
	bool DropFilesList::Load(LPDATAOBJECT pDataObject)
	{
		Unload();

		FORMATETC fetc{ CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		if (SUCCEEDED(pDataObject->GetData(&fetc, &stgm_)))
		{
			auto pDropFiles{ static_cast<DROPFILES*>(stgm_.hGlobal) };
			pList_ = reinterpret_cast<LPCWSTR>((BYTE*)pDropFiles + pDropFiles->pFiles);
			fWide_ = pDropFiles->fWide;
		}

		return pList_ != nullptr;
	}

	void DropFilesList::Unload()
	{
		if (pList_ != nullptr)
		{
			pList_ = nullptr;
			ReleaseStgMedium(&stgm_);
		}
	}

	bool DropFilesList::IsMultiItems() const
	{
		if (fWide_)
		{
			return IsMultiSzItems(pList_);
		}

		return IsMultiSzItems(reinterpret_cast<LPCSTR>(pList_));
	}

	std::wstring DropFilesList::GetFirstItem() const
	{
		if (fWide_)
		{
			return std::wstring{ pList_ };
		}

		return WStrFromStr(reinterpret_cast<LPCSTR>(pList_));
	}

	std::vector<std::wstring> DropFilesList::GetItems() const
	{
		if (fWide_)
		{
			return GetMultiSzItems(pList_);
		}

		return GetMultiSzItemsWide(reinterpret_cast<LPCSTR>(pList_));
	}

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      device
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

#pragma warning(suppress: 26495) // hwndParent_ and flags_ don't need to be initialized.

	DeviceInformationSet::DeviceInformationSet() : hDevInfo_{ INVALID_HANDLE_VALUE }
	{
	}

	DeviceInformationSet::~DeviceInformationSet()
	{
		Unload();
	}

	bool DeviceInformationSet::Load(LPCGUID classGuid, LPCWSTR enumerator, HWND hwndParent, DWORD flags)
	{
		Unload();

		hDevInfo_ = SetupDiGetClassDevsW(classGuid, enumerator, hwndParent, flags);
		auto succeeded{ hDevInfo_ != INVALID_HANDLE_VALUE };

		if (succeeded)
		{
			if (classGuid != nullptr)
			{
				classGuid_.reset(new GUID());
				memcpy(classGuid_.get(), classGuid, sizeof(GUID));
			}
			else
			{
				classGuid_.reset();
			}

			if (enumerator != nullptr)
			{
				auto length{ wcslen(enumerator) + 1 };
				enumerator_.reset(new WCHAR[length]);
				memcpy(enumerator_.get(), enumerator, length * sizeof(WCHAR));
			}
			else
			{
				enumerator_.reset();
			}

			hwndParent_ = hwndParent;
			flags_ = flags;
		}

		return succeeded;
	}

	void DeviceInformationSet::Unload()
	{
		if (hDevInfo_ != INVALID_HANDLE_VALUE)
		{
			SetupDiDestroyDeviceInfoList(hDevInfo_);
			hDevInfo_ = INVALID_HANDLE_VALUE;
		}
	}

	std::wstring DeviceInformationSet::GetDevicePath(LPCWSTR pDeviceName) const
	{
		if (hDevInfo_ != INVALID_HANDLE_VALUE)
		{
			StorageDeviceNumber deviceSdn;
			if (deviceSdn.Load(pDeviceName))
			{
				StorageDeviceNumber disSdn;
				for (const auto& devicePath : GetDevicePaths())
				{
					if (disSdn.Load(devicePath.c_str()))
					{
						if (disSdn.DeviceNumber() == deviceSdn.DeviceNumber())
						{
							return devicePath;
						}
					}
				}
			}
		}

		return std::wstring{};
	}

	std::vector<std::wstring> DeviceInformationSet::GetDevicePaths() const
	{
		std::vector<std::wstring> devicePaths{};

		if (hDevInfo_ != INVALID_HANDLE_VALUE)
		{
			SP_DEVICE_INTERFACE_DATA diData{ sizeof(SP_DEVICE_INTERFACE_DATA) };
			DWORD memberIndex{ 0 };

			while (SetupDiEnumDeviceInterfaces(hDevInfo_, nullptr, classGuid_.get(), memberIndex++, &diData))
			{
				DWORD requiredSize;

				SetupDiGetDeviceInterfaceDetailW(hDevInfo_, &diData, nullptr, 0, &requiredSize, nullptr);
				if (GetLastError() == ERROR_INSUFFICIENT_BUFFER)
				{
					auto buffer{ std::make_unique<BYTE[]>(requiredSize) };
					auto pDiDetailData{ reinterpret_cast<PSP_DEVICE_INTERFACE_DETAIL_DATA_W>(buffer.get()) };
					pDiDetailData->cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA_W);

					if (!SetupDiGetDeviceInterfaceDetailW(hDevInfo_, &diData, pDiDetailData, requiredSize, nullptr, nullptr))
					{
						devicePaths.clear();
						break;
					}

					devicePaths.push_back(std::wstring{ pDiDetailData->DevicePath });
				}
			}
		}

		return devicePaths;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	int GetShortPathCreationValue(LPCWSTR path)
	{
		auto result{ -1 };
		auto volumeGuidPath{ GetVolumeGuidPath(path, false) };

		auto hDevice{ CreateFileW(volumeGuidPath.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr) };
		if (hDevice != INVALID_HANDLE_VALUE)
		{
			DWORD unused;
			FILE_FS_PERSISTENT_VOLUME_INFORMATION outBuffer;
			FILE_FS_PERSISTENT_VOLUME_INFORMATION inBuffer{};

			inBuffer.FlagMask = PERSISTENT_VOLUME_STATE_SHORT_NAME_CREATION_DISABLED;
			inBuffer.Version = 1;

			if (DeviceIoControl(hDevice, FSCTL_QUERY_PERSISTENT_VOLUME_STATE, &inBuffer, sizeof(inBuffer), &outBuffer, sizeof(outBuffer), &unused, nullptr))
			{
				result = outBuffer.VolumeFlags;
			}

			CloseHandle(hDevice);
		}

		return result;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::wstring GetVolumeGuidPath(LPCWSTR path, bool trailingBackslash)
	{
		WCHAR volumeMountPoint[MAX_PATH];
		if (GetVolumePathNameW(path, volumeMountPoint, _countof(volumeMountPoint)))
		{
			WCHAR volumeName[MAX_PATH];
			if (GetVolumeNameForVolumeMountPointW(volumeMountPoint, volumeName, _countof(volumeName)))
			{
				if (!trailingBackslash)
				{
					auto lastCharIndex{ wcslen(volumeName) - 1 };
					if (volumeName[lastCharIndex] == L'\\')
					{
						volumeName[lastCharIndex] = L'\0';
					}
				}

				return std::wstring{ volumeName };
			}
		}

		return std::wstring{};
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	StorageDeviceDescriptor::StorageDeviceDescriptor() : pDescriptor_{ nullptr }
	{
	}

	bool StorageDeviceDescriptor::Load(LPCWSTR pDeviceName)
	{
		auto succeeded{ false };

		auto hDevice{ CreateFileW(pDeviceName, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr) };
		if (hDevice != INVALID_HANDLE_VALUE)
		{
			DWORD unused;
			STORAGE_DESCRIPTOR_HEADER descriptorHeader;
			STORAGE_PROPERTY_QUERY propertyQuery{ StorageDeviceProperty, PropertyStandardQuery };

			if (DeviceIoControl(hDevice, IOCTL_STORAGE_QUERY_PROPERTY, &propertyQuery, sizeof(propertyQuery), &descriptorHeader, sizeof(descriptorHeader), &unused, nullptr))
			{
				descriptorBuffer_.reset(new BYTE[descriptorHeader.Size]);
				pDescriptor_ = reinterpret_cast<PSTORAGE_DEVICE_DESCRIPTOR>(descriptorBuffer_.get());

				succeeded = DeviceIoControl(hDevice, IOCTL_STORAGE_QUERY_PROPERTY, &propertyQuery, sizeof(propertyQuery), pDescriptor_, descriptorHeader.Size, &unused, nullptr);
			}

			CloseHandle(hDevice);
		}

		if (!succeeded)
		{
			Unload();
		}

		return succeeded;
	}

	void StorageDeviceDescriptor::Unload()
	{
		if (pDescriptor_)
		{
			descriptorBuffer_.reset();
			pDescriptor_ = nullptr;
		}
	}

	BYTE StorageDeviceDescriptor::DeviceType() const
	{
		return pDescriptor_ ? pDescriptor_->DeviceType : 0;
	}

	BYTE StorageDeviceDescriptor::DeviceTypeModifier() const
	{
		return pDescriptor_ ? pDescriptor_->DeviceTypeModifier : 0;
	}

	BOOLEAN StorageDeviceDescriptor::RemovableMedia() const
	{
		return pDescriptor_ ? pDescriptor_->RemovableMedia : 0;
	}

	BOOLEAN StorageDeviceDescriptor::CommandQueueing() const
	{
		return pDescriptor_ ? pDescriptor_->CommandQueueing : 0;
	}

	STORAGE_BUS_TYPE StorageDeviceDescriptor::BusType() const
	{
		return pDescriptor_ ? pDescriptor_->BusType : BusTypeUnknown;
	}

	std::wstring StorageDeviceDescriptor::VendorId() const
	{
		return GetStringData(pDescriptor_ ? pDescriptor_->VendorIdOffset : 0);
	}

	std::wstring StorageDeviceDescriptor::ProductId() const
	{
		return GetStringData(pDescriptor_ ? pDescriptor_->ProductIdOffset : 0);
	}

	std::wstring StorageDeviceDescriptor::ProductRevision() const
	{
		return GetStringData(pDescriptor_ ? pDescriptor_->ProductRevisionOffset : 0);
	}

	std::wstring StorageDeviceDescriptor::SerialNumber() const
	{
		// TODO: this is not the correct display of the serial number - apparently there is byte swapping to do.
		return GetStringData(pDescriptor_ ? pDescriptor_->SerialNumberOffset : 0);
	}

	std::vector<BYTE> StorageDeviceDescriptor::RawDeviceProperties() const
	{
		if (pDescriptor_ && pDescriptor_->RawPropertiesLength > 0)
		{
			return std::vector<BYTE>{pDescriptor_->RawDeviceProperties, pDescriptor_->RawDeviceProperties + pDescriptor_->RawPropertiesLength};
		}

		return std::vector<BYTE>{};
	}

	std::wstring StorageDeviceDescriptor::GetStringData(DWORD offset) const
	{
		if (offset != 0)
		{
			auto pStr{ reinterpret_cast<LPCSTR>((BYTE*)pDescriptor_ + offset) };
			return std::wstring{ pStr, pStr + strlen(pStr) };
		}

		return std::wstring{};
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	StorageDeviceNumber::StorageDeviceNumber() : sdn_{}
	{
	}

	bool StorageDeviceNumber::Load(LPCWSTR pDeviceName)
	{
		auto succeeded{ false };

		auto hDevice{ CreateFileW(pDeviceName, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr) };
		if (hDevice != INVALID_HANDLE_VALUE)
		{
			DWORD unused;

			succeeded = DeviceIoControl(hDevice, IOCTL_STORAGE_GET_DEVICE_NUMBER, nullptr, 0, &sdn_, sizeof(sdn_), &unused, nullptr);

			CloseHandle(hDevice);
		}

		if (!succeeded)
		{
			Unload();
		}

		return succeeded;
	}

	void StorageDeviceNumber::Unload()
	{
		memset(&sdn_, 0, sizeof(sdn_));
	}

	DEVICE_TYPE StorageDeviceNumber::DeviceType() const
	{
		return sdn_.DeviceType;
	}

	DWORD StorageDeviceNumber::DeviceNumber() const
	{
		return sdn_.DeviceNumber;
	}

	DWORD StorageDeviceNumber::PartitionNumber() const
	{
		return sdn_.PartitionNumber;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	VolumeDiskExtents::VolumeDiskExtents() : pVolDiskExt_{ nullptr }
	{
	}

	bool VolumeDiskExtents::Load(LPCWSTR pDeviceName)
	{
		auto succeeded{ false };

		auto hDevice{ CreateFileW(pDeviceName, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr) };
		if (hDevice != INVALID_HANDLE_VALUE)
		{
			DWORD unused;
			DWORD vdeBufferSize{ sizeof(VOLUME_DISK_EXTENTS) };
			CreateVolDiskExtBuffer(vdeBufferSize);

			succeeded = DeviceIoControl(hDevice, IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS, nullptr, 0, pVolDiskExt_, vdeBufferSize, &unused, nullptr);
			if (!succeeded && GetLastError() == ERROR_MORE_DATA)
			{
				vdeBufferSize = sizeof(VOLUME_DISK_EXTENTS) + sizeof(DISK_EXTENT) * pVolDiskExt_->NumberOfDiskExtents;
				CreateVolDiskExtBuffer(vdeBufferSize);

				succeeded = DeviceIoControl(hDevice, IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS, nullptr, 0, pVolDiskExt_, vdeBufferSize, &unused, nullptr);
			}

			CloseHandle(hDevice);
		}

		if (!succeeded)
		{
			Unload();
		}

		return succeeded;
	}

	void VolumeDiskExtents::Unload()
	{
		if (pVolDiskExt_)
		{
			volDiskExtBuffer_.reset();
			pVolDiskExt_ = nullptr;
		}
	}

	std::vector<DISK_EXTENT> VolumeDiskExtents::GetDiskExtents() const
	{
		std::vector<DISK_EXTENT> diskExtents{};

		if (pVolDiskExt_)
		{
			DWORD nDiskExtents{ pVolDiskExt_->NumberOfDiskExtents };
			for (DWORD i = 0; i < nDiskExtents; ++i)
			{
				diskExtents.push_back(pVolDiskExt_->Extents[i]);
			}
		}

		return diskExtents;
	}

	void VolumeDiskExtents::CreateVolDiskExtBuffer(DWORD size)
	{
		volDiskExtBuffer_.reset(new BYTE[size]);
		pVolDiskExt_ = reinterpret_cast<PVOLUME_DISK_EXTENTS>(volDiskExtBuffer_.get());
	}

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      environment
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

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
	//
	//                                      image
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

		// http://msdn.microsoft.com/en-us/library/bb757020.aspx
	HBITMAP BitmapFromIcon(HICON hIcon)
	{
		HBITMAP hBitmap{ nullptr };

		IWICImagingFactory* pWICImagingFactory;
		if (SUCCEEDED(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory))))
		{
			IWICBitmap* pWICBitmap;
			if (SUCCEEDED(pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmap)))
			{
				IWICFormatConverter* pWICFormatConverter;
				if (SUCCEEDED(pWICImagingFactory->CreateFormatConverter(&pWICFormatConverter)))
				{
					if (SUCCEEDED(pWICFormatConverter->Initialize(pWICBitmap, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, nullptr, 0.0f, WICBitmapPaletteTypeCustom)))
					{
						UINT cx, cy;
						if (SUCCEEDED(pWICFormatConverter->GetSize(&cx, &cy)))
						{
							HDC hDC{ GetDC(nullptr) };
							if (hDC != nullptr)
							{
								LPBYTE pBuffer;
								BITMAPINFO bmi{};
								bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
								bmi.bmiHeader.biPlanes = 1;
								bmi.bmiHeader.biCompression = BI_RGB;
								bmi.bmiHeader.biWidth = static_cast<LONG>(cx);
								bmi.bmiHeader.biHeight = -static_cast<LONG>(cy);
								bmi.bmiHeader.biBitCount = 32;
								hBitmap = CreateDIBSection(hDC, &bmi, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pBuffer), nullptr, 0);
								if (hBitmap != nullptr)
								{
									UINT cbStride{ cx * sizeof(ARGB) };
									UINT cbBuffer{ cy * cbStride };
									if (FAILED(pWICFormatConverter->CopyPixels(nullptr, cbStride, cbBuffer, pBuffer)))
									{
										hBitmap = nullptr;
									}
								}
								ReleaseDC(nullptr, hDC);
							}
						}
					}
					pWICFormatConverter->Release();
				}
				pWICBitmap->Release();
			}
			pWICImagingFactory->Release();
		}

		return hBitmap;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	HBITMAP BitmapFromIconResource(HMODULE hModule, WORD idIcon, int width, int height)
	{
		HBITMAP hBitmap{ nullptr };

		auto hIcon{ LoadIconResource(hModule, idIcon, width, height) };
		if (hIcon != nullptr)
		{
			hBitmap = BitmapFromIcon(hIcon);
			DestroyIcon(hIcon);
		}

		return hBitmap;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      menu
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

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
	//
	//                                      path
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

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
	//
	//                                      registry
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

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
			cbData = GetMultiSzSize(pValData);
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
	//
	//                                      resource
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

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
	//
	//                                      string
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

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

	bool IsMultiSzItems(LPCSTR pMultiSz)
	{
		if (pMultiSz != nullptr)
		{
			pMultiSz += strlen(pMultiSz) + 1;

			return *pMultiSz != 0;
		}

		return false;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	bool IsMultiSzItems(LPCWSTR pMultiSz)
	{
		if (pMultiSz != nullptr)
		{
			pMultiSz += wcslen(pMultiSz) + 1;

			return *pMultiSz != 0;
		}

		return false;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	size_t GetMultiSzCount(LPCSTR pMultiSz)
	{
		size_t count{ 0 };

		if (pMultiSz != nullptr)
		{
			while (*pMultiSz)
			{
				++count;
				pMultiSz += strlen(pMultiSz) + 1;
			}
		}

		return count;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	size_t GetMultiSzCount(LPCWSTR pMultiSz)
	{
		size_t count{ 0 };

		if (pMultiSz != nullptr)
		{
			while (*pMultiSz)
			{
				++count;
				pMultiSz += wcslen(pMultiSz) + 1;
			}
		}

		return count;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	size_t GetMultiSzSize(LPCSTR pMultiSz)
	{
		if (pMultiSz != nullptr)
		{
			auto p{ pMultiSz };
			while (*p)
			{
				p += strlen(p) + 1;
			}

			return ((p - pMultiSz) + 1) * sizeof(CHAR);
		}

		return 0;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	size_t GetMultiSzSize(LPCWSTR pMultiSz)
	{
		if (pMultiSz != nullptr)
		{
			auto p{ pMultiSz };
			while (*p)
			{
				p += wcslen(p) + 1;
			}

			return ((p - pMultiSz) + 1) * sizeof(WCHAR);
		}

		return 0;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::vector<std::string> GetMultiSzItems(LPCSTR pMultiSz)
	{
		std::vector<std::string> items{};

		if (pMultiSz != nullptr)
		{
			while (*pMultiSz)
			{
				items.push_back(std::string{ pMultiSz });
				pMultiSz += strlen(pMultiSz) + 1;
			}
		}

		return items;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::vector<std::wstring> GetMultiSzItemsWide(LPCSTR pMultiSz)
	{
		std::vector<std::wstring> items{};

		if (pMultiSz != nullptr)
		{
			while (*pMultiSz)
			{
				auto length{ strlen(pMultiSz) };
				items.push_back(std::wstring{ pMultiSz, pMultiSz + length });
				pMultiSz += length + 1;
			}
		}

		return items;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

	std::vector<std::wstring> GetMultiSzItems(LPCWSTR pMultiSz)
	{
		std::vector<std::wstring> items{};

		if (pMultiSz != nullptr)
		{
			while (*pMultiSz)
			{
				items.push_back(std::wstring{ pMultiSz });
				pMultiSz += wcslen(pMultiSz) + 1;
			}
		}

		return items;
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

	std::wstring WStrFromStr(LPCSTR pStr)
	{
		return std::wstring{ pStr, pStr + strlen(pStr) };
	}

	///////////////////////////////////////////////////////////////////////////////////////////////
	//
	//                                      synchronization
	//
	///////////////////////////////////////////////////////////////////////////////////////////////

	bool MutexExists(LPCWSTR pName)
	{
		auto mutex{ OpenMutexW(SYNCHRONIZE, FALSE, pName) };
		auto exists{ mutex != nullptr };
		if (exists)
		{
			CloseHandle(mutex);
		}

		return exists;
	}

	///////////////////////////////////////////////////////////////////////////////////////////////

}
