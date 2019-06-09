#include "pch.h"

///////////////////////////////////////////////////////////////////////////////////////////////////

typedef DWORD ARGB;

///////////////////////////////////////////////////////////////////////////////////////////////////

namespace hlp
{

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

}
