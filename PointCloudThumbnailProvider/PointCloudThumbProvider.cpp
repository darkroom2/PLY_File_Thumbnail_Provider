#include "happly.h"
#include <Shlwapi.h>
#include <thumbcache.h> // For IThumbnailProvider.
#include <gdiplus.h>
#include <iostream>
#include <limits>

// Include all GLM core / GLSL features
#include <glm/glm.hpp> // vec2, vec3, mat4, radians

// Include all GLM extensions
#include <glm/ext.hpp> // perspective, translate, rotate

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "gdiplus.lib")


// this thumbnail provider implements IInitializeWithStream to enable being hosted
// in an isolated process for robustness

class CPointCloudThumbProvider : public IInitializeWithStream, IThumbnailProvider {
public:
	CPointCloudThumbProvider() : _cRef(1), _pStream(nullptr) {}

	virtual ~CPointCloudThumbProvider() {
		if (_pStream) {
			_pStream->Release();
		}
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		static const QITAB qit[] = {
				QITABENT(CPointCloudThumbProvider, IInitializeWithStream),
				QITABENT(CPointCloudThumbProvider, IThumbnailProvider),
				{nullptr}
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef() override {
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release() override {
		ULONG cRef = InterlockedDecrement(&_cRef);
		if (!cRef) {
			delete this;
		}
		return cRef;
	}

	// IInitializeWithStream
	IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode) override;

	// IThumbnailProvider
	IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) override;

private:
	long _cRef;
	IStream* _pStream;     // provided during initialization.
};

HRESULT CPointCloudThumbProvider_CreateInstance(REFIID riid, void** ppv) {
	auto* pNew = new(std::nothrow) CPointCloudThumbProvider();
	HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr)) {
		hr = pNew->QueryInterface(riid, ppv);
		pNew->Release();
	}
	return hr;
}

// IInitializeWithStream
IFACEMETHODIMP CPointCloudThumbProvider::Initialize(IStream* pStream, DWORD) {
	HRESULT hr = E_UNEXPECTED;  // can only be inited once
	if (_pStream == nullptr) {
		// take a reference to the stream if we have not been inited yet
		hr = pStream->QueryInterface(&_pStream);
	}
	return hr;
}

// Funkcja dostarczaj¹ca odpowiedni¹ miniaturkê
IFACEMETHODIMP CPointCloudThumbProvider::GetThumbnail(UINT cx, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) {
	// wy³¹czenie kana³u alfa w miniaturce
	*pdwAlpha = WTSAT_RGB;

	// wczytanie pliku jako strumieñ
	STATSTG statstg;
	HRESULT hr = _pStream->Stat(&statstg, STATFLAG_NONAME);
	if (SUCCEEDED(hr)) {
		auto fileSize = statstg.cbSize.QuadPart;
		// utworzenie bufora na plik
		std::vector<char> buffer(fileSize + 1);
		// zasadnicze wczytanie strumienia do bufora
		hr = _pStream->Read(buffer.data(), fileSize, nullptr);
		if (SUCCEEDED(hr)) {
			// zamiana bufora na string oraz stringstream (wymóg biblioteki happly)
			std::string ret(buffer.data());
			std::stringstream ss;
			ss << ret;

			// inicjalizacja obiektu reprezentujacego dany plik .ply
			happly::PLYData plyIn(ss);

			// wektor trójek punktów (X, Y, Z)
			std::vector<std::array<double, 3>> vPos = plyIn.getVertexPositions();

			// inicjalizacja bitmapy z biblioteki Windows Gdiplus
			Gdiplus::GdiplusStartupInput startupInput;
			ULONG_PTR gdiplusToken;
			Gdiplus::GdiplusStartup(&gdiplusToken, &startupInput, nullptr);
			// cx jest rozmiarem bitmapy dostarczanym przez system w argumencie wywo³ania tej funkcji
			Gdiplus::Bitmap* gdiPlusBitmap = new Gdiplus::Bitmap(cx, cx, PixelFormat24bppRGB);

			// glm::perspective tworzy macierz projekcji perspektywicznej 4x4, która jest u¿ywana do przekszta³cania punktów.
			glm::mat4 Projection = glm::perspective(glm::radians(45.0f), 1.f, 0.1f, 10.0f);

			// znajdz max i min w tablicy wspolrzednych (wystarczy x, y i z)
			double max_x, max_y, max_z;
			max_x = max_y = max_z = std::numeric_limits<double>::min();
			double min_x, min_y, min_z;
			min_x = min_y = min_z = std::numeric_limits<double>::max();

			for (const auto& i : vPos) {
				if (max_x < i[0])	max_x = i[0];
				if (max_y < i[1])	max_y = i[1];
				if (max_z < i[2])	max_z = i[2];

				if (min_x > i[0])	min_x = i[0];
				if (min_y > i[1])	min_y = i[1];
				if (min_z > i[2])	min_z = i[2];
			}

			// macierz widoku kamery
			glm::mat4 View = glm::lookAt(
				glm::vec3(max_x * 2, max_y * 2, 2 * std::max({ max_x, max_y, max_z })), // punkt w którym znajduje siê kamera
				glm::vec3((min_x + max_x) / 2.0, (min_y + max_y) / 2.0, (min_z + max_z) / 2.0), // punkt na który patrzymy
				glm::vec3(0, 1, 0)  // wektor okreœlaj¹cym kierunek "w górê" œwiata (przestrzeni).
			);

			// wektor do przechowywania skonwertowanych punktów przestrzeni 3d do punktów 2d naszej bitmapy wraz z wartoœci¹ G i B koloru
			std::vector<glm::vec3> screen_coords;

			// dla ka¿dej trójki (X, Y, Z)
			for (auto& i : vPos) {

				auto old_range = max_z - min_z;
				auto new_range = 255 - 0;
				auto new_value = (i[2] - min_z) * new_range / old_range;

				// utwórz punkt 4d z ostatni¹ wspó³rzêdn¹ 1.
				glm::vec4 point4d(i[0], i[1], i[2], 1.0);

				// utworzenie punktu po zastosowaniu perspektywy i uwzglêdnieniu obserwatora
				glm::vec4 clipSpacePos = Projection * (View * point4d);

				// normalizacja punktu
				glm::vec3 ndcSpacePos = glm::vec3(clipSpacePos.x, clipSpacePos.y, clipSpacePos.z) / clipSpacePos.w;

				// przeskalowanie punktu do rozmiarow ekranu
				glm::vec3 screen(((ndcSpacePos.x + 1.0) / 2.0) * cx, ((1.0 - ndcSpacePos.y) / 2.0) * cx, new_value);

				// dodatnie punktu do wektora
				screen_coords.push_back(screen);
			}

			// rysowanie wszystkich punktów
			for (auto& i : screen_coords) {
				gdiPlusBitmap->SetPixel((int)(i.x), (int)(i.y), Gdiplus::Color::Color(126, (int)(i.z), (int)(i.z)));
			}

			// funkcja tworz¹ca bitmapê HBITMAP wymagan¹ przez funkjcê GetThumbnail, phbmp jest wskaŸnikiem do tej bitmapy
			Gdiplus::Status status = gdiPlusBitmap->GetHBITMAP(Gdiplus::Color::Black, phbmp);
			delete gdiPlusBitmap;
			if (status == Gdiplus::Status::Ok)
				hr = S_OK;
			else
				hr = E_FAIL;

			Gdiplus::GdiplusShutdown(gdiplusToken); // zamkniêcie biblioteki GDIPlus
			return hr;
		}
	}
	return E_FAIL;
}
