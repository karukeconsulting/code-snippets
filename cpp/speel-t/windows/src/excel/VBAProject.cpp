#include <atlbase.h>
#include "SpeelTApp.h"
#include "excel/VBAProject.h"
#include <vector>
#include <stdint.h>
#include <iostream>

namespace SpeelT::Excel {
	std::wstring VBAProject::VBAContentType = L"application/vnd.ms-office.vbaProject";

	VBAProject::VBAProject() : _exists{ false } {

	}

	VBAProject::VBAProject(IOpcPartPtr vbaPart) : _exists{ false } {
		IStreamPtr vbaStream;
		LPWSTR contentType;
		STATSTG stats;

		TraceLoggingWrite(g_hSpeelTLoggingProvider, "Test", TraceLoggingWideString(L"FieldValue", "Subtest"));

		// Get the IOpcPart's content type, and check that it is a VBA part
		CHK_HR_THROW(vbaPart->GetContentType(&contentType));
		std::wstring ct(contentType);

		if (ct != VBAContentType) {
			throw "";
		}

		// Get the VBA IOpcPart's stream
		CHK_HR_THROW(vbaPart->GetContentStream(&vbaStream));
		// Get the stream's information
		CHK_HR_THROW(vbaStream->Stat(&stats, STATFLAG_DEFAULT));

		// Save the stream's name
		if (stats.pwcsName) {
			_streamName = stats.pwcsName;
			CoTaskMemFree(stats.pwcsName);
		}

		// Save the stream's CLSID
		CComHeapPtr<wchar_t> clsid_str;
		CHK_HR_THROW(StringFromCLSID(stats.clsid, &clsid_str));

		if (clsid_str) {
			_streamCLSID = clsid_str;
		}

		ULONG cbRead{};
		ULONGLONG cbSize{ stats.cbSize.QuadPart };
		HRESULT hrRead = S_OK;

		ILockBytesPtr vbaLockBytes;
		CHK_HR_THROW(CreateILockBytesOnHGlobal(NULL, TRUE, &vbaLockBytes));
		vbaLockBytes->SetSize(stats.cbSize);

		std::vector<byte> streamData;
		streamData.reserve(cbSize);
		ULONG sliceSize{};

		if (cbSize > static_cast<ULONGLONG>(UINT32_MAX)) {
			sliceSize = UINT32_MAX;
		}
		else {
			sliceSize = cbSize;
		}
		std::vector<byte> streamSlice(static_cast<size_t>(sliceSize));
		bool more_data;

		// Loop through the stream's bytes, and store them in an ILockBytes object
		do {
			ULARGE_INTEGER ulOffset;
			ulOffset.QuadPart = 0ull;
			ULONG cbWritten{};

			hrRead = vbaStream->Read(streamSlice.data(), sliceSize, &cbRead);
			more_data = SUCCEEDED(hrRead) && (cbRead > 0);
			if (more_data) {
				CHK_HR_THROW(vbaLockBytes->WriteAt(ulOffset, streamSlice.data(), sliceSize, &cbWritten));
				ulOffset.QuadPart += static_cast<ULONGLONG>(cbWritten);
				streamData.insert(streamData.end(), streamSlice.begin(), streamSlice.begin() + cbRead);
			}
		} while (cbRead > 0);

		std::wcout << L"\ncbSize: " << cbSize << "; streamData.size(): " << streamData.size() << "; ";

		CHK_HR_THROW(vbaLockBytes->Stat(&stats, STATFLAG_NONAME));
		std::wcout << "vbaLockBytes->Stat(...).cbsize: " << stats.cbSize.QuadPart << "\n";

		// Open IStorage from ILockBytes
		IStoragePtr vbaStorage;
		HRESULT hrOpen = StgOpenStorageOnILockBytes(vbaLockBytes, NULL, STGM_READ | STGM_PRIORITY | STGM_SIMPLE, NULL, 0, &vbaStorage);
		if (SUCCEEDED(hrOpen)) {
			std::wcout << "Storage opened from ILockBytes\n";
		}
		else {
			std::wcout << "Storage opening from ILockBytes failed!! ; ";
			_com_error errOpen(hrOpen);
			if (errOpen.ErrorMessage()) {
				std::wcout << "ErrorMessage: " << errOpen.ErrorMessage() << " ; ";
			}
			//if (errCreate.Description().GetAddress()) {
			//	std::wcout << "Description: " << errCreate.Description() << " ; ";
			//}
			//if (errCreate.Source().GetAddress()) {
			//	std::wcout << "Source: " << errCreate.Source() << "\n";
			//}
			throw errOpen;
		}

		vbaStorage->Stat(&stats, STATFLAG_DEFAULT);
		std::wcout << "Stats retrieved from storage\n";

		if (stats.pwcsName) {
			std::wcout << L"VBA Storage name: " << stats.pwcsName << "; ";
			CoTaskMemFree(stats.pwcsName);
		}

		IEnumSTATSTGPtr stgEnum;
		CHK_HR_THROW(vbaStorage->EnumElements(0, NULL, 0, &stgEnum));

		STATSTG statsArray[25];
		ULONG celtFetched;

		do {
			CHK_HR_THROW(stgEnum->Next(25, statsArray, &celtFetched));

			if (celtFetched) {
				for (ULONG idx = 0; idx < celtFetched; idx++) {
					std::wcout << "Storage element #" << idx << ": " << statsArray[idx].pwcsName << "; ";
				}
			}
		} while (celtFetched > 0);
		std::wcout << "\n";

		_exists = true;
	}

	bool VBAProject::exists() {
		return _exists;
	}
}