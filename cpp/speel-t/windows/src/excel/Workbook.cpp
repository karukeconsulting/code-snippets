#include <iostream>
#include "excel/Workbook.h"

namespace SpeelT::Excel {
	Workbook::Workbook(std::wstring& filename) {
		_opcPackage = std::make_unique<SpeelT::OPC::Package>(filename);

		std::vector<IOpcPartPtr> vbaPartsVec = _opcPackage->GetPartsByType(L"application/vnd.ms-office.vbaProject");
		if (vbaPartsVec.empty()) {
			_ourProject = std::make_unique<VBAProject>();
		}
		else {
			try {
				_ourProject = std::make_unique<VBAProject>(vbaPartsVec.at(0));
			}
			catch (_com_error& e) {
				std::wcout << "Could not create Workbook from " << filename << ": " << e.ErrorMessage() << std::endl;
			}
		}
	}

	bool Workbook::hasVBAProject() {
		return _ourProject->exists();
	}

	VBAProject* Workbook::GetVBAProject() {
		return _ourProject.get();
	}

	SpeelT::OPC::Package* Workbook::GetOPCPackage() {
		return _opcPackage.get();
	}
}