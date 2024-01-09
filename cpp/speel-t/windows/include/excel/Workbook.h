#include <string>
#include "excel/VBAProject.h"
#include "excel/Worksheet.h"
#include "opc/Package.h"

namespace SpeelT::Excel {
	class Workbook {
	public:
		Workbook(std::wstring& filename);
		bool hasVBAProject();
		VBAProject* GetVBAProject();
		SpeelT::OPC::Package* GetOPCPackage();

	private:
		std::unique_ptr<VBAProject> _ourProject;
		std::unique_ptr<SpeelT::OPC::Package> _opcPackage;
	};
}