#include "SpeelTComUtils.h"
#include <string>

namespace SpeelT::Excel {
	class VBAProject {
	public:
		VBAProject();
		VBAProject(IOpcPartPtr vbaPart);
		bool exists();

		static std::wstring VBAContentType;
	private:
		bool _exists;
		std::wstring _streamName, _streamCLSID;
	};
}