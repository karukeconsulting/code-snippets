#include <stdint.h>
#include <string>
#include <unordered_map>
#include <vector>
#include "SpeelTComUtils.h"

namespace SpeelT::OPC {
	class Package {
	public:
		Package(std::wstring& filename);
		uint64_t GetNumParts();
		uint64_t GetNumRelationships();
		const std::vector<IOpcPartPtr>& GetPartsByType(const std::wstring& typeName);

	private:
		uint64_t _numParts;
		uint64_t _numRelationships;
		std::unordered_map<std::wstring, std::vector<IOpcPartPtr>> _partMap;
		std::vector<IOpcPartPtr> _emptyPartsVector;
	};
}