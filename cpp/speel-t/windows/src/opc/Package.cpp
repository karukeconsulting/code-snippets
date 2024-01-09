#include <iostream>
#include "opc/Package.h"
#include "SpeelTComUtils.h"
#include "SpeelTTraceLogging.h"

namespace SpeelT::OPC {
	Package::Package(std::wstring& filename) : _numParts{}, _numRelationships{}, _partMap{}, _emptyPartsVector{} {
		IOpcFactoryPtr factory;
		IOpcPackagePtr package;

		IOpcPartPtr part;
		IOpcPartUriPtr partUri;
		IOpcPartSetPtr partSet;
		IOpcPartEnumeratorPtr partEnumerator;

		IOpcRelationshipPtr rel;
		IOpcRelationshipSetPtr relationshipSet;
		IOpcRelationshipEnumeratorPtr relEnumerator;

		IStreamPtr sourceFileStream;

		BOOL hasNext = false;

		TraceLoggingActivity<g_hSpeelTLoggingProvider> packageActivity;
		TraceLoggingActivity<g_hSpeelTLoggingProvider> partsEnumerationActivity;

		//partsEnumerationActivity.SetRelatedActivity(packageActivity);

		TraceLoggingWriteStart(packageActivity, "OPC Package Processing");
		try {
			CHK_HR_THROW(factory.CreateInstance(CLSID_OpcFactory));
			TraceLoggingWriteTagged(packageActivity, "Read package from file", TraceLoggingValue(filename.c_str(), "filename"));

			CHK_HR_THROW(factory->CreateStreamOnFile(filename.c_str(), OPC_STREAM_IO_READ, NULL, 0, &sourceFileStream));
			CHK_HR_THROW(factory->ReadPackageFromStream(sourceFileStream.GetInterfacePtr(), OPC_CACHE_ON_ACCESS, &package));
						
			// Parts
			CHK_HR_THROW(package->GetPartSet(&partSet));
			CHK_HR_THROW(partSet->GetEnumerator(&partEnumerator));

			TraceLoggingWriteStart(partsEnumerationActivity, "Enumerating parts");
			hasNext = false;
			CHK_HR_THROW(partEnumerator->MoveNext(&hasNext));

			while (hasNext) {
				_numParts++;

				std::cout << "\nGetting Part #" << _numParts << "; ";
				CHK_HR_THROW(partEnumerator->GetCurrent(&part));

				CHK_HR_THROW(part->GetName(&partUri));

				_bstr_t uri;
				CHK_HR_THROW(partUri->GetAbsoluteUri(uri.GetAddress()));
				std::wcout << L"URI: " << (LPCWSTR)uri << "; ";

				LPWSTR contentType;

				// Content Type
				CHK_HR_THROW(part->GetContentType(&contentType));
				std::wcout << L"Content Type: ";
				if (contentType) {
					std::wcout << contentType;

					std::wstring ct(contentType);
					_partMap[ct].push_back(part);
					CoTaskMemFree(contentType);
				}
				else {
					std::wcout << L"UNDETERMINED";
				}

				CHK_HR_THROW(partEnumerator->MoveNext(&hasNext));
			}
			TraceLoggingWriteTagged(partsEnumerationActivity, "Parts Analysis", TraceLoggingValue(_numParts, "# package parts"));
			TraceLoggingWriteStop(partsEnumerationActivity, "Enumerating parts");

			// Relationships
			CHK_HR_THROW(package->GetRelationshipSet(&relationshipSet));
			CHK_HR_THROW(relationshipSet->GetEnumerator(&relEnumerator));

			hasNext = false;
			CHK_HR_THROW(relEnumerator->MoveNext(&hasNext));

			while (hasNext) {
				_numRelationships++;

				std::cout << "\nGetting relationship #" << _numRelationships << "; ";
				CHK_HR_THROW(relEnumerator->GetCurrent(&rel));

				LPWSTR relType;
				IOpcUriPtr sourceUri;
				IUriPtr targetUri;
				OPC_URI_TARGET_MODE targetMode;

				// Relationship Type
				CHK_HR_THROW(rel->GetRelationshipType(&relType));
				std::wcout << "Relationship Type: ";
				if (relType) {
					std::wcout << relType << "; ";
					CoTaskMemFree(relType);
				}
				else {
					std::wcout << L"UNDETERMINED; ";
				}

				// Source Uri
				CHK_HR_THROW(rel->GetSourceUri(&sourceUri));

				_bstr_t absoluteUri_s, displayUri_s, path_s;
				CHK_HR_THROW(sourceUri->GetAbsoluteUri(absoluteUri_s.GetAddress()));
				CHK_HR_THROW(sourceUri->GetDisplayUri(displayUri_s.GetAddress()));
				CHK_HR_THROW(sourceUri->GetPath(path_s.GetAddress()));

				std::wcout << L"Absolute Uri (Source): " << (LPCWSTR)absoluteUri_s << "; ";
				std::wcout << L"Display Uri (Source): " << (LPCWSTR)displayUri_s << "; ";
				std::wcout << L"Path (Source): " << (LPCWSTR)path_s << "; ";

				// Target Uri
				CHK_HR_THROW(rel->GetTargetUri(&targetUri));

				_bstr_t absoluteUri_t, displayUri_t, path_t;
				CHK_HR_THROW(targetUri->GetAbsoluteUri(absoluteUri_t.GetAddress()));
				CHK_HR_THROW(targetUri->GetDisplayUri(displayUri_t.GetAddress()));
				CHK_HR_THROW(targetUri->GetPath(path_t.GetAddress()));

				std::wcout << "Absolute Uri (Target): " << (LPCWSTR)absoluteUri_t << "; ";
				std::wcout << "Display Uri (Target): " << (LPCWSTR)displayUri_t << "; ";
				std::wcout << "Path (Target): " << (LPCWSTR)path_t << "; ";

				// Target Mode
				CHK_HR_THROW(rel->GetTargetMode(&targetMode));
				std::wcout << "Target Mode: ";
				if (targetMode == OPC_URI_TARGET_MODE_EXTERNAL) {
					std::wcout << "External; ";
				}
				else {
					std::wcout << "Internal; ";
				}

				CHK_HR_THROW(relEnumerator->MoveNext(&hasNext));
			}
		}
		catch (_com_error& e) {
			std::wcout << L"IOpcFactory processing error: " << e.ErrorMessage() << std::endl;
		}
		TraceLoggingWriteStop(packageActivity, "OPC Package Processing");
	}

	uint64_t Package::GetNumParts() {
		return _numParts;
	}

	uint64_t Package::GetNumRelationships() {
		return _numRelationships;
	}

	const std::vector<IOpcPartPtr>& Package::GetPartsByType(const std::wstring& typeName) {
		if (_partMap.contains(typeName)) {
			return _partMap.at(typeName);
		}
		else {
			return _emptyPartsVector;
		}
	}
}