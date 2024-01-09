#include "SpeelTTraceLogging.h"
#include "gtest/gtest.h"
#include <windows.h>
#include <stdio.h>
#include "excel/Workbook.h"

TRACELOGGING_DEFINE_PROVIDER(g_hSpeelTLoggingProvider, "SpeelTLoggingProvider", (0x60f1baa2, 0x8f3f, 0x479b, 0xba, 0x22, 0xe, 0x4, 0xa7, 0xbb, 0x60, 0xe7));

TEST(SpeelT_App_OpcFactory, ReadFromOpcFactory) {
	TraceLoggingWrite(g_hSpeelTLoggingProvider, "SpeelT_App_OpcFactory", TraceLoggingWideString(L"Starting Test", "ReadFromOpcFactory"));

	WCHAR bindir[MAX_PATH];
	DWORD ncch = GetCurrentDirectoryW(MAX_PATH, bindir);

	if (!ncch)
	{
		LPWSTR buffer = nullptr;
		FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER, nullptr, GetLastError(), 0, reinterpret_cast<LPWSTR>(&buffer), 10, nullptr);

		std::wcout << "Error: " << buffer << std::endl;
	}

	ASSERT_GT(ncch, 0);

	std::wcout << L"Current directory is '" << bindir << "'\n";

	std::wstring speelt_tests_dir{ bindir };
	speelt_tests_dir += L"\\tests\\data\\";

	std::wstring sh_rolls_file = speelt_tests_dir + L"SH Rolls.xlsx";
	std::wstring struct_hedges_file = speelt_tests_dir + L"Structural Hedges.xlsm";

	SpeelT::Excel::Workbook sh_rolls_book(sh_rolls_file);
	SpeelT::Excel::Workbook struct_hedges_book(struct_hedges_file);

	ASSERT_EQ(sh_rolls_book.GetOPCPackage()->GetNumParts(), 110ull);
	ASSERT_EQ(sh_rolls_book.GetOPCPackage()->GetNumRelationships(), 4ul);
	ASSERT_EQ(sh_rolls_book.hasVBAProject(), false);

	ASSERT_EQ(struct_hedges_book.GetOPCPackage()->GetNumParts(), 40ull);
	ASSERT_EQ(struct_hedges_book.GetOPCPackage()->GetNumRelationships(), 3ul);
	ASSERT_EQ(struct_hedges_book.hasVBAProject(), true);

	TraceLoggingWrite(g_hSpeelTLoggingProvider, "SpeelT_App_OpcFactory", TraceLoggingWideString(L"Ending Test", "ReadFromOpcFactory"));
}

int main(int argc, char** argv) {
	::testing::InitGoogleTest(&argc, argv);

	sentry_options_t* options = sentry_options_new();
	sentry_options_set_debug(options, 1);
	sentry_init(options);

	TraceLoggingRegister(g_hSpeelTLoggingProvider);

	CoInitialize(nullptr);
	int retValue = RUN_ALL_TESTS();
	CoUninitialize();

	sentry_close();

	return retValue;
}