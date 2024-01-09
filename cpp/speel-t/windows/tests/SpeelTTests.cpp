#pragma once

#include "gtest/gtest.h"
#include <windows.h>
#include <stdio.h>
#include "msxml6.tlh"
#include "zip.h"

TEST(SpeelT_App_Characteristics, WindowsProcessCreation) {
	WCHAR bindir[MAX_PATH];
	DWORD ncch = GetCurrentDirectoryW(MAX_PATH, bindir);

	if (!ncch)
	{
		LPWSTR buffer = nullptr;
		FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER, nullptr, GetLastError(), 0, reinterpret_cast<LPWSTR>(&buffer), 10, nullptr);

		std::wcout << "Error: " << buffer << std::endl;
	}

	ASSERT_GT(ncch, 0);
  
	std::wcout << L"Current directory is '" << bindir << "'" << std::endl;

	std::wstring speelt_app_exe{ bindir };
	speelt_app_exe += L"\\speelt-app.exe";

	std::wcout << L"speelt-app executable: " << speelt_app_exe << std::endl;

	STARTUPINFOEXW startup_info{};
	startup_info.StartupInfo.cb = sizeof(STARTUPINFOEXW);


	PROCESS_INFORMATION process_info;

	BOOL b = CreateProcessW(
		/* [in, optional]      LPCWSTR               lpApplicationName */ speelt_app_exe.c_str(),
		/* [in, out, optional] LPWSTR                lpCommandLine */ nullptr,
		/* [in, optional]      LPSECURITY_ATTRIBUTES lpProcessAttributes */ nullptr,
		/* [in, optional]      LPSECURITY_ATTRIBUTES lpThreadAttributes */ nullptr,
		/* [in]                BOOL                  bInheritHandles */ FALSE,
		/* [in]                DWORD                 dwCreationFlags */ CREATE_UNICODE_ENVIRONMENT | EXTENDED_STARTUPINFO_PRESENT | NORMAL_PRIORITY_CLASS,
		/* [in, optional]      LPVOID                lpEnvironment */ nullptr,
		/* [in, optional]      LPCWSTR               lpCurrentDirectory */ nullptr,
		/* [in]                LPSTARTUPINFOW        lpStartupInfo */ reinterpret_cast<LPSTARTUPINFOW>(&startup_info),
		/* [out]               LPPROCESS_INFORMATION lpProcessInformation */ &process_info);

	EXPECT_TRUE(b);

	if (b)
	{
		WaitForSingleObject(process_info.hProcess, INFINITE);

		DWORD exitCode;
		if (GetExitCodeProcess(process_info.hProcess, &exitCode))
		{
			EXPECT_EQ(exitCode, 0);
		}

		CloseHandle(process_info.hThread);
		CloseHandle(process_info.hProcess);
	}
}

TEST(SpeelT_App_Zip, ReadArchive) {
	FILE* xlfile;
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
	std::wcout << L"Test Excel file is: " << sh_rolls_file << "\n";

	zip_source_t* src = nullptr;
	zip_t* za = nullptr;
	std::unique_ptr<zip_error_t> error = std::make_unique<zip_error>();

	zip_error_init(error.get());
	src = zip_source_win32w_create(sh_rolls_file.c_str(), 0, -1, error.get());

	if (!src) {
		if (error) {
			std::cout << "libzip error: " << zip_error_strerror(error.get()) << "\n";
		}
		else {
			std::cout << "Could not open the zip file, and no error message was provided!\n";
		}
	}

	ASSERT_TRUE(src != nullptr);

	za = zip_open_from_source(src, ZIP_CHECKCONS | ZIP_RDONLY, error.get());
	ASSERT_TRUE(za);

	zip_int64_t num_entries = zip_get_num_entries(za, ZIP_FL_UNCHANGED);
	ASSERT_EQ(num_entries, 157);

	int res = zip_close(za);
	if (res != 0) {
		zip_error_t* error = zip_get_error(za);
		if (error) {
			std::cout << "libzip error: " << zip_error_strerror(error) << std::endl;
		}
		else {
			std::cout << "Could not fetch libzip error string!!" << std::endl;
		}
	}

	zip_error_fini(error.get());
}
