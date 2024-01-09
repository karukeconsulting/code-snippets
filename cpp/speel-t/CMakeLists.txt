cmake_minimum_required (VERSION 3.22)

cmake_policy(VERSION 3.22)
cmake_policy(SET CMP0126 NEW)

include(GoogleTest)

if(${WIN32})
	set(VCPKG_TARGET_TRIPLET "x64-windows")
	set(CXXFLAGS /std:c++20)
elseif(${UNIX})
	set(VCPKG_TARGET_TRIPLET "x64-linux")
endif()

project ("speelt-app")

enable_testing()

find_package(GTest CONFIG REQUIRED)
find_package(sentry CONFIG REQUIRED)
find_package(spdlog REQUIRED)

add_executable(speelt-app "./src/SpeelTApp.h")
add_executable(speelt-tests)
add_executable(speelt-tests-workbook)

target_link_libraries(speelt-app PRIVATE spdlog::spdlog_header_only)

target_link_libraries(speelt-tests PRIVATE GTest::gtest GTest::gtest_main)
target_link_libraries(speelt-tests-workbook PRIVATE GTest::gtest GTest::gtest_main)
target_link_libraries(speelt-tests-workbook PRIVATE sentry::sentry)

find_path(SENTRY_INCLUDE_DIR sentry.h)
include_directories(${SENTRY_INCLUDE_DIR})

include_directories("./src" "./include")

if(${WIN32})
	message(STATUS "Platform is Windows")
	
	include_directories("./windows/src" "./windows/include")

	target_sources(speelt-app PRIVATE "./windows/src/SpeelTApp.cpp" "./windows/src/SpeelTOpcFactory.cpp" "./windows/src/excel/Workbook.cpp" "windows/src/opc/Package.cpp" "windows/src/excel/VBAProject.cpp" "src/SpeelTUtils.h" "windows/src/SpeelTUtils.cpp")
	target_sources(speelt-tests PRIVATE "./windows/tests/SpeelTTests.cpp" "src/SpeelTUtils.h" "windows/src/SpeelTUtils.cpp")
	target_sources(speelt-tests-workbook PRIVATE "./windows/tests/SpeelTWorkbookTests.cpp" "./windows/src/excel/Workbook.cpp" "windows/src/opc/Package.cpp" "windows/src/excel/VBAProject.cpp" "src/SpeelTUtils.h" "windows/src/SpeelTUtils.cpp")

	target_compile_definitions(speelt-app PUBLIC UNICODE)
	target_compile_definitions(speelt-tests PUBLIC UNICODE)
	target_compile_definitions(speelt-tests-workbook PUBLIC UNICODE)

	target_compile_options(speelt-app PUBLIC /std:c++20)
	target_compile_options(speelt-tests PUBLIC /std:c++20)
	target_compile_options(speelt-tests-workbook PUBLIC /std:c++20)

	#set_property(TARGET BZip2 PROPERTY IMPORTED_LOCATION $ENV{VCPKG_INSTALLATION_ROOT}\\packages\\bzip2_x64-windows\\lib\\bz2.lib)
	#set_property(TARGET BZip2 PROPERTY IMPORTED_LOCATION_Debug $ENV{VCPKG_INSTALLATION_ROOT}\\packages\\bzip2_x64-windows\\debug\\lib\\bz2d.lib)
elseif(${UNIX})
	message(STATUS "Platform is Linux")

	include_directories("./linux/src" "./linux/include")

	target_sources(speelt-app PRIVATE "./linux/src/SpeelTApp.cpp")
	target_sources(speelt-tests PRIVATE "./linux/tests/SpeelTTests.cpp")

	set_property(TARGET BZip2 PROPERTY IMPORTED_LOCATION $ENV{VCPKG_INSTALLATION_ROOT}/packages/bzip2_x64-linux/lib/bz2.lib)
	set_property(TARGET BZip2 PROPERTY IMPORTED_LOCATION_Debug $ENV{VCPKG_INSTALLATION_ROOT}/packages/bzip2_x64-linux/debug/lib/bz2.lib)
endif()

add_custom_command(TARGET speelt-tests POST_BUILD COMMAND ${CMAKE_SOURCE_DIR}/../scripts/speelt_robocopy.bat ${CMAKE_SOURCE_DIR}/../tests ${CMAKE_BINARY_DIR}/tests /E)
add_custom_command(TARGET speelt-tests POST_BUILD COMMAND ${CMAKE_SOURCE_DIR}/../scripts/speelt_robocopy.bat ${CMAKE_SOURCE_DIR}/../scripts ${CMAKE_BINARY_DIR} /E)
add_custom_command(TARGET speelt-app PRE_BUILD COMMAND "${CMAKE_COMMAND}" -E copy_directory "$<TARGET_FILE_DIR:sentry_crashpad::crashpad_handler>" "$<TARGET_FILE_DIR:speelt-app>")

# Compiler Definitions

# Compiler Options

# Google Tests
gtest_discover_tests(speelt-tests)
gtest_discover_tests(speelt-tests-workbook)

# Other Tests
add_test(NAME WPRTest COMMAND ${CMAKE_BINARY_DIR}/speelt_wpr_tests.bat)