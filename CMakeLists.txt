cmake_minimum_required(VERSION 3.16)
project(7ZipShellExtension LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Build the DLL
add_library(7Zip.ShellExtension SHARED 7Zip.ShellExtension.cpp)

# Compile flags/defines
target_compile_definitions(7Zip.ShellExtension PRIVATE UNICODE _UNICODE)

# Add include directories for 7-Zip core
target_include_directories(7Zip.ShellExtension PRIVATE
    "${CMAKE_SOURCE_DIR}/../SevenZip/CPP"
)

# Link against Windows system libraries
target_link_libraries(7Zip.ShellExtension PRIVATE Ole32 Shlwapi Comdlg32)

# Set DLL properties correctly
set_target_properties(7Zip.ShellExtension PROPERTIES
    OUTPUT_NAME "7Zip.ShellExtension"
    PREFIX ""
    WINDOWS_EXPORT_ALL_SYMBOLS ON
)
