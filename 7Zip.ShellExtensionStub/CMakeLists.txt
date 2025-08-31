cmake_minimum_required(VERSION 3.16)
project(SevenZipHost LANGUAGES CXX)
add_executable(7Zip.ShellExtensionStub WIN32 7Zip.ShellExtensionStub.cpp)
target_compile_features(7Zip.ShellExtensionStub PRIVATE cxx_std_17)

set_target_properties(7Zip.ShellExtensionStub PROPERTIES OUTPUT_NAME "7Zip.ShellExtensionStub")
