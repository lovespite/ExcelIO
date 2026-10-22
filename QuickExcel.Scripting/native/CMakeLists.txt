cmake_minimum_required(VERSION 3.21)
project(quickexcel_js LANGUAGES C)

# Static CRT makes the produced DLL usable without a separate VC runtime install.
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
set(QJS_ENABLE_INSTALL OFF CACHE BOOL "" FORCE)
set(QJS_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
add_subdirectory(vendor/quickjs-0.17.0 EXCLUDE_FROM_ALL)
add_library(quickexcel-js SHARED quickexcel-js.c)
target_link_libraries(quickexcel-js PRIVATE qjs)
set_target_properties(quickexcel-js PROPERTIES C_STANDARD 11 C_STANDARD_REQUIRED ON)
target_compile_definitions(quickexcel-js PRIVATE WIN32_LEAN_AND_MEAN _CRT_SECURE_NO_WARNINGS)
