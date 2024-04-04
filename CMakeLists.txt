cmake_minimum_required(VERSION 3.0)
project(CppDiscType)

add_executable(${PROJECT_NAME} main.cpp utils.cpp utils.h HResultException.h)

enable_testing()
add_subdirectory(tests)
