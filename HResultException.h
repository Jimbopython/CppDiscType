#pragma once

#include <Windows.h>
#include <stdexcept>

#include "utils.h"

class HResultException : public std::runtime_error
{
    HRESULT m_errorCode;

  public:
    HResultException(const std::string msg, HRESULT errorCode = -1)
        : std::runtime_error(msg), m_errorCode(errorCode)
    {
    }

    const char *what() const noexcept override
    {
        PrintHR(m_errorCode);
        return std::runtime_error::what();
    }
};