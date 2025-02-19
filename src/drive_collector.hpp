#pragma once
#include <string>

class IDrive
{
  public:
//    void getDiskType();
    virtual std::string getName() = 0;
};

class WindowsDrive : public IDrive
{
  public:
    std::string getName() override
    {
        return "D";
    }
};

//TODO change void to ILaufwerk
IDrive* findDrive(const std::string& mountPoint){return new WindowsDrive{};}
