#pragma once
#include <string>
#include <memory>

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
// IDrive* niemals dumme pointer, immer smart pointer!
std::shared_ptr<IDrive> findDrive(const std::string& mountPoint) {
    // return new WindowsDrive;
    return std::make_shared<WindowsDrive>();
}
