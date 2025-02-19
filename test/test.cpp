#include <catch2/catch_test_macros.hpp>
#include "../src/drive_collector.hpp"
#include <string_view>

TEST_CASE( "Todo" ) {
    constexpr std::string_view mntPoint = "D";
    auto drive = findDrive(std::string(mntPoint));
    REQUIRE(drive != nullptr);
    REQUIRE(drive->getName() == std::string(mntPoint));

    // todo: test the drive class to return the correct name
}