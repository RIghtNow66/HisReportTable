#include "TaosDataFetcher.h"
#include <sstream>
#include <iomanip>
#include <ctime>
#include <iostream>
#include <algorithm>
#include <QMessageBox>
#include <qDebug>

TaosDataFetcher::TaosDataFetcher()
    : tdb(new taosdbapi())
{
}

TaosDataFetcher::~TaosDataFetcher()
{
    delete tdb;
}

std::map<int64_t, std::vector<float>> TaosDataFetcher::fetchDataFromAddress(const std::string& address)
{
    std::vector<std::string> ycnoList;
    std::string startTime;
    std::string endTime;
    int interval = 5; // 默认间隔

    if (!parseAddress(address, ycnoList, startTime, endTime, interval)) {
        throw std::runtime_error("地址解析失败: " + address);
    }

    if (ycnoList.empty()) {
        throw std::runtime_error("地址中未包含有效的YCNO编号");
    }

    try {
        // 使用带时间间隔的查询
        auto result = tdb->read(ycnoList, startTime, endTime, interval);
        //std::map<int64_t, vector<float>> result = tdb->read(ycnoList, startTime, endTime);
        if (result.empty())
        {
            QMessageBox::warning(NULL, "警告", "未获取到有效数据，请检查taos连接", QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes);
        }
        return tdb->read(ycnoList, startTime, endTime, interval);
    }
    catch (const std::exception& e) {
        throw std::runtime_error(std::string("数据查询失败: ") + e.what());
    }
}

std::map<int64_t, std::vector<float>> TaosDataFetcher::parseAndFetchData(const std::string& address)
{
    return fetchDataFromAddress(address);
}

std::map<std::string, std::map<int64_t, std::vector<float>>> TaosDataFetcher::fetchMultipleData(const std::vector<std::string>& addresses)
{
    std::map<std::string, std::map<int64_t, std::vector<float>>> result;

    for (const auto& address : addresses) {
        try {
            result[address] = fetchDataFromAddress(address);
        }
        catch (const std::exception& e) {
            std::cerr << "获取地址 " << address << " 的数据失败: " << e.what() << std::endl;
            // 可以选择继续处理其他地址或抛出异常
        }
    }

    return result;
}

bool TaosDataFetcher::parseAddress(const std::string& address,
    std::vector<std::string>& ycnoList,
    std::string& startTime,
    std::string& endTime,
    int& interval)
{
    // 清除之前的输出
    ycnoList.clear();
    startTime.clear();
    endTime.clear();
    interval = 5;

    // 简单的地址解析逻辑 - 根据您的实际地址格式调整
    // 示例地址格式: "YC001,YC002@2024-01-01 00:00:00~2024-01-02 23:59:59#10"

    size_t atPos = address.find('@');
    size_t tildePos = address.find('~');
    size_t hashPos = address.find('#');

    if (atPos == std::string::npos || tildePos == std::string::npos) {
        return false;
    }

    // 解析YCNO列表
    std::string ycnoPart = address.substr(0, atPos);
    std::istringstream ycnoStream(ycnoPart);
    std::string ycno;
    while (std::getline(ycnoStream, ycno, ',')) {
        if (!ycno.empty()) {
            ycnoList.push_back(ycno);
        }
    }

    // 解析时间范围
    if (hashPos != std::string::npos) {
        // 有时间间隔
        startTime = address.substr(atPos + 1, tildePos - atPos - 1);
        endTime = address.substr(tildePos + 1, hashPos - tildePos - 1);

        // 解析时间间隔
        try {
            interval = std::stoi(address.substr(hashPos + 1));
        }
        catch (...) {
            interval = 5; // 使用默认值
        }
    }
    else {
        // 没有时间间隔
        startTime = address.substr(atPos + 1, tildePos - atPos - 1); 
        endTime = address.substr(tildePos + 1);
    }

    return !ycnoList.empty() && !startTime.empty() && !endTime.empty();
}

std::string TaosDataFetcher::timestampToString(time_t timestamp)
{
    std::tm* timeinfo = std::localtime(&timestamp);
    std::stringstream ss;
    ss << std::put_time(timeinfo, "%Y-%m-%d %H:%M:%S");
    return ss.str();
}
