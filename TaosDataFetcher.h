#pragma once
#ifndef TAOSDATAFETCHER_H
#define TAOSDATAFETCHER_H

#include <map>
#include <vector>
#include <string>
#include "taosdbapi.h"
#include <QObject>
#include <QString>

class TaosDataFetcher
{
public:
    TaosDataFetcher();
    ~TaosDataFetcher();

    // 从地址字符串获取数据
    std::map<int64_t, std::vector<float>> fetchDataFromAddress(const std::string& address);

    // 解析地址字符串并获取数据
    std::map<int64_t, std::vector<float>> parseAndFetchData(const std::string& address);

    // 批量获取多个地址的数据
    std::map<std::string, std::map<int64_t, std::vector<float>>> fetchMultipleData(const std::vector<std::string>& addresses);

    // 工具方法：时间戳转字符串
    static std::string timestampToString(time_t timestamp);

private:
    // 解析地址字符串
    bool parseAddress(const std::string& address,
        std::vector<std::string>& ycnoList,
        std::string& startTime,
        std::string& endTime,
        int& interval);

    // 数据库API对象
    taosdbapi* tdb;
};

#endif // TAOSDATAFETCHER_H