#pragma once
#pragma once
#ifndef UNIVERSALQUERYENGINE_H
#define UNIVERSALQUERYENGINE_H

#include <map>
#include <vector>
#include <string>
#include <QVariant>
#include <QHash>

// 通用查询引擎
class  UniversalQueryEngine {
public:
    static UniversalQueryEngine& instance();

    // 执行RTDB查询
    QHash<QString, QVariant> queryValuesForBindingKeys(
        QList<QString> bindingKeys
    );

private:
    UniversalQueryEngine() = default;
    ~UniversalQueryEngine() = default;
    UniversalQueryEngine(const UniversalQueryEngine&) = delete;
    UniversalQueryEngine& operator=(const UniversalQueryEngine&) = delete;

    struct RtdbRequest
    {
        QString originalKey;
        std::string tableName;
        std::string objName;
        std::string fieldName;
    };


    // 解析描述字符串
    void executeRtdQuery(
        const std::vector<RtdbRequest>& requests,
        QHash<QString, QVariant>& results
    );

};

#endif // UNIVERSALQUERYENGINE_H