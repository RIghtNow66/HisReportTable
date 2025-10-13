#include "UniversalQueryEngine.h"
#include "platform/rdbapi.h"  // RTDB相关头文件
#include <QRegularExpression>
#include <QDebug>
#include <sstream>

UniversalQueryEngine& UniversalQueryEngine::instance() {
    static UniversalQueryEngine instance;
    return instance;
}

// --- 新的公共接口实现 ---
QHash<QString, QVariant> UniversalQueryEngine::queryValuesForBindingKeys(QList<QString> bindingKeys)
{
    QHash<QString, QVariant> finalResults;
    if (bindingKeys.isEmpty()) {
        return finalResults;
    }

    // 1. 解析输入，构建请求列表
    std::vector<RtdbRequest> requests;
    for (const QString& key : bindingKeys) {
        if (!key.startsWith("##")) continue;

        QString cleanKey = key.mid(2); // 去掉 "##"
        QStringList parts = cleanKey.split('/');

        if (parts.size() >= 3) {
            RtdbRequest req;
            req.originalKey = key;
            req.tableName = parts[0].toStdString();
            req.objName = parts[1].toStdString();
            req.fieldName = parts[2].toStdString();
            requests.push_back(req);
        }
        else {
            // 对于格式错误的Key，直接在结果中标记
            finalResults[key] = "Invalid Format";
        }
    }

    // 2. 执行查询
    if (!requests.empty()) {
        executeRtdQuery(requests, finalResults);
    }

    // 3. 返回填充好的结果
    return finalResults;
}


void UniversalQueryEngine::executeRtdQuery(
    const std::vector<RtdbRequest>& requests,
    QHash<QString, QVariant>& results // 直接填充结果
)
{
    int N = requests.size();
    if (N == 0) return;

    // 1. 准备RTDB查询的输入结构
    std::vector<RDB_FIELD_STRU> getfinfo(N);
    for (int i = 0; i < N; ++i) {
        // 使用C++11的安全函数 strcpy_s，或在确保缓冲区足够大的情况下使用 strcpy
        strncpy(getfinfo[i].tabname, requests[i].tableName.c_str(), sizeof(getfinfo[i].tabname) - 1);
        strncpy(getfinfo[i].objname, requests[i].objName.c_str(), sizeof(getfinfo[i].objname) - 1);
        strncpy(getfinfo[i].fldname, requests[i].fieldName.c_str(), sizeof(getfinfo[i].fldname) - 1);
        // 确保字符串以 null 结尾
        getfinfo[i].tabname[sizeof(getfinfo[i].tabname) - 1] = '\0';
        getfinfo[i].objname[sizeof(getfinfo[i].objname) - 1] = '\0';
        getfinfo[i].fldname[sizeof(getfinfo[i].fldname) - 1] = '\0';
    }

    // 2. 执行RTDB查询
    Rdb_QuickPolling rsp;
    Rdb_MultiTypeValue rmtv;
    int ret = rsp.RdbGetFieldValue(SYS_USER, "", N, getfinfo.data(), &rmtv);

    // 3. 处理查询结果
    if (ret <= 0) {
        qDebug() << "RdbGetFieldValue API failed, ret =" << ret;
        // 如果API调用失败，将所有请求标记为错误
        for (const auto& req : requests) {
            results[req.originalKey] = "Query Error";
        }
        return;
    }

    // 创建一个布尔向量来标记哪些请求已收到响应
    std::vector<bool> responded(N, false);

    for (int i = 0; i < ret; i++) {
        int parano = rmtv.RdbGetValOrderno(i); // 获取响应对应的原始请求索引

        if (parano >= 0 && parano < N) {
            // 获取对应的原始请求
            const RtdbRequest& originalRequest = requests[parano];

            // 从RTDB结果中获取值 (这里假设是int，您可以根据需要使用 RdbGetVal_double, RdbGetVal_string 等)
            int resultValue = rmtv.RdbGetVal_int(i);

            // 使用原始Key，将结果存入哈希表
            results[originalRequest.originalKey] = QVariant(resultValue);

            // 标记此请求已响应
            responded[parano] = true;
        }
    }

    // 4. 处理未收到响应的请求
    for (int i = 0; i < N; ++i) {
        if (!responded[i]) {
            results[requests[i].originalKey] = "No Response";
        }
    }
}