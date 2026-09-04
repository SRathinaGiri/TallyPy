#include "TallyService.h"

#include <QByteArray>
#include <QCryptographicHash>
#include <QDate>
#include <QDateTime>
#include <QEventLoop>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QMap>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QNetworkRequest>
#include <QStandardPaths>
#include <QThread>
#include <QRegularExpression>
#include <QSet>
#include <QStringConverter>
#include <QTextStream>
#include <QTimer>
#include <QUrl>
#include <QDomDocument>
#include <algorithm>
#include <atomic>
#include <cmath>
#include <functional>
#include <mutex>
#include <stdexcept>

namespace {
const QSet<QString> kAccountingVoucherTypes = {
    "Sales",
    "Purchase",
    "Journal",
    "Receipt",
    "Payment",
    "Debit Note",
    "Credit Note",
    "Contra",
    "Memorandum",
    "Reversing Journal",
};

const QSet<QString> kPredefinedVoucherTypes = {
    "Contra",
    "Payment",
    "Receipt",
    "Journal",
    "Sales",
    "Purchase",
    "Debit Note",
    "Credit Note",
    "Memorandum",
    "Reversing Journal",
    "Delivery Note",
    "Receipt Note",
    "Rejections In",
    "Rejections Out",
    "Stock Journal",
    "Physical Stock",
    "Material In",
    "Material Out",
    "Sales Order",
    "Purchase Order",
    "Job Work In Order",
    "Job Work Out Order",
    "Payroll",
    "Attendance",
};

const QSet<QString> kInventoryVoucherTypes = {
    "Delivery Note",
    "Receipt Note",
    "Rejections In",
    "Rejections Out",
    "Stock Journal",
    "Physical Stock",
    "Material In",
    "Material Out",
};

const QSet<QString> kOrderVoucherTypes = {
    "Sales Order",
    "Purchase Order",
    "Job Work In Order",
    "Job Work Out Order",
};

const QSet<QString> kPayrollVoucherTypes = {
    "Payroll",
    "Attendance",
};

struct ExportChunk {
    QString fromDate;
    QString toDate;
    int masterFrom = -1;
    int masterTo = -1;
    int expectedCount = -1;
};

std::atomic_bool gCancelRequested{false};
std::function<void(const QString &)> gLogCallback;
std::mutex gLogMutex;

void checkCancelled() {
    if (gCancelRequested.load()) {
        throw std::runtime_error("Operation cancelled by user.");
    }
}

void emitLog(const QString &message) {
    std::function<void(const QString &)> callback;
    {
        std::lock_guard<std::mutex> lock(gLogMutex);
        callback = gLogCallback;
    }
    if (callback) {
        callback(message);
    }
}

const QSet<QString> kBsPrimaryGroups = {
    "Capital Account", "Reserves & Surplus",
    "Loans (Liability)", "Bank OD A/c", "Secured Loans", "Unsecured Loans",
    "Current Liabilities", "Duties & Taxes", "Provisions", "Sundry Creditors",
    "Fixed Assets",
    "Investments",
    "Current Assets", "Stock-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Bank Accounts", "Cash-in-hand", "Sundry Debtors",
    "Misc. Expenses (ASSET)",
    "Suspense Account",
    "Branch / Divisions",
};

const QSet<QString> kPlPrimaryGroups = {
    "Sales Accounts",
    "Purchase Accounts",
    "Direct Incomes",
    "Indirect Incomes",
    "Direct Expenses",
    "Indirect Expenses",
};

const QSet<QString> kPrimaryGroups = []() {
    QSet<QString> groups = kBsPrimaryGroups;
    groups.unite(kPlPrimaryGroups);
    return groups;
}();

const QMap<QString, QString> kCurrencyFallbacks = {
    {"INR", QString::fromUtf8("₹")},
    {"INDIAN RUPEE", QString::fromUtf8("₹")},
    {"RUPEE", QString::fromUtf8("₹")},
    {"RUPEES", QString::fromUtf8("₹")},
    {"RS", QString::fromUtf8("₹")},
    {"RS.", QString::fromUtf8("₹")},
    {"USD", "$"},
    {"US DOLLAR", "$"},
    {"DOLLAR", "$"},
    {"EUR", QString::fromUtf8("€")},
    {"EURO", QString::fromUtf8("€")},
    {"GBP", QString::fromUtf8("£")},
    {"POUND", QString::fromUtf8("£")},
    {"POUND STERLING", QString::fromUtf8("£")},
    {"AED", QString::fromUtf8("د.إ")},
    {"DIRHAM", QString::fromUtf8("د.إ")},
    {"", ""},
};

const QStringList kVoucherColumns = {
    "Date", "VoucherTypeName", "BaseVoucherType", "VoucherNumber", "LedgerName",
    "MasterID", "Amount", "DrCr", "DebitAmount", "CreditAmount", "ParentLedger",
    "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "PartyLedgerName",
    "PartyGSTIN", "LedgerGSTIN", "VoucherNarration", "IsOptional", "CompanyName",
    "FromDate", "ToDate"
};

const QStringList kAllVoucherColumns = {
    "Date", "VoucherTypeName", "BaseVoucherType", "VoucherNumber", "LedgerName",
    "MasterID", "Amount", "DrCr", "DebitAmount", "CreditAmount", "ParentLedger",
    "PrimaryGroup", "Nature", "NatureOfGroup", "PAN", "PartyLedgerName",
    "PartyGSTIN", "LedgerGSTIN", "VoucherNarration", "IsOptional", "CompanyName",
    "FromDate", "ToDate", "VoucherCategory"
};

const QStringList kLedgerColumns = {
    "MasterID", "Name", "PrimaryGroup", "Nature", "NatureOfGroup", "PAN",
    "StartingFrom", "CurrencyName", "StateName", "Parent", "PartyGSTIN",
    "OpeningBalance", "ClosingBalance", "CompanyName", "FromDate", "ToDate"
};

const QStringList kStockItemColumns = {
    "Name", "Parent", "Category", "LedgerName", "OpeningBalance", "OpeningValue",
    "BasicValue", "BasicQty", "OpeningRate", "ClosingBalance", "ClosingValue",
    "ClosingRate", "CompanyName", "FromDate", "ToDate"
};

const QStringList kStockVoucherColumns = {
    "Date", "VoucherTypeName", "VoucherNumber", "StockItemName", "BilledQty",
    "Rate", "Amount", "GodownName", "BatchName", "VoucherNarration", "CompanyName",
    "FromDate", "ToDate"
};

struct GroupInfo {
    QString parent;
    QString nature;
    QString primaryGroup;
};

QString cleanText(const QString &text) {
    QString value = text;
    static const QRegularExpression controlChars("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]");
    value.remove(controlChars);
    return value.trimmed();
}

QString stripNs(const QString &tag) {
    const int idx = tag.indexOf('}');
    return idx >= 0 ? tag.mid(idx + 1) : tag;
}

QString xmlCleanup(QString xmlText) {
    static const QRegularExpression charRefs(R"(&#(x[0-9A-Fa-f]+|\d+);)");
    QString cleaned;
    cleaned.reserve(xmlText.size());
    int last = 0;
    auto it = charRefs.globalMatch(xmlText);
    while (it.hasNext()) {
        auto match = it.next();
        cleaned += xmlText.mid(last, match.capturedStart() - last);
        const QString value = match.captured(1);
        bool ok = false;
        uint codepoint = 0;
        if (value.startsWith('x', Qt::CaseInsensitive)) {
            codepoint = value.mid(1).toUInt(&ok, 16);
        } else {
            codepoint = value.toUInt(&ok, 10);
        }
        if (ok && (codepoint == 9 || codepoint == 10 || codepoint == 13 ||
                   (codepoint >= 32 && codepoint <= 55295) ||
                   (codepoint >= 57344 && codepoint <= 65533) ||
                   (codepoint >= 65536 && codepoint <= 1114111))) {
            cleaned += match.captured(0);
        }
        last = match.capturedEnd();
    }
    cleaned += xmlText.mid(last);
    xmlText = cleaned;

    static const QRegularExpression controlChars("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]");
    xmlText.remove(controlChars);
    xmlText.replace(QRegularExpression(R"(&(?!#\d+;|#x[0-9A-Fa-f]+;|[A-Za-z_:][A-Za-z0-9_.:-]*;))"), "&amp;");
    xmlText.replace(QRegularExpression(R"(<(/?)[A-Za-z_][\w.-]*:([A-Za-z_][\w.-]*))"), "<\\1\\2");
    xmlText.replace(QRegularExpression(R"(\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*"[^"]*")"), "");
    xmlText.replace(QRegularExpression(R"(\s+xmlns:[A-Za-z_][\w.-]*\s*=\s*'[^']*')"), "");
    return xmlText;
}

QDomDocument parseXmlRoot(const QString &xmlText) {
    QDomDocument doc;
    QString errorMsg;
    int errorLine = 0;
    int errorColumn = 0;
    if (!doc.setContent(xmlCleanup(xmlText), &errorMsg, &errorLine, &errorColumn)) {
        throw std::runtime_error(QString("XML parse error at line %1, column %2: %3")
                                     .arg(errorLine)
                                     .arg(errorColumn)
                                     .arg(errorMsg)
                                     .toStdString());
    }
    return doc;
}

QList<QDomElement> directChildren(const QDomElement &elem, const QString &localName) {
    QList<QDomElement> children;
    QDomNode node = elem.firstChild();
    const QString wanted = localName.toUpper();
    while (!node.isNull()) {
        if (node.isElement()) {
            QDomElement child = node.toElement();
            if (stripNs(child.tagName()).toUpper() == wanted) {
                children.append(child);
            }
        }
        node = node.nextSibling();
    }
    return children;
}

QString directChildText(const QDomElement &elem, const QString &localName) {
    QDomNode node = elem.firstChild();
    const QString wanted = localName.toUpper();
    while (!node.isNull()) {
        if (node.isElement()) {
            QDomElement child = node.toElement();
            if (stripNs(child.tagName()).toUpper() == wanted) {
                return cleanText(child.text());
            }
        }
        node = node.nextSibling();
    }
    return {};
}

bool isRealVoucher(const QDomElement &elem) {
    if (stripNs(elem.tagName()).toUpper() != "VOUCHER") {
        return false;
    }
    if (!elem.attribute("VCHKEY").isEmpty() || !elem.attribute("REMOTEID").isEmpty() || !elem.attribute("VCHTYPE").isEmpty()) {
        return true;
    }
    return !directChildText(elem, "DATE").isEmpty()
        || !directChildText(elem, "VOUCHERTYPENAME").isEmpty()
        || !directChildText(elem, "VOUCHERNUMBER").isEmpty()
        || !directChildText(elem, "MASTERID").isEmpty();
}

int voucherMasterId(const QDomElement &voucher) {
    bool ok = false;
    const int attrValue = cleanText(voucher.attribute("MASTERID")).toInt(&ok);
    if (ok) {
        return attrValue;
    }
    const int childValue = directChildText(voucher, "MASTERID").toInt(&ok);
    return ok ? childValue : -1;
}

int countRealVouchers(const QDomDocument &doc) {
    int count = 0;
    const QDomNodeList vouchers = doc.elementsByTagName("VOUCHER");
    for (int i = 0; i < vouchers.size(); ++i) {
        if (isRealVoucher(vouchers.at(i).toElement())) {
            ++count;
        }
    }
    return count;
}

QString firstDescendantText(const QDomElement &elem, const QString &localName) {
    const QString wanted = localName.toUpper();
    QDomNode node = elem.firstChild();
    while (!node.isNull()) {
        if (node.isElement()) {
            const QDomElement child = node.toElement();
            if (stripNs(child.tagName()).toUpper() == wanted) {
                const QString value = cleanText(child.text());
                if (!value.isEmpty()) {
                    return value;
                }
            }
            const QString nested = firstDescendantText(child, localName);
            if (!nested.isEmpty()) {
                return nested;
            }
        }
        node = node.nextSibling();
    }
    return {};
}

QString firstNonEmptyText(const QDomElement &elem, const QStringList &names) {
    for (const QString &name : names) {
        const QString value = directChildText(elem, name);
        if (!value.isEmpty()) {
            return value;
        }
    }
    return {};
}

QString formatTallyDate(const QString &value) {
    const QString text = cleanText(value);
    static const QRegularExpression eightDigits(R"(^\d{8}$)");
    if (eightDigits.match(text).hasMatch()) {
        return QString("%1-%2-%3").arg(text.mid(0, 4), text.mid(4, 2), text.mid(6, 2));
    }
    return text;
}

QString canonicalVoucherTypeName(const QString &value) {
    const QString cleaned = cleanText(value);
    const QString lowered = cleaned.toLower();
    if (lowered == "rejection in" || lowered == "rejections in") {
        return "Rejections In";
    }
    if (lowered == "rejection out" || lowered == "rejections out") {
        return "Rejections Out";
    }
    return cleaned;
}

QString voucherCategoryFromBaseType(const QString &baseType) {
    if (kAccountingVoucherTypes.contains(baseType)) {
        return "Accounting";
    }
    if (kInventoryVoucherTypes.contains(baseType)) {
        return "Inventory";
    }
    if (kOrderVoucherTypes.contains(baseType)) {
        return "Orders";
    }
    if (kPayrollVoucherTypes.contains(baseType)) {
        return "Payroll";
    }
    return "Unknown";
}

QString escapeXml(const QString &value) {
    QString escaped = value.toHtmlEscaped();
    escaped.replace('\'', "&apos;");
    return escaped;
}

QString normalizeAmountText(const QString &value) {
    QString text = cleanText(value);
    text.remove(',');
    if (text.isEmpty()) {
        return {};
    }
    static const QRegularExpression numberPattern(R"([-+]?\d+(?:\.\d+)?)");
    auto matches = numberPattern.globalMatch(text);
    QString token;
    while (matches.hasNext()) {
        token = matches.next().captured(0);
    }
    return token.isEmpty() ? text : token;
}

double toDoubleValue(const QString &value, double defaultValue = 0.0) {
    const QString normalized = normalizeAmountText(value);
    if (normalized.isEmpty()) {
        return defaultValue;
    }
    bool ok = false;
    const double result = normalized.toDouble(&ok);
    return ok ? result : defaultValue;
}

QString numberToString(double value) {
    if (std::abs(value) < 0.0000001) {
        value = 0.0;
    }
    return QString::number(value, 'f', 2);
}

QString csvEscape(const QString &value) {
    QString escaped = value;
    escaped.replace('"', "\"\"");
    if (escaped.contains(',') || escaped.contains('"') || escaped.contains('\n') || escaped.contains('\r')) {
        return "\"" + escaped + "\"";
    }
    return escaped;
}

QString makeSessionDir() {
    const QString root = QStandardPaths::writableLocation(QStandardPaths::TempLocation) + "/TallyQtExporter";
    const QString stamp = QDateTime::currentDateTimeUtc().toString("yyyyMMdd_HHmmss_zzz");
    const QString path = root + "/" + stamp;
    QDir().mkpath(path);
    return path;
}

class CsvTableWriter {
public:
    CsvTableWriter() = default;

    CsvTableWriter(const QString &path, const QStringList &columns)
        : path_(path), columns_(columns), file_(path) {
        QDir().mkpath(QFileInfo(path).absolutePath());
        if (!file_.open(QIODevice::WriteOnly | QIODevice::Text | QIODevice::Truncate)) {
            throw std::runtime_error(QString("Unable to create export cache file: %1").arg(path).toStdString());
        }
        stream_.setDevice(&file_);
        stream_.setEncoding(QStringConverter::Utf8);
        stream_ << QChar(0xFEFF);
        writeValues(columns_);
    }

    void appendRow(QVariantMap row, const QString &companyName, const QString &fromDate, const QString &toDate) {
        row.insert("CompanyName", companyName);
        row.insert("FromDate", formatTallyDate(fromDate));
        row.insert("ToDate", formatTallyDate(toDate));
        QStringList values;
        values.reserve(columns_.size());
        for (const QString &column : columns_) {
            values << row.value(column).toString();
        }
        writeValues(values);
        ++rowCount_;
    }

    void appendRows(const QVector<QVariantMap> &rows, const QString &companyName, const QString &fromDate, const QString &toDate) {
        for (const QVariantMap &row : rows) {
            appendRow(row, companyName, fromDate, toDate);
        }
    }

    void flush() {
        stream_.flush();
        file_.flush();
    }

    TallyTable table(const QString &id, const QString &title, const QString &fileName) const {
        TallyTable table;
        table.id = id;
        table.title = title;
        table.defaultFileName = fileName;
        table.columns = columns_;
        table.csvPath = path_;
        table.rowCount = rowCount_;
        return table;
    }

private:
    void writeValues(const QStringList &values) {
        QStringList escaped;
        escaped.reserve(values.size());
        for (const QString &value : values) {
            escaped << csvEscape(value);
        }
        stream_ << escaped.join(',') << "\n";
    }

    QString path_;
    QStringList columns_;
    QFile file_;
    QTextStream stream_;
    int rowCount_ = 0;
};

QPair<QString, QString> natureFromPrimaryGroup(const QString &primaryGroup) {
    const QString pg = cleanText(primaryGroup).toLower();
    if (QStringList({
            "current assets", "fixed assets", "investments", "misc. expenses (asset)",
            "bank accounts", "cash-in-hand", "deposits (asset)", "loans & advances (asset)",
            "stock-in-hand", "sundry debtors"
        }).contains(pg)) {
        return {"BS", "Assets"};
    }
    if (QStringList({
            "capital account", "current liabilities", "loans (liability)", "suspense account",
            "branch / divisions", "bank od a/c", "duties & taxes", "provisions",
            "reserves & surplus", "secured loans", "sundry creditors", "unsecured loans"
        }).contains(pg)) {
        return {"BS", "Liabilities"};
    }
    if (QStringList({"direct incomes", "indirect incomes", "sales accounts"}).contains(pg)) {
        return {"PL", "Income"};
    }
    if (QStringList({"direct expenses", "indirect expenses", "purchase accounts"}).contains(pg)) {
        return {"PL", "Expenses"};
    }
    return {"Unknown", "Unknown"};
}

QString ledgerPrimaryGroup(const QString &ledgerName, const QMap<QString, QVariantMap> &ledgerMeta) {
    QSet<QString> seen;
    QString current = cleanText(ledgerName);
    while (!current.isEmpty() && !seen.contains(current)) {
        seen.insert(current);
        const QVariantMap meta = ledgerMeta.value(current);
        const QString parent = cleanText(meta.value("Parent").toString());
        if (parent.isEmpty()) {
            return {};
        }
        if (kPrimaryGroups.contains(parent)) {
            return parent;
        }
        current = parent;
    }
    return {};
}

QString postToTally(const QString &url, const QString &xmlText, int timeoutMs = 120000) {
    checkCancelled();
    QNetworkAccessManager manager;
    QNetworkRequest request{QUrl(url)};
    request.setHeader(QNetworkRequest::ContentTypeHeader, "text/xml; charset=utf-8");

    QEventLoop loop;
    QNetworkReply *reply = manager.post(request, xmlText.toUtf8());
    QObject::connect(reply, &QNetworkReply::finished, &loop, &QEventLoop::quit);

    QTimer timer;
    timer.setSingleShot(true);
    QObject::connect(&timer, &QTimer::timeout, &loop, &QEventLoop::quit);
    timer.start(timeoutMs);

    QTimer cancelTimer;
    QObject::connect(&cancelTimer, &QTimer::timeout, &loop, [&]() {
        if (gCancelRequested.load()) {
            reply->abort();
            loop.quit();
        }
    });
    cancelTimer.start(100);
    loop.exec();
    cancelTimer.stop();

    checkCancelled();

    if (timer.isActive()) {
        timer.stop();
    } else {
        reply->abort();
        reply->deleteLater();
        throw std::runtime_error("Request to Tally timed out.");
    }

    const auto networkError = reply->error();
    const QString errorText = reply->errorString();
    const QByteArray body = reply->readAll();
    reply->deleteLater();

    if (networkError != QNetworkReply::NoError) {
        throw std::runtime_error(QString("Tally request failed: %1").arg(errorText).toStdString());
    }
    checkCancelled();
    return QString::fromUtf8(body);
}

QDate parseTallyDateValue(const QString &value) {
    const QString text = cleanText(value);
    if (text.isEmpty()) {
        return {};
    }
    const QStringList formats = {"yyyyMMdd", "yyyy-MM-dd", "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "d/M/yyyy", "dd-MMM-yyyy", "d-MMM-yyyy", "dd-MMMM-yyyy", "d-MMMM-yyyy"};
    for (const QString &format : formats) {
        const QDate parsed = QDate::fromString(text, format);
        if (parsed.isValid()) {
            return parsed;
        }
    }
    return {};
}

QString tallyRequestDate(const QString &value) {
    const QDate parsed = parseTallyDateValue(value);
    return parsed.isValid() ? parsed.toString("yyyyMMdd") : cleanText(value);
}

QVector<QPair<QString, QString>> splitPeriod(const QDate &startDate, const QDate &endDate, const QString &mode) {
    QVector<QPair<QString, QString>> chunks;
    QDate current = startDate;
    while (current <= endDate) {
        QDate chunkEnd;
        if (mode == "weekly") {
            chunkEnd = std::min(endDate, current.addDays(6));
        } else if (mode == "daily") {
            chunkEnd = current;
        } else {
            chunkEnd = std::min(endDate, QDate(current.year(), current.month(), current.daysInMonth()));
        }
        chunks.append({current.toString("yyyyMMdd"), chunkEnd.toString("yyyyMMdd")});
        current = chunkEnd.addDays(1);
    }
    return chunks;
}

QString buildVoucherProbeRequestXml(const QString &company, const QString &fromDate, const QString &toDate) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    staticVars << QString("<SVFROMDATE TYPE='Date'>%1</SVFROMDATE>").arg(escapeXml(tallyRequestDate(fromDate)));
    staticVars << QString("<SVTODATE TYPE='Date'>%1</SVTODATE>").arg(escapeXml(tallyRequestDate(toDate)));
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyVoucherProbe</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE><COLLECTION NAME=\"MyVoucherProbe\"><TYPE>Voucher</TYPE>"
        "<FETCH>Date, MasterID, VoucherTypeName</FETCH>"
        "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QVector<ExportChunk> buildMasterIdChunks(const QVector<int> &masterIds, const QDate &startDate, const QDate &endDate, int maxVouchers) {
    QVector<int> ids = masterIds;
    std::sort(ids.begin(), ids.end());
    QVector<ExportChunk> chunks;
    for (int index = 0; index < ids.size(); index += maxVouchers) {
        const int endIndex = std::min(index + maxVouchers, static_cast<int>(ids.size()));
        ExportChunk chunk;
        chunk.fromDate = startDate.toString("yyyyMMdd");
        chunk.toDate = endDate.toString("yyyyMMdd");
        chunk.masterFrom = ids[index];
        chunk.masterTo = ids[endIndex - 1];
        chunk.expectedCount = endIndex - index;
        chunks.append(chunk);
    }
    return chunks;
}

QVector<ExportChunk> planExportChunks(const QString &url, const QString &company, const QString &fromDate, const QString &toDate) {
    const QDate startDate = parseTallyDateValue(fromDate);
    const QDate endDate = parseTallyDateValue(toDate);
    if (!startDate.isValid() || !endDate.isValid() || startDate > endDate) {
        return {{cleanText(fromDate), cleanText(toDate), -1, -1, -1}};
    }

    QString mode = "monthly";
    try {
        emitLog(QString("Probe started for %1 to %2").arg(formatTallyDate(fromDate), formatTallyDate(toDate)));
        const QDomDocument doc = parseXmlRoot(postToTally(url, buildVoucherProbeRequestXml(company, fromDate, toDate), 180000));
        const QDomNodeList vouchers = doc.elementsByTagName("VOUCHER");
        QMap<QString, int> monthCounts;
        QVector<int> masterIds;
        int realVoucherCount = 0;
        for (int i = 0; i < vouchers.size(); ++i) {
            const QDomElement voucher = vouchers.at(i).toElement();
            if (!isRealVoucher(voucher)) {
                continue;
            }
            ++realVoucherCount;
            const QDate voucherDate = parseTallyDateValue(directChildText(voucher, "DATE"));
            if (voucherDate.isValid()) {
                const QString key = QString("%1-%2").arg(voucherDate.year()).arg(voucherDate.month());
                monthCounts[key] = monthCounts.value(key) + 1;
            }
            const int masterId = voucherMasterId(voucher);
            if (masterId >= 0) {
                masterIds.append(masterId);
            }
        }
        emitLog(QString("Probe completed. Voucher headers found: %1. MasterIDs found: %2.").arg(realVoucherCount).arg(masterIds.size()));
        const int target = std::max(250, qEnvironmentVariableIntValue("TALLYXML_CHUNK_VOUCHERS") > 0 ? qEnvironmentVariableIntValue("TALLYXML_CHUNK_VOUCHERS") : 1000);
        if (realVoucherCount <= target) {
            emitLog(QString("Chunk plan: one full-period request because probe count is <= %1.").arg(target));
            return {{tallyRequestDate(fromDate), tallyRequestDate(toDate), -1, -1, realVoucherCount}};
        }
        if (!masterIds.isEmpty()) {
            QVector<ExportChunk> chunks = buildMasterIdChunks(masterIds, startDate, endDate, target);
            emitLog(QString("Chunk plan: %1 MasterID chunk(s), target <= %2 probed vouchers each.").arg(chunks.size()).arg(target));
            return chunks;
        }
        int maxMonthCount = 0;
        for (int count : monthCounts) {
            maxMonthCount = std::max(maxMonthCount, count);
        }
        mode = maxMonthCount <= 2000 ? "monthly" : "weekly";
    } catch (...) {
        mode = "monthly";
        emitLog("Probe failed. Falling back to monthly chunks.");
    }

    QVector<ExportChunk> chunks;
    for (const auto &period : splitPeriod(startDate, endDate, mode)) {
        chunks.append({period.first, period.second, -1, -1, -1});
    }
    emitLog(QString("Chunk plan: %1 %2 chunk(s).").arg(chunks.size()).arg(mode));
    return chunks;
}

QString cacheFilePath(const QString &company, const QString &tableName, const ExportChunk &chunk) {
    QString cacheRoot = QStandardPaths::writableLocation(QStandardPaths::CacheLocation);
    if (cacheRoot.isEmpty()) {
        cacheRoot = QDir::currentPath() + "/.tally_cache";
    } else {
        cacheRoot += "/tally_xml";
    }
    const QString keyText = cleanText(company) + "|" + tableName + "|" + cleanText(chunk.fromDate) + "|" + cleanText(chunk.toDate)
        + "|" + QString::number(chunk.masterFrom) + "|" + QString::number(chunk.masterTo) + "|xml-v6";
    const QString fileName = QString(QCryptographicHash::hash(keyText.toUtf8(), QCryptographicHash::Sha1).toHex()) + ".xml";
    return QDir(cacheRoot).filePath(fileName);
}

QString fetchXmlCached(const QString &url, const QString &xmlText, const QString &company, const QString &tableName,
                       const ExportChunk &chunk) {
    const QString path = cacheFilePath(company, tableName, chunk);
    QFile existing(path);
    if (existing.exists() && existing.open(QIODevice::ReadOnly | QIODevice::Text)) {
        return QString::fromUtf8(existing.readAll());
    }

    std::exception_ptr lastError;
    for (int attempt = 0; attempt < 3; ++attempt) {
        try {
            const QString xml = postToTally(url, xmlText);
            QDir().mkpath(QFileInfo(path).absolutePath());
            QFile file(path);
            if (file.open(QIODevice::WriteOnly | QIODevice::Text)) {
                file.write(xml.toUtf8());
            }
            return xml;
        } catch (...) {
            lastError = std::current_exception();
            QThread::sleep(1 + attempt * 2);
        }
    }
    if (lastError) {
        std::rethrow_exception(lastError);
    }
    return {};
}

QVector<QVariantMap> fetchChunkedRows(
    const QString &url,
    const QString &company,
    const QString &fromDate,
    const QString &toDate,
    const QString &tableName,
    const std::function<QString(const QString &, const QString &, const QString &, int, int)> &buildRequest,
    const std::function<QVector<QVariantMap>(const QDomDocument &, const QString &, const QString &)> &parseChunk) {
    QVector<QVariantMap> rows;
    std::function<void(const ExportChunk &)> processChunk;
    int chunkIndex = 0;
    const QVector<ExportChunk> chunks = planExportChunks(url, company, fromDate, toDate);
    emitLog(QString("%1: starting %2 chunk(s).").arg(tableName, QString::number(chunks.size())));
    processChunk = [&](const ExportChunk &chunk) {
        ++chunkIndex;
        const QString chunkFrom = chunk.fromDate;
        const QString chunkTo = chunk.toDate;
        const bool logChunk = chunks.size() <= 20 || chunkIndex <= 5 || chunkIndex == chunks.size() || chunkIndex % 50 == 0;
        try {
            if (logChunk) {
                QString label = QString("%1/%2").arg(chunkIndex).arg(chunks.size());
                if (chunk.masterFrom >= 0) {
                    emitLog(QString("%1: chunk %2 started, MasterID %3-%4, expected vouchers %5.")
                                .arg(tableName, label)
                                .arg(chunk.masterFrom)
                                .arg(chunk.masterTo)
                                .arg(chunk.expectedCount));
                } else {
                    emitLog(QString("%1: chunk %2 started, %3 to %4.").arg(tableName, label, formatTallyDate(chunkFrom), formatTallyDate(chunkTo)));
                }
            }
            const QString xml = fetchXmlCached(url, buildRequest(company, chunkFrom, chunkTo, chunk.masterFrom, chunk.masterTo), company, tableName, chunk);
            const QDomDocument doc = parseXmlRoot(xml);
            const QString status = cleanText(firstDescendantText(doc.documentElement(), "STATUS"));
            if (status == "0") {
                const QString errorText = firstDescendantText(doc.documentElement(), "LINEERROR");
                throw std::runtime_error((errorText.isEmpty() ? QString("Tally returned STATUS=0") : errorText).toStdString());
            }
            const int beforeRows = rows.size();
            const QVector<QVariantMap> chunkRows = parseChunk(doc, chunkFrom, chunkTo);
            if (chunk.masterFrom >= 0 && chunk.expectedCount >= 0) {
                const int detailCount = countRealVouchers(doc);
                if (detailCount > 0 && detailCount < chunk.expectedCount) {
                    throw std::runtime_error(QString("MasterID-filtered response returned %1 voucher header(s), expected %2.")
                                                 .arg(detailCount)
                                                 .arg(chunk.expectedCount)
                                                 .toStdString());
                }
            }
            rows += chunkRows;
            if (logChunk) {
                emitLog(QString("%1: chunk parsed %2 row(s) from %3 voucher header(s). Total rows: %4.")
                            .arg(tableName)
                            .arg(rows.size() - beforeRows)
                            .arg(countRealVouchers(doc))
                            .arg(rows.size()));
            }
        } catch (...) {
            if (chunk.masterFrom >= 0 && chunk.masterTo >= 0) {
                if (chunk.masterTo - chunk.masterFrom <= 5) {
                    throw;
                }
                const int midpoint = (chunk.masterFrom + chunk.masterTo) / 2;
                ExportChunk left = chunk;
                left.masterTo = midpoint;
                left.expectedCount = -1;
                ExportChunk right = chunk;
                right.masterFrom = midpoint + 1;
                right.expectedCount = -1;
                emitLog(QString("%1: splitting MasterID range %2-%3 into smaller ranges.").arg(tableName).arg(chunk.masterFrom).arg(chunk.masterTo));
                processChunk(left);
                processChunk(right);
                return;
            }
            const QDate startDate = parseTallyDateValue(chunkFrom);
            const QDate endDate = parseTallyDateValue(chunkTo);
            if (!startDate.isValid() || !endDate.isValid() || startDate >= endDate) {
                throw;
            }
            for (const auto &day : splitPeriod(startDate, endDate, "daily")) {
                ExportChunk dayChunk{day.first, day.second, -1, -1, -1};
                const QString xml = fetchXmlCached(url, buildRequest(company, day.first, day.second, -1, -1), company, tableName, dayChunk);
                rows += parseChunk(parseXmlRoot(xml), day.first, day.second);
            }
        }
    };
    for (const auto &chunk : chunks) {
        processChunk(chunk);
    }
    emitLog(QString("%1: completed with %2 row(s).").arg(tableName).arg(rows.size()));
    return rows;
}

int fetchChunkedRowsStream(
    const QString &url,
    const QString &company,
    const QString &fromDate,
    const QString &toDate,
    const QString &tableName,
    const std::function<QString(const QString &, const QString &, const QString &, int, int)> &buildRequest,
    const std::function<QVector<QVariantMap>(const QDomDocument &, const QString &, const QString &)> &parseChunk,
    const std::function<void(const QVector<QVariantMap> &)> &handleRows) {
    int totalRows = 0;
    std::function<void(const ExportChunk &)> processChunk;
    int chunkIndex = 0;
    const QVector<ExportChunk> chunks = planExportChunks(url, company, fromDate, toDate);
    emitLog(QString("%1: starting %2 chunk(s).").arg(tableName, QString::number(chunks.size())));
    processChunk = [&](const ExportChunk &chunk) {
        ++chunkIndex;
        const QString chunkFrom = chunk.fromDate;
        const QString chunkTo = chunk.toDate;
        const bool logChunk = chunks.size() <= 20 || chunkIndex <= 5 || chunkIndex == chunks.size() || chunkIndex % 50 == 0;
        try {
            if (logChunk) {
                const QString label = QString("%1/%2").arg(chunkIndex).arg(chunks.size());
                if (chunk.masterFrom >= 0) {
                    emitLog(QString("%1: chunk %2 started, MasterID %3-%4, expected vouchers %5.")
                                .arg(tableName, label)
                                .arg(chunk.masterFrom)
                                .arg(chunk.masterTo)
                                .arg(chunk.expectedCount));
                } else {
                    emitLog(QString("%1: chunk %2 started, %3 to %4.").arg(tableName, label, formatTallyDate(chunkFrom), formatTallyDate(chunkTo)));
                }
            }
            const QString xml = fetchXmlCached(url, buildRequest(company, chunkFrom, chunkTo, chunk.masterFrom, chunk.masterTo), company, tableName, chunk);
            const QDomDocument doc = parseXmlRoot(xml);
            const QString status = cleanText(firstDescendantText(doc.documentElement(), "STATUS"));
            if (status == "0") {
                const QString errorText = firstDescendantText(doc.documentElement(), "LINEERROR");
                throw std::runtime_error((errorText.isEmpty() ? QString("Tally returned STATUS=0") : errorText).toStdString());
            }
            const QVector<QVariantMap> chunkRows = parseChunk(doc, chunkFrom, chunkTo);
            if (chunk.masterFrom >= 0 && chunk.expectedCount >= 0) {
                const int detailCount = countRealVouchers(doc);
                if (detailCount > 0 && detailCount < chunk.expectedCount) {
                    throw std::runtime_error(QString("MasterID-filtered response returned %1 voucher header(s), expected %2.")
                                                 .arg(detailCount)
                                                 .arg(chunk.expectedCount)
                                                 .toStdString());
                }
            }
            handleRows(chunkRows);
            totalRows += chunkRows.size();
            if (logChunk) {
                emitLog(QString("%1: chunk parsed %2 row(s) from %3 voucher header(s). Total rows: %4.")
                            .arg(tableName)
                            .arg(chunkRows.size())
                            .arg(countRealVouchers(doc))
                            .arg(totalRows));
            }
        } catch (...) {
            if (chunk.masterFrom >= 0 && chunk.masterTo >= 0) {
                if (chunk.masterTo - chunk.masterFrom <= 5) {
                    throw;
                }
                const int midpoint = (chunk.masterFrom + chunk.masterTo) / 2;
                ExportChunk left = chunk;
                left.masterTo = midpoint;
                left.expectedCount = -1;
                ExportChunk right = chunk;
                right.masterFrom = midpoint + 1;
                right.expectedCount = -1;
                emitLog(QString("%1: splitting MasterID range %2-%3 into smaller ranges.").arg(tableName).arg(chunk.masterFrom).arg(chunk.masterTo));
                processChunk(left);
                processChunk(right);
                return;
            }
            const QDate startDate = parseTallyDateValue(chunkFrom);
            const QDate endDate = parseTallyDateValue(chunkTo);
            if (!startDate.isValid() || !endDate.isValid() || startDate >= endDate) {
                throw;
            }
            for (const auto &day : splitPeriod(startDate, endDate, "daily")) {
                ExportChunk dayChunk{day.first, day.second, -1, -1, -1};
                const QString xml = fetchXmlCached(url, buildRequest(company, day.first, day.second, -1, -1), company, tableName, dayChunk);
                const QVector<QVariantMap> dayRows = parseChunk(parseXmlRoot(xml), day.first, day.second);
                handleRows(dayRows);
                totalRows += dayRows.size();
            }
        }
    };
    for (const auto &chunk : chunks) {
        processChunk(chunk);
    }
    emitLog(QString("%1: completed with %2 row(s).").arg(tableName).arg(totalRows));
    return totalRows;
}

QString buildLedgerRequestXml(const QString &company) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyLedgers</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyLedgers\"><TYPE>Ledger</TYPE>"
        "<FETCH>Name, Parent, PartyGSTIN, MasterID, StartingFrom, CurrencyName, StateName, OpeningBalance, ClosingBalance, IncomeTaxNumber</FETCH>"
        "<COMPUTE>PrimaryGroup:$_PrimaryGroup</COMPUTE>"
        "<COMPUTE>CurrencyFormalName:$FormalName:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencySymbol:$UnicodeSymbol:Currency:$CurrencyName</COMPUTE>"
        "<COMPUTE>CurrencyOriginalSymbol:$OriginalSymbol:Currency:$CurrencyName</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QPair<QString, QString> masterIdFilterXml(int masterFrom, int masterTo) {
    if (masterFrom < 0 || masterTo < 0) {
        return {"", ""};
    }
    return {
        "<FILTER>MasterIdRange</FILTER>",
        QString("<SYSTEM TYPE=\"Formulae\" NAME=\"MasterIdRange\">$MasterID &gt;= %1 AND $MasterID &lt;= %2</SYSTEM>").arg(masterFrom).arg(masterTo)
    };
}

QString buildFlatVoucherRequestXml(const QString &company, const QString &fromDate, const QString &toDate, int masterFrom, int masterTo) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    staticVars << QString("<SVFROMDATE TYPE='Date'>%1</SVFROMDATE>").arg(escapeXml(tallyRequestDate(fromDate)));
    staticVars << QString("<SVTODATE TYPE='Date'>%1</SVTODATE>").arg(escapeXml(tallyRequestDate(toDate)));
    const auto filter = masterIdFilterXml(masterFrom, masterTo);
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>TXMLFlatVoucherRows</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"TXMLBaseVouchers\"><TYPE>Voucher</TYPE>"
        + filter.first +
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration, PartyLedgerName, PartyGSTIN, IsOptional</FETCH>"
        "</COLLECTION>"
        "<COLLECTION NAME=\"TXMLFlatVoucherRows\"><SOURCECOLLECTION>TXMLBaseVouchers</SOURCECOLLECTION>"
        "<WALK>All Ledger Entries</WALK>"
        "<FETCH>LedgerName, Amount, IsDeemedPositive, TXMLDate, TXMLVoucherTypeName, TXMLVoucherNumber, "
        "TXMLPartyLedgerName, TXMLPartyGSTIN, TXMLVoucherNarration, TXMLSignedAmount, TXMLDebitAmount, "
        "TXMLCreditAmount, TXMLDrCr, TXMLEntryLedgerMasterID, TXMLEntryParentLedger, TXMLEntryPrimaryGroup, "
        "TXMLEntryLedgerGSTIN, TXMLStatusOptional, TXMLCompanyName</FETCH>"
        "<COMPUTE>TXMLDate:$..Date</COMPUTE>"
        "<COMPUTE>TXMLVoucherTypeName:$..VoucherTypeName</COMPUTE>"
        "<COMPUTE>TXMLVoucherNumber:$..VoucherNumber</COMPUTE>"
        "<COMPUTE>TXMLPartyLedgerName:If NOT $$IsEmpty:$..PartyLedgerName Then $..PartyLedgerName Else \"N/A\"</COMPUTE>"
        "<COMPUTE>TXMLPartyGSTIN:$..PartyGSTIN</COMPUTE>"
        "<COMPUTE>TXMLVoucherNarration:$..Narration</COMPUTE>"
        "<COMPUTE>TXMLSignedAmount:If ($IsDeemedPositive OR $Amount < 0) Then $$Abs:$Amount * -1 Else $$Abs:$Amount</COMPUTE>"
        "<COMPUTE>TXMLDebitAmount:If ($IsDeemedPositive OR $Amount < 0) Then $$Abs:$Amount Else 0</COMPUTE>"
        "<COMPUTE>TXMLCreditAmount:If ($IsDeemedPositive OR $Amount < 0) Then 0 Else $$Abs:$Amount</COMPUTE>"
        "<COMPUTE>TXMLDrCr:If ($IsDeemedPositive OR $Amount < 0) Then \"Dr\" Else \"Cr\"</COMPUTE>"
        "<COMPUTE>TXMLEntryLedgerMasterID:$MasterID:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLEntryParentLedger:$Parent:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLEntryPrimaryGroup:$_PrimaryGroup:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLEntryLedgerGSTIN:$PartyGSTIN:Ledger:$LedgerName</COMPUTE>"
        "<COMPUTE>TXMLStatusOptional:If $..IsOptional Then \"Yes\" Else \"No\"</COMPUTE>"
        "<COMPUTE>TXMLCompanyName:##SVCurrentCompany</COMPUTE>"
        "</COLLECTION>"
        + filter.second +
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QString buildStockItemRequestXml(const QString &company) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyStockItems</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyStockItems\"><TYPE>StockItem</TYPE>"
        "<FETCH>Name, Parent, Category, LedgerName, OpeningBalance, OpeningValue, BasicValue, BasicQty, OpeningRate</FETCH>"
        "<COMPUTE>ClosingBalance:$_ClosingBalance</COMPUTE>"
        "<COMPUTE>ClosingValue:$_ClosingValue</COMPUTE>"
        "<COMPUTE>ClosingRate:$_ClosingRate</COMPUTE>"
        "</COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QString buildFlatInventoryEntriesRequestXml(const QString &company, const QString &fromDate, const QString &toDate, int masterFrom, int masterTo) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }
    staticVars << QString("<SVFROMDATE TYPE='Date'>%1</SVFROMDATE>").arg(escapeXml(tallyRequestDate(fromDate)));
    staticVars << QString("<SVTODATE TYPE='Date'>%1</SVTODATE>").arg(escapeXml(tallyRequestDate(toDate)));
    const auto filter = masterIdFilterXml(masterFrom, masterTo);
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>TXMLFlatInventoryRows</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"TXMLBaseInventoryVouchers\"><TYPE>Voucher</TYPE>"
        + filter.first +
        "<FETCH>Date, VoucherTypeName, VoucherNumber, Narration</FETCH>"
        "</COLLECTION>"
        "<COLLECTION NAME=\"TXMLFlatInventoryRows\"><SOURCECOLLECTION>TXMLBaseInventoryVouchers</SOURCECOLLECTION>"
        "<WALK>All Inventory Entries</WALK>"
        "<FETCH>StockItemName, BilledQty, Rate, Amount, TXMLDate, TXMLVoucherTypeName, TXMLVoucherNumber, "
        "TXMLVoucherNarration, TXMLCompanyName, TXMLSignedQty, TXMLSignedAmount, TXMLGodownName, TXMLBatchName</FETCH>"
        "<COMPUTE>TXMLDate:$..Date</COMPUTE>"
        "<COMPUTE>TXMLVoucherTypeName:$..VoucherTypeName</COMPUTE>"
        "<COMPUTE>TXMLVoucherNumber:$..VoucherNumber</COMPUTE>"
        "<COMPUTE>TXMLVoucherNarration:$..Narration</COMPUTE>"
        "<COMPUTE>TXMLCompanyName:##SVCurrentCompany</COMPUTE>"
        "<COMPUTE>TXMLSignedQty:If $IsDeemedPositive Then $$Abs:$BilledQty Else $$Abs:$BilledQty * -1</COMPUTE>"
        "<COMPUTE>TXMLSignedAmount:If $IsDeemedPositive Then $$Abs:$Amount Else $$Abs:$Amount * -1</COMPUTE>"
        "<COMPUTE>TXMLGodownName:$GodownName</COMPUTE>"
        "<COMPUTE>TXMLBatchName:$BatchName</COMPUTE>"
        "</COLLECTION>"
        + filter.second +
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    );
}

QString buildCompanyListRequestXml() {
    return (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>List of Companies</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "</DESC></BODY></ENVELOPE>"
    );
}

QPair<QMap<QString, QString>, QMap<QString, GroupInfo>> fetchTallyMetadata(const QString &url, const QString &company) {
    QStringList staticVars = {"<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"};
    if (!company.isEmpty()) {
        staticVars << QString("<SVCURRENTCOMPANY>%1</SVCURRENTCOMPANY>").arg(escapeXml(company));
    }

    const QString vtypeXml =
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>AllVoucherTypes</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllVoucherTypes\"><TYPE>VoucherType</TYPE><FETCH>Name, Parent</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>";

    const QString groupXml =
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>AllGroups</ID></HEADER><BODY><DESC>"
        + QString("<STATICVARIABLES>%1</STATICVARIABLES>").arg(staticVars.join(""))
        + "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"AllGroups\"><TYPE>Group</TYPE><FETCH>Name, Parent, Nature, _PrimaryGroup</FETCH></COLLECTION>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>";

    QMap<QString, QString> vtypeMap;
    QMap<QString, GroupInfo> groupMap;

    try {
        const QDomDocument vtypeDoc = parseXmlRoot(postToTally(url, vtypeXml));
        QDomNodeList vtypes = vtypeDoc.elementsByTagName("VOUCHERTYPE");
        for (int i = 0; i < vtypes.size(); ++i) {
            const QDomElement elem = vtypes.at(i).toElement();
            const QString name = canonicalVoucherTypeName(cleanText(elem.attribute("NAME").isEmpty() ? directChildText(elem, "NAME") : elem.attribute("NAME")));
            const QString parent = canonicalVoucherTypeName(directChildText(elem, "PARENT"));
            if (!name.isEmpty()) {
                vtypeMap.insert(name, parent.isEmpty() ? name : parent);
            }
        }

        const QDomDocument groupDoc = parseXmlRoot(postToTally(url, groupXml));
        QDomNodeList groups = groupDoc.elementsByTagName("GROUP");
        for (int i = 0; i < groups.size(); ++i) {
            const QDomElement elem = groups.at(i).toElement();
            const QString name = directChildText(elem, "NAME");
            if (!name.isEmpty()) {
                groupMap.insert(name, {
                    directChildText(elem, "PARENT"),
                    directChildText(elem, "NATURE"),
                    directChildText(elem, "_PRIMARYGROUP")
                });
            }
        }

        for (int pass = 0; pass < 5; ++pass) {
            for (auto it = vtypeMap.begin(); it != vtypeMap.end(); ++it) {
                const QString parentName = it.value();
                if (!parentName.isEmpty() && !kPredefinedVoucherTypes.contains(parentName) && vtypeMap.contains(parentName)) {
                    it.value() = vtypeMap.value(parentName);
                }
            }
        }

        for (int pass = 0; pass < 5; ++pass) {
            for (auto it = groupMap.begin(); it != groupMap.end(); ++it) {
                GroupInfo &info = it.value();
                if (!info.parent.isEmpty() && groupMap.contains(info.parent)) {
                    if (info.nature.isEmpty()) {
                        info.nature = groupMap.value(info.parent).nature;
                    }
                    if (info.primaryGroup.isEmpty()) {
                        info.primaryGroup = groupMap.value(info.parent).primaryGroup;
                    }
                }
            }
        }
    } catch (...) {
    }

    return {vtypeMap, groupMap};
}

QVector<QVariantMap> parseLedgers(const QDomDocument &doc, const QMap<QString, GroupInfo> &groupMap) {
    QVector<QVariantMap> rows;
    QMap<QString, QVariantMap> ledgerLookup;
    QDomNodeList nodes = doc.elementsByTagName("LEDGER");

    for (int i = 0; i < nodes.size(); ++i) {
        const QDomElement elem = nodes.at(i).toElement();
        if (stripNs(elem.tagName()).toUpper() != "LEDGER") {
            continue;
        }

        const QString name = cleanText(elem.attribute("NAME").isEmpty() ? directChildText(elem, "NAME") : elem.attribute("NAME"));
        if (name.isEmpty()) {
            continue;
        }

        const QString parent = directChildText(elem, "PARENT");
        const GroupInfo gInfo = groupMap.value(parent);
        const QString primaryGroup = !gInfo.primaryGroup.isEmpty() ? gInfo.primaryGroup :
                                     firstNonEmptyText(elem, {"PRIMARYGROUP"}).isEmpty() ? firstDescendantText(elem, "PRIMARYGROUP") :
                                                                                           firstNonEmptyText(elem, {"PRIMARYGROUP"});

        QVariantMap row;
        row.insert("MasterID", cleanText(elem.attribute("MASTERID").isEmpty() ? directChildText(elem, "MASTERID") : elem.attribute("MASTERID")));
        row.insert("Name", name);
        row.insert("PrimaryGroup", primaryGroup);
        row.insert("Nature", "");
        row.insert("NatureOfGroup", gInfo.nature);
        row.insert("PAN", firstNonEmptyText(elem, {"INCOMETAXNUMBER", "PAN"}).isEmpty() ? firstDescendantText(elem, "INCOMETAXNUMBER") : firstNonEmptyText(elem, {"INCOMETAXNUMBER", "PAN"}));
        row.insert("StartingFrom", firstNonEmptyText(elem, {"STARTINGFROM"}).isEmpty() ? firstDescendantText(elem, "STARTINGFROM") : firstNonEmptyText(elem, {"STARTINGFROM"}));
        row.insert("CurrencyNameRaw", firstNonEmptyText(elem, {"CURRENCYNAME"}).isEmpty() ? firstDescendantText(elem, "CURRENCYNAME") : firstNonEmptyText(elem, {"CURRENCYNAME"}));
        row.insert("CurrencyFormalNameRaw", firstNonEmptyText(elem, {"CURRENCYFORMALNAME"}).isEmpty() ? firstDescendantText(elem, "CURRENCYFORMALNAME") : firstNonEmptyText(elem, {"CURRENCYFORMALNAME"}));
        row.insert("CurrencySymbolRaw", firstNonEmptyText(elem, {"CURRENCYSYMBOL"}).isEmpty() ? firstDescendantText(elem, "CURRENCYSYMBOL") : firstNonEmptyText(elem, {"CURRENCYSYMBOL"}));
        row.insert("CurrencyOriginalSymbolRaw", firstNonEmptyText(elem, {"CURRENCYORIGINALSYMBOL"}).isEmpty() ? firstDescendantText(elem, "CURRENCYORIGINALSYMBOL") : firstNonEmptyText(elem, {"CURRENCYORIGINALSYMBOL"}));
        row.insert("StateName", firstNonEmptyText(elem, {"STATENAME"}).isEmpty() ? firstDescendantText(elem, "STATENAME") : firstNonEmptyText(elem, {"STATENAME"}));
        row.insert("Parent", parent);
        row.insert("PartyGSTIN", firstNonEmptyText(elem, {"PARTYGSTIN", "GSTIN"}).isEmpty() ? firstDescendantText(elem, "PARTYGSTIN") : firstNonEmptyText(elem, {"PARTYGSTIN", "GSTIN"}));
        row.insert("OpeningBalance", numberToString(toDoubleValue(firstNonEmptyText(elem, {"OPENINGBALANCE"}).isEmpty() ? firstDescendantText(elem, "OPENINGBALANCE") : firstNonEmptyText(elem, {"OPENINGBALANCE"}))));
        row.insert("ClosingBalance", numberToString(toDoubleValue(firstNonEmptyText(elem, {"CLOSINGBALANCE"}).isEmpty() ? firstDescendantText(elem, "CLOSINGBALANCE") : firstNonEmptyText(elem, {"CLOSINGBALANCE"}))));
        rows.append(row);
        ledgerLookup.insert(name, row);
    }

    for (QVariantMap &row : rows) {
        QString primaryGroup = row.value("PrimaryGroup").toString();
        if (primaryGroup.isEmpty()) {
            primaryGroup = ledgerPrimaryGroup(row.value("Name").toString(), ledgerLookup);
            row.insert("PrimaryGroup", primaryGroup);
        }

        if (row.value("NatureOfGroup").toString().isEmpty() && groupMap.contains(primaryGroup)) {
            row.insert("NatureOfGroup", groupMap.value(primaryGroup).nature);
        }

        QString nature = row.value("Nature").toString();
        QString natureOfGroup = row.value("NatureOfGroup").toString();
        if (!natureOfGroup.isEmpty()) {
            const QString lowered = natureOfGroup.toLower();
            if (lowered == "assets" || lowered == "liabilities") {
                nature = "BS";
            } else if (lowered == "income" || lowered == "expenses") {
                nature = "PL";
            }
        }
        if (nature.isEmpty() && !primaryGroup.isEmpty()) {
            const auto pair = natureFromPrimaryGroup(primaryGroup);
            nature = pair.first;
            natureOfGroup = pair.second;
        }
        row.insert("Nature", nature);
        row.insert("NatureOfGroup", natureOfGroup);

        const QString currencyKey = cleanText((row.value("CurrencyFormalNameRaw").toString().isEmpty()
                                               ? row.value("CurrencyNameRaw").toString()
                                               : row.value("CurrencyFormalNameRaw").toString())).toUpper();
        row.insert("CurrencyName", kCurrencyFallbacks.value(currencyKey,
                                                            cleanText(row.value("CurrencySymbolRaw").toString().isEmpty()
                                                                          ? row.value("CurrencyOriginalSymbolRaw").toString()
                                                                          : row.value("CurrencySymbolRaw").toString())));
    }

    std::sort(rows.begin(), rows.end(), [](const QVariantMap &a, const QVariantMap &b) {
        return a.value("MasterID").toInt() == b.value("MasterID").toInt()
                   ? a.value("Name").toString() < b.value("Name").toString()
                   : a.value("MasterID").toInt() < b.value("MasterID").toInt();
    });

    QVector<QVariantMap> output;
    for (const QVariantMap &row : rows) {
        QVariantMap ordered;
        for (const QString &column : kLedgerColumns) {
            ordered.insert(column, row.value(column).toString());
        }
        output.append(ordered);
    }
    return output;
}

QVector<QVariantMap> parseFlatVouchers(const QDomDocument &doc, const QMap<QString, QVariantMap> &ledgerMeta,
                                       const QString &company, const QString &fromDate, const QString &toDate,
                                       const QMap<QString, QString> &vtypeMap) {
    QVector<QVariantMap> rows;
    const QString formattedFromDate = formatTallyDate(fromDate);
    const QString formattedToDate = formatTallyDate(toDate);
    QDomNodeList flatNodes = doc.elementsByTagName("LEDGERENTRY");
    if (!flatNodes.isEmpty()) {
        for (int i = 0; i < flatNodes.size(); ++i) {
            const QDomElement entry = flatNodes.at(i).toElement();
            const QString ledgerName = directChildText(entry, "LEDGERNAME");
            const double rawAmount = toDoubleValue(directChildText(entry, "AMOUNT"));
            double amountValue = toDoubleValue(firstNonEmptyText(entry, {"TXMLSIGNEDAMOUNT", "SIGNEDAMOUNT", "AMOUNT"}));
            if (ledgerName.isEmpty() || (std::abs(amountValue) < 0.0000001 && std::abs(rawAmount) < 0.0000001)) {
                continue;
            }
            const double baseAmount = std::abs(std::abs(rawAmount) >= 0.0000001 ? rawAmount : amountValue);
            const bool isDebit = directChildText(entry, "ISDEEMEDPOSITIVE").toUpper() == "YES" || rawAmount < 0;
            const double signedAmount = isDebit ? -baseAmount : baseAmount;

            const QString voucherType = canonicalVoucherTypeName(firstNonEmptyText(entry, {"TXMLVOUCHERTYPENAME", "VOUCHERTYPENAME"}));
            const QString baseType = canonicalVoucherTypeName(vtypeMap.value(voucherType, voucherType));
            const QString voucherCategory = voucherCategoryFromBaseType(baseType);
            const double debitAmount = signedAmount < 0 ? baseAmount : 0.0;
            const double creditAmount = signedAmount > 0 ? baseAmount : 0.0;

            QVariantMap meta = ledgerMeta.value(ledgerName);
            QString primaryGroup = firstNonEmptyText(entry, {"TXMLENTRYPRIMARYGROUP", "ENTRYPRIMARYGROUP", "PRIMARYGROUP"});
            if (primaryGroup.isEmpty()) primaryGroup = meta.value("PrimaryGroup").toString();
            QString nature = meta.value("Nature").toString();
            QString natureOfGroup = meta.value("NatureOfGroup").toString();
            if (nature.isEmpty() && !primaryGroup.isEmpty()) {
                const auto pair = natureFromPrimaryGroup(primaryGroup);
                nature = pair.first;
                natureOfGroup = pair.second;
            }

            QString optional = firstNonEmptyText(entry, {"TXMLSTATUSOPTIONAL", "STATUSOPTIONAL", "ISOPTIONAL"});
            if (optional.toUpper() == "YES") {
                optional = "Yes";
            } else if (optional.toUpper() == "NO" || optional.isEmpty()) {
                optional = "No";
            }

            QVariantMap row;
            row.insert("Date", formatTallyDate(firstNonEmptyText(entry, {"TXMLDATE", "DATE"})));
            row.insert("VoucherTypeName", voucherType);
            row.insert("BaseVoucherType", baseType);
            row.insert("VoucherNumber", firstNonEmptyText(entry, {"TXMLVOUCHERNUMBER", "VOUCHERNUMBER"}));
            row.insert("LedgerName", ledgerName);
            row.insert("MasterID", firstNonEmptyText(entry, {"TXMLENTRYLEDGERMASTERID", "ENTRYLEDGERMASTERID", "LEDMASTERID"}).isEmpty()
                                       ? meta.value("MasterID").toString()
                                       : firstNonEmptyText(entry, {"TXMLENTRYLEDGERMASTERID", "ENTRYLEDGERMASTERID", "LEDMASTERID"}));
            row.insert("Amount", numberToString(signedAmount));
            row.insert("DrCr", signedAmount < 0 ? "Dr" : "Cr");
            row.insert("DebitAmount", numberToString(debitAmount));
            row.insert("CreditAmount", numberToString(creditAmount));
            row.insert("ParentLedger", firstNonEmptyText(entry, {"TXMLENTRYPARENTLEDGER", "ENTRYPARENTLEDGER", "PARENTLEDGER"}).isEmpty()
                                          ? meta.value("Parent").toString()
                                          : firstNonEmptyText(entry, {"TXMLENTRYPARENTLEDGER", "ENTRYPARENTLEDGER", "PARENTLEDGER"}));
            row.insert("PrimaryGroup", primaryGroup);
            row.insert("Nature", nature);
            row.insert("NatureOfGroup", natureOfGroup);
            row.insert("PAN", meta.value("PAN").toString());
            row.insert("PartyLedgerName", firstNonEmptyText(entry, {"TXMLPARTYLEDGERNAME", "PARTYLEDGERNAME"}).isEmpty() ? "N/A" : firstNonEmptyText(entry, {"TXMLPARTYLEDGERNAME", "PARTYLEDGERNAME"}));
            row.insert("PartyGSTIN", firstNonEmptyText(entry, {"TXMLPARTYGSTIN", "PARTYGSTIN"}));
            row.insert("LedgerGSTIN", firstNonEmptyText(entry, {"TXMLENTRYLEDGERGSTIN", "ENTRYLEDGERGSTIN", "LEDGERGSTIN"}).isEmpty()
                                          ? meta.value("PartyGSTIN").toString()
                                          : firstNonEmptyText(entry, {"TXMLENTRYLEDGERGSTIN", "ENTRYLEDGERGSTIN", "LEDGERGSTIN"}));
            row.insert("VoucherNarration", firstNonEmptyText(entry, {"TXMLVOUCHERNARRATION", "NARRATION", "VOUCHERNARRATION"}));
            row.insert("IsOptional", optional);
            row.insert("CompanyName", firstNonEmptyText(entry, {"TXMLCOMPANYNAME", "COMPANYNAME"}).isEmpty() ? company : firstNonEmptyText(entry, {"TXMLCOMPANYNAME", "COMPANYNAME"}));
            row.insert("FromDate", formattedFromDate);
            row.insert("ToDate", formattedToDate);
            row.insert("VoucherCategory", voucherCategory);
            rows.append(row);
        }
    }
    return rows;
}

QVector<QVariantMap> parseStockItems(const QDomDocument &doc) {
    QVector<QVariantMap> rows;
    QDomNodeList nodes = doc.elementsByTagName("STOCKITEM");
    for (int i = 0; i < nodes.size(); ++i) {
        const QDomElement elem = nodes.at(i).toElement();
        if (stripNs(elem.tagName()).toUpper() != "STOCKITEM") {
            continue;
        }

        const QString name = cleanText(elem.attribute("NAME").isEmpty() ? directChildText(elem, "NAME") : elem.attribute("NAME"));
        if (name.isEmpty()) {
            continue;
        }

        QVariantMap row;
        row.insert("Name", name);
        row.insert("Parent", directChildText(elem, "PARENT"));
        row.insert("Category", directChildText(elem, "CATEGORY"));
        row.insert("LedgerName", directChildText(elem, "LEDGERNAME"));
        row.insert("OpeningBalance", numberToString(toDoubleValue(directChildText(elem, "OPENINGBALANCE"))));
        row.insert("OpeningValue", numberToString(toDoubleValue(directChildText(elem, "OPENINGVALUE"))));
        row.insert("BasicValue", numberToString(toDoubleValue(directChildText(elem, "BASICVALUE"))));
        row.insert("BasicQty", numberToString(toDoubleValue(directChildText(elem, "BASICQTY"))));
        row.insert("OpeningRate", numberToString(toDoubleValue(directChildText(elem, "OPENINGRATE"))));
        row.insert("ClosingBalance", numberToString(toDoubleValue(directChildText(elem, "CLOSINGBALANCE"))));
        row.insert("ClosingValue", numberToString(toDoubleValue(directChildText(elem, "CLOSINGVALUE"))));
        row.insert("ClosingRate", numberToString(toDoubleValue(directChildText(elem, "CLOSINGRATE"))));
        rows.append(row);
    }
    return rows;
}

QVector<QVariantMap> parseFlatInventoryEntries(const QDomDocument &doc, const QString &company) {
    QVector<QVariantMap> rows;
    QDomNodeList nodes = doc.elementsByTagName("INVENTORYENTRY");
    for (int i = 0; i < nodes.size(); ++i) {
        const QDomElement inv = nodes.at(i).toElement();
        const QString itemName = firstNonEmptyText(inv, {"STOCKITEMNAME"});
        const QString voucherType = firstNonEmptyText(inv, {"TXMLVOUCHERTYPENAME", "VOUCHERTYPENAME"});
        if (itemName.isEmpty() || voucherType.contains("Order", Qt::CaseInsensitive)) {
            continue;
        }

        QVariantMap row;
        row.insert("Date", formatTallyDate(firstNonEmptyText(inv, {"TXMLDATE", "DATE"})));
        row.insert("VoucherTypeName", voucherType);
        row.insert("VoucherNumber", firstNonEmptyText(inv, {"TXMLVOUCHERNUMBER", "VOUCHERNUMBER"}));
        row.insert("StockItemName", itemName.trimmed());
        row.insert("BilledQty", numberToString(toDoubleValue(firstNonEmptyText(inv, {"TXMLSIGNEDQTY", "BILLEDQTY"}))));
        row.insert("Rate", numberToString(toDoubleValue(firstNonEmptyText(inv, {"RATE"}))));
        row.insert("Amount", numberToString(toDoubleValue(firstNonEmptyText(inv, {"TXMLSIGNEDAMOUNT", "AMOUNT"}))));
        row.insert("GodownName", firstNonEmptyText(inv, {"TXMLGODOWNNAME", "GODOWNNAME"}));
        row.insert("BatchName", firstNonEmptyText(inv, {"TXMLBATCHNAME", "BATCHNAME"}));
        row.insert("VoucherNarration", firstNonEmptyText(inv, {"TXMLVOUCHERNARRATION", "NARRATION", "VOUCHERNARRATION"}));
        row.insert("CompanyName", firstNonEmptyText(inv, {"TXMLCOMPANYNAME", "COMPANYNAME"}).isEmpty()
                                      ? company
                                      : firstNonEmptyText(inv, {"TXMLCOMPANYNAME", "COMPANYNAME"}));
        rows.append(row);
    }
    return rows;
}

TallyTable makeTable(const QString &id, const QString &title, const QString &fileName, const QStringList &columns,
                     const QVector<QVariantMap> &rows, const QString &companyName, const QString &fromDate, const QString &toDate) {
    TallyTable table;
    table.id = id;
    table.title = title;
    table.defaultFileName = fileName;
    table.columns = columns;

    for (QVariantMap row : rows) {
        row.insert("CompanyName", companyName);
        row.insert("FromDate", formatTallyDate(fromDate));
        row.insert("ToDate", formatTallyDate(toDate));
        for (const QString &column : columns) {
            if (!row.contains(column)) {
                row.insert(column, "");
            }
        }
        QVariantMap ordered;
        for (const QString &column : columns) {
            ordered.insert(column, row.value(column).toString());
        }
        table.rows.append(ordered);
    }
    return table;
}
}

void TallyService::resetCancellation() {
    gCancelRequested.store(false);
}

void TallyService::cancelCurrentOperation() {
    gCancelRequested.store(true);
}

void TallyService::setLogCallback(std::function<void(const QString &)> callback) {
    std::lock_guard<std::mutex> lock(gLogMutex);
    gLogCallback = std::move(callback);
}

CompanyInfo TallyService::getCompanyInfo(const QString &host, const QString &port) {
    checkCancelled();
    const QString url = QString("http://%1:%2").arg(host, port);
    const QString activeCompanyXml =
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>EXPORT</TALLYREQUEST>"
        "<TYPE>COLLECTION</TYPE><ID>MyCompanyInfo</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE>"
        "<COLLECTION NAME=\"MyCompanyInfo\"><TYPE>Company</TYPE>"
        "<FETCH>Name, StartingFrom, EndingAt, Guid</FETCH>"
        "<FILTER>IsActiveCompany</FILTER>"
        "</COLLECTION>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"IsActiveCompany\">$Name = ##SVCURRENTCOMPANY</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>";

    const auto extractCompany = [](const QDomDocument &doc) -> CompanyInfo {
        const QDomNodeList nodes = doc.elementsByTagName("COMPANY");
        for (int i = 0; i < nodes.size(); ++i) {
            const QDomElement cmp = nodes.at(i).toElement();
            const QString name = cleanText(cmp.attribute("NAME").isEmpty() ? directChildText(cmp, "NAME") : cmp.attribute("NAME"));
            if (!name.isEmpty()) {
                return {name, directChildText(cmp, "STARTINGFROM"), directChildText(cmp, "ENDINGAT")};
            }
        }

        const QDomElement root = doc.documentElement();
        const QString currentCompany = firstDescendantText(root, "SVCURRENTCOMPANY");
        if (!currentCompany.isEmpty()) {
            return {currentCompany, firstDescendantText(root, "STARTINGFROM"), firstDescendantText(root, "ENDINGAT")};
        }

        const QString companyName = firstDescendantText(root, "COMPANYNAME");
        if (!companyName.isEmpty()) {
            return {companyName, firstDescendantText(root, "STARTINGFROM"), firstDescendantText(root, "ENDINGAT")};
        }

        return {};
    };

    const QDomDocument activeDoc = parseXmlRoot(postToTally(url, activeCompanyXml, 10000));
    CompanyInfo info = extractCompany(activeDoc);
    if (!info.name.isEmpty()) {
        return info;
    }

    const QDomDocument companyListDoc = parseXmlRoot(postToTally(url, buildCompanyListRequestXml(), 10000));
    info = extractCompany(companyListDoc);
    if (!info.name.isEmpty()) {
        return info;
    }

    return {};
}

TallyDataBundle TallyService::loadAllData(const QString &host, const QString &port, const QString &company,
                                          const QString &fromDate, const QString &toDate) {
    checkCancelled();
    const QString url = QString("http://%1:%2").arg(host, port);
    QString selectedCompany = cleanText(company);
    QString selectedFrom = cleanText(fromDate);
    QString selectedTo = cleanText(toDate);
    emitLog(QString("Load started. Target: %1.").arg(url));

    if (selectedCompany.isEmpty() || selectedFrom.isEmpty() || selectedTo.isEmpty()) {
        emitLog("Company/date input incomplete; requesting active company metadata.");
        const CompanyInfo info = getCompanyInfo(host, port);
        if (selectedCompany.isEmpty()) selectedCompany = info.name;
        if (selectedFrom.isEmpty()) selectedFrom = info.startDateRaw;
        if (selectedTo.isEmpty()) selectedTo = info.endDateRaw;
        emitLog(QString("Company metadata received: company='%1', starting='%2', ending='%3'.")
                    .arg(info.name, formatTallyDate(info.startDateRaw), formatTallyDate(info.endDateRaw)));
    }
    emitLog(QString("Using company='%1', period=%2 to %3.").arg(selectedCompany, formatTallyDate(selectedFrom), formatTallyDate(selectedTo)));
    const QString sessionDir = makeSessionDir();

    emitLog("Fetching voucher type and group metadata.");
    const auto metadata = fetchTallyMetadata(url, selectedCompany);
    const QMap<QString, QString> &vtypeMap = metadata.first;
    const QMap<QString, GroupInfo> &groupMap = metadata.second;
    emitLog(QString("Metadata received. Voucher types: %1. Groups: %2.").arg(vtypeMap.size()).arg(groupMap.size()));

    emitLog("Fetching ledgers.");
    const QDomDocument ledgerDoc = parseXmlRoot(postToTally(url, buildLedgerRequestXml(selectedCompany)));
    const QVector<QVariantMap> ledgerRows = parseLedgers(ledgerDoc, groupMap);
    QMap<QString, QVariantMap> ledgerMeta;
    for (const QVariantMap &row : ledgerRows) {
        ledgerMeta.insert(row.value("Name").toString(), row);
    }
    emitLog(QString("Ledgers parsed. Rows: %1.").arg(ledgerRows.size()));

    CsvTableWriter voucherWriter(sessionDir + "/vouchers.csv", kVoucherColumns);
    CsvTableWriter allVoucherWriter(sessionDir + "/allvouchers.csv", kAllVoucherColumns);
    CsvTableWriter ledgerWriter(sessionDir + "/ledgers.csv", kLedgerColumns);
    CsvTableWriter stockItemWriter(sessionDir + "/stock_items.csv", kStockItemColumns);
    CsvTableWriter inventoryWriter(sessionDir + "/stock_vouchers.csv", kStockVoucherColumns);

    ledgerWriter.appendRows(ledgerRows, selectedCompany, selectedFrom, selectedTo);

    emitLog("Fetching accounting/all voucher rows with Tally-side flat ledger-entry TDL and probe-based chunking.");
    int voucherCount = 0;
    const int allVoucherCount = fetchChunkedRowsStream(
        url,
        selectedCompany,
        selectedFrom,
        selectedTo,
        "vouchers_flat",
        buildFlatVoucherRequestXml,
        [&](const QDomDocument &doc, const QString &chunkFrom, const QString &chunkTo) {
            return parseFlatVouchers(doc, ledgerMeta, selectedCompany, chunkFrom, chunkTo, vtypeMap);
        },
        [&](const QVector<QVariantMap> &chunkRows) {
            allVoucherWriter.appendRows(chunkRows, selectedCompany, selectedFrom, selectedTo);
            for (const QVariantMap &row : chunkRows) {
                if (row.value("VoucherCategory").toString() == "Accounting") {
                    voucherWriter.appendRow(row, selectedCompany, selectedFrom, selectedTo);
                    ++voucherCount;
                }
            }
        });

    emitLog("Fetching stock items.");
    const QDomDocument stockDoc = parseXmlRoot(postToTally(url, buildStockItemRequestXml(selectedCompany)));
    const QVector<QVariantMap> stockRows = parseStockItems(stockDoc);
    stockItemWriter.appendRows(stockRows, selectedCompany, selectedFrom, selectedTo);
    emitLog(QString("Stock items parsed. Rows: %1.").arg(stockRows.size()));

    emitLog("Fetching stock voucher rows with Tally-side flat inventory-entry TDL and probe-based chunking.");
    const int inventoryCount = fetchChunkedRowsStream(
        url,
        selectedCompany,
        selectedFrom,
        selectedTo,
        "inventory_flat",
        buildFlatInventoryEntriesRequestXml,
        [&](const QDomDocument &doc, const QString &, const QString &) {
            return parseFlatInventoryEntries(doc, selectedCompany);
        },
        [&](const QVector<QVariantMap> &chunkRows) {
            inventoryWriter.appendRows(chunkRows, selectedCompany, selectedFrom, selectedTo);
        });
    voucherWriter.flush();
    allVoucherWriter.flush();
    ledgerWriter.flush();
    stockItemWriter.flush();
    inventoryWriter.flush();

    TallyDataBundle bundle;
    bundle.companyName = selectedCompany;
    bundle.fromDateRaw = selectedFrom;
    bundle.toDateRaw = selectedTo;
    bundle.tables.insert("voucher_df", voucherWriter.table("voucher_df", "Vouchers", "vouchers.csv"));
    bundle.tables.insert("all_voucher_df", allVoucherWriter.table("all_voucher_df", "All Vouchers", "allvouchers.csv"));
    bundle.tables.insert("ledger_df", ledgerWriter.table("ledger_df", "Ledgers", "ledgers.csv"));
    bundle.tables.insert("stock_item_df", stockItemWriter.table("stock_item_df", "Stock Items", "stock_items.csv"));
    bundle.tables.insert("inventory_df", inventoryWriter.table("inventory_df", "Stock Vouchers", "stock_vouchers.csv"));
    emitLog(
        QString("Load completed. Accounting voucher rows: %1. All voucher rows: %2. Ledgers: %3. Stock items: %4. Stock vouchers: %5.")
            .arg(voucherCount)
            .arg(allVoucherCount)
            .arg(ledgerRows.size())
            .arg(stockRows.size())
            .arg(inventoryCount));
    return bundle;
}
