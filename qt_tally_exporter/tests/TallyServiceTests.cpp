// Include the implementation to exercise private XML builders and parsers without a live Tally server.
#include "../src/TallyService.cpp"
#include <QCoreApplication>
#include <iostream>

void check(bool condition, const char *message) {
    if (!condition) throw std::runtime_error(message);
}

int main(int argc, char **argv) {
    QCoreApplication app(argc, argv);
    try {
        for (const auto range : {qMakePair(-1, -1), qMakePair(10, 20)}) {
            QDomDocument request;
            check(bool(request.setContent(buildFlatVoucherRequestXml("A & B", "20260401", "20260430", range.first, range.second))),
                  "Voucher request must be well-formed XML");
            check(request.toString().contains("&lt; 0"), "Formula comparisons must be XML escaped");
        }
        for (const QString tag : {QString("LEDGERENTRY"), QString("ALLLEDGERENTRIES.LIST"), QString("LEDGERENTRIES.LIST"), QString("CUSTOMROW")}) {
            const QString fields = "<LEDGERNAME>Cash</LEDGERNAME><AMOUNT>-125.50</AMOUNT>"
                "<TXMLVOUCHERTYPENAME>Payment</TXMLVOUCHERTYPENAME><TXMLDATE>20260401</TXMLDATE>";
            const auto doc = parseXmlRoot("<ENVELOPE><COLLECTION><" + tag + ">" + fields + "</" + tag + "></COLLECTION></ENVELOPE>");
            const auto rows = parseFlatVouchers(doc, {}, "Test", "20260401", "20260430", {});
            check(rows.size() == 1, "Ledger row must be exported exactly once regardless of wrapper tag");
            check(rows[0].value("VoucherCategory") == "Accounting", "Payment must enter vouchers.csv");
            check(rows[0].value("DebitAmount").toDouble() == 125.50, "Debit amount must be preserved");
        }
        for (const QString xml : {QString("<RESPONSE>Unknown Request, cannot be processed</RESPONSE>"),
                                 QString("<ENVELOPE><LINEERROR>Invalid company</LINEERROR></ENVELOPE>"),
                                 QString("<ENVELOPE><HEADER><STATUS>0</STATUS></HEADER></ENVELOPE>")}) {
            bool rejected = false;
            try { parseXmlRoot(xml); } catch (const std::runtime_error &) { rejected = true; }
            check(rejected, "Tally error response must not become an empty successful export");
        }
        check(parseFlatVouchers(parseXmlRoot("<ENVELOPE><COLLECTION/></ENVELOPE>"), {}, "", "", "", {}).isEmpty(),
              "A valid empty collection must remain supported");
        std::cout << "All voucher regression checks passed\n";
        if (app.arguments().contains("--live")) {
            const auto info = TallyService::getCompanyInfo("localhost", "9000");
            check(!info.name.isEmpty(), "No active Tally company");
            const auto probe = parseXmlRoot(postToTally("http://localhost:9000",
                buildVoucherProbeRequestXml(info.name, info.startDateRaw, info.endDateRaw)));
            const auto vouchers = probe.elementsByTagName("VOUCHER");
            check(!vouchers.isEmpty(), "No voucher headers available for live verification");
            QString date;
            for (int i = 0; i < vouchers.size() && date.isEmpty(); ++i) {
                if (isRealVoucher(vouchers.at(i).toElement())) {
                    date = directChildText(vouchers.at(i).toElement(), "DATE");
                }
            }
            check(!date.isEmpty(), "Probe voucher has no date");
            const auto response = parseXmlRoot(postToTally("http://localhost:9000",
                buildFlatVoucherRequestXml(info.name, date, date, -1, -1)));
            const auto rows = parseFlatVouchers(response, {}, info.name, date, date, {});
            check(!rows.isEmpty(), "Live voucher response contains no ledger rows");
            int accountingRows = 0;
            for (const auto &row : rows) {
                if (row.value("VoucherCategory") == "Accounting") ++accountingRows;
            }
            std::cout << "Live single-day check: " << rows.size() << " all-voucher rows, "
                      << accountingRows << " accounting rows\n";
        }
    } catch (const std::exception &error) {
        std::cerr << error.what() << '\n';
        return 1;
    }
    return 0;
}
