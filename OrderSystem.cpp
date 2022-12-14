#include <iostream>
#include "COrder.h"
#include "CCenter.h"
#include "CReply.h"
#include "CQuote.h"
#include "COSQuote.h"
#include "COOQuote.h"
#include <thread>

using namespace std;
COrder* pOrderLib;
CCenter* pCenterLib;
CReply* pReplyLib;
CQuote* pQuoteLib;
COSQuote* pOSQuoteLib;
COOQuote* pOOQuoteLib;
long g_nCode = 0;
string g_strUserId;

void Order(string g_strUserId, int nMarket,int num)
{
    long nQty = 0, nTradeType = 0, nSpecialTradeType = 0;
    int nSeqOrBook = 0,nSeqOrBookOrNum = 0;
    string strNo = "", strPrice = "", strOrderNumbertor = "", strStrikePrice = "",strExchangeNo = "",
        strStockNo = "", strYearMonth = "", strOrderDenominator = "", strOrder = "",strTriggerNumbertor = "", strTrigger = "";
    short sCallPut = 0, sSpecialTradeType = 0, sNewClose = 0,sBuySell = 0, sDayTrade = 0,
        sTradeType = 0,sPrime = 0, sPeriod = 0, sFlag = 0, sReserved = 0;
    //初始
    pOrderLib->Initialize();
    //讀取憑證
    g_nCode = pOrderLib->ReadCertByID(g_strUserId);
    pCenterLib->PrintCodeMessage("Order", "ReadCertByID", g_nCode);
    //取得帳號
    g_nCode = pOrderLib->GetUserAccount();
    pCenterLib->PrintCodeMessage("Order", "GetUserAccount", g_nCode);
    switch(num)
    {
        case 0:
            if (nMarket == 0)//證券
            {
                //輸入下單資料
                cout << "\n------------------\n";
                cout << "請輸入以下資料:\n\n";
                cout << "委託股票代號: ";
                cin >> strStockNo;

                cout << "公司類別: (0)上市上櫃 (1)興櫃 ";
                cin >> sPrime;

                cout << "委託方式: (0)整股 (1)盤後 (2)零股 ";
                cin >> sPeriod;

                cout << "當沖與否: (0)現股 (1)融資 (2)融券 (3)無券 ";
                cin >> sFlag;

                cout << "委託條件: (0)ROD (1)IOC (2)FOK ";
                cin >> nTradeType;

                cout << "委託類型: (1)市價 (2)限價 ";
                cin >> nSpecialTradeType;

                cout << "買賣類別: (0)買 (1)賣 ";
                cin >> sBuySell;

                cout << "委託量 (整股交易為張數，零股則為股數): ";
                cin >> nQty;

                cout << "委託價 (市價單為0 ; 限價單不可為0 ;昨收價為「M」): ";
                cin >> strPrice;

                g_nCode = pOrderLib->SentStockOrder(g_strUserId, false, strStockNo, sPrime, sPeriod, sFlag, nTradeType, nSpecialTradeType, sBuySell, nQty, strPrice);
                pCenterLib->PrintCodeMessage("Order", "SentStockOrder", g_nCode);
            }
            else if (nMarket == 1 || nMarket == 2) //期貨、選擇權
            {
                //輸入下單資料
                cout << "\n------------------\n";
                cout << "請輸入以下資料:\n\n";
                cout << "委託期貨/選擇權代號: ";//EX:期貨TX12，選擇權:TXO13000K2
                cin >> strStockNo;

                cout << "委託條件: (0)ROD (1)IOC (2)FOK ";
                cin >> sTradeType;

                cout << "當沖與否: (0)否 (1)是 (可當沖商品參考交易所規定) ";
                cin >> sDayTrade;

                cout << "新平倉: (0)新倉 (1)平倉 (2)自動 ";
                cin >> sNewClose;

                cout << "買賣類別: (0)買 (1)賣 ";
                cin >> sBuySell;

                cout << "委託口數: ";
                cin >> nQty;

                cout << "委託價(市價為M，範圍市價為P): ";
                cin >> strPrice;

                cout << "盤別: (0)盤中 (1)T盤預約 ";
                cin >> sReserved;

                if (nMarket == 1)//期貨
                {
                    g_nCode = pOrderLib->SentFutureOrder(g_strUserId, false, strStockNo, sTradeType, sDayTrade, sNewClose, sBuySell, nQty, strPrice, sReserved);
                    pCenterLib->PrintCodeMessage("Order", "SentFutureOrder", g_nCode);
                }
                else if (nMarket == 2)//選擇權
                {
                    g_nCode = pOrderLib->SentOptionOrder(g_strUserId, false, strStockNo, sTradeType, sDayTrade, sNewClose, sBuySell, nQty, strPrice, sReserved);
                    pCenterLib->PrintCodeMessage("Order", "SentOptionOrder", g_nCode);
                }
            }
            else if (nMarket == 3)//海外期貨
            {
                //下載海期商品檔
                cout << "下載海期商品當中...\n";
                short nCode = pOrderLib->LoadOSCommodity();
                pCenterLib->PrintCodeMessage("Order", "LoadOSCommodity", nCode);

                //輸入下單資料
                cout << "\n------------------\n";
                cout << "請輸入以下資料:\n\n";
                cout << "交易所代碼: ";
                cin >> strExchangeNo;

                cout << "海外期選代號: ";
                cin >> strStockNo;

                cout << "近月商品年月(YYYYMM):  ";
                cin >> strYearMonth;

                cout << "委託條件: (0)ROD (1)IOC (2)FOK ";
                cin >> sTradeType;

                cout << "委託價: ";
                cin >> strOrder;

                cout << "委託價分子: ";
                cin >> strOrderNumbertor;

                cout << "觸發價: ";
                cin >> strTrigger;

                cout << "觸發價分子: ";
                cin >> strTriggerNumbertor;

                cout << "買賣類別: (0)買 (1)賣 ";
                cin >> sBuySell;

                cout << "新平倉: (0)新倉 \n";//預設
                sNewClose = 0;

                cout << "當沖與否: (0)否 (1)是 (海期價差單不開放當沖) ";
                cin >> sDayTrade;

                cout << "交易口數: ";
                cin >> nQty;

                cout << "單別: (0)LMT限價單 (1)MKT市價單 (2)STL停損限價 (3)STP停損市價 ";
                cin >> sSpecialTradeType;

                g_nCode = pOrderLib->SentOverSeaFutureOrder(g_strUserId, false, strExchangeNo, strStockNo, strYearMonth, sTradeType, strOrder,
                    strOrderNumbertor, strTrigger, strTriggerNumbertor, sDayTrade, sNewClose, sBuySell, nQty, sSpecialTradeType);
                pCenterLib->PrintCodeMessage("Order", "SentOverSeaFutureOrder", g_nCode);
            }
            else if (nMarket == 4)//海外選擇權
            {
                //下載海選商品檔
                cout << "下載海選商品當中...\n";
                short nCode = pOrderLib->LoadOOCommodity();
                pCenterLib->PrintCodeMessage("Order", "LoadOOCommodity", nCode);

                //輸入下單資料
                cout << "\n------------------\n";
                cout << "請輸入以下資料:\n\n";
                cout << "交易所代碼: ";
                cin >> strExchangeNo;

                cout << "海外選擇權代號: ";
                cin >> strStockNo;

                cout << "近月商品年月(YYYYMM):  ";
                cin >> strYearMonth;

                cout << "委託條件: (0)ROD\n";//預設
                sTradeType = 0;

                cout << "委託價: ";
                cin >> strOrder;

                cout << "委託價分子: ";
                cin >> strOrderNumbertor;

                cout << "倉別: (0)新倉 (1)平倉 ";
                cin >> sNewClose;

                cout << "當沖與否: (0)否 (1)是 (海期價差單不開放當沖) ";
                cin >> sDayTrade;

                cout << "履約價: ";
                cin >> strStrikePrice;

                cout << "買賣別: (0)買 (1)賣 ";
                cin >> sBuySell;

                cout << "買賣權: (0)CALL (1)PUT ";
                cin >> sCallPut;

                cout << "交易口數: ";
                cin >> nQty;

                cout << "單別: (0)LMT限價單 ";//預設
                sSpecialTradeType=0;

                g_nCode = pOrderLib->SentOverSeaOptionOrder(g_strUserId, false, strExchangeNo,
                    strStockNo, strYearMonth, sTradeType, strOrder, strOrderNumbertor, sDayTrade, sNewClose, sCallPut, sBuySell, strStrikePrice, nQty, sSpecialTradeType);
                pCenterLib->PrintCodeMessage("Order", "SentOverSeaOptionOrder", g_nCode);
            }
            break;
        case 1:
            //輸入下單資料
            cout << "\n------------------\n";
            cout << "請輸入以下資料:\n\n";
            cout << "改量序號: ";
            cin >> strNo;
            cout << "減量數量: ";
            cin >> nQty;
            if (nMarket == 0 || nMarket == 1 || nMarket == 2)
            {
                g_nCode = pOrderLib->DecreaseOrder(g_strUserId, false, strNo, nQty, nMarket);
                pCenterLib->PrintCodeMessage("Order", "DecreaseOrder", g_nCode);
            }
            if (nMarket == 3 || nMarket == 4)
            {
                g_nCode = pOrderLib->OverSeaDecreaseOrder(g_strUserId, false, strNo, nQty, nMarket);
                pCenterLib->PrintCodeMessage("Order", "OverSeaDecreaseOrder", g_nCode);
            }
            break;
        case 2:
            //輸入下單資料
            cout << "\n------------------\n";
            cout << "請輸入以下資料:\n\n";
            if (nMarket == 0 || nMarket == 1 || nMarket == 2)
            {
                cout << "改價方式: (0)序號改價 (1)書號改價 ";
                cin >> nSeqOrBook;
                cout << "序號/書號: ";
                cin >> strNo;
                cout << "改價價格: ";
                cin >> strPrice;
                cout << "委託條件: (0)ROD (1)IOC (2)FOK\n";
                cin >> nTradeType;
                g_nCode = pOrderLib->CorrectPrice(g_strUserId, false, strNo, strPrice, nTradeType, nSeqOrBook, nMarket);
            }
            else if (nMarket == 3 || nMarket == 4)
            {
                cout << "改價方式: (0)書號改價\n";//預設
                nSeqOrBook = 0;

                cout << "書號: ";
                cin >> strNo;

                cout << "交易所代碼: ";
                cin >> strExchangeNo;

                cout << "海期代號: ";
                cin >> strStockNo;

                cout << "商品年月(YYYYMM):  ";
                cin >> strYearMonth;

                cout << "委託價: ";
                cin >> strPrice;

                cout << "委託價分子: ";
                cin >> strOrderNumbertor;

                cout << "委託價分母: ";
                cin >> strOrderDenominator;

                cout << "委託條件: (0)ROD\n";//預設
                sTradeType = 0;

                cout << "委託類型: (0)LMT限價單\n";//預設
                sSpecialTradeType = 0;

                if (nMarket == 3)
                {
                    cout << "倉別: (0)新倉\n";//預設
                    sNewClose = 0;
                }
                else if (nMarket == 4)
                {
                    cout << "倉別: (0)新倉 (1)平倉 ";
                    cin >> sNewClose;

                    cout << "買賣權: (0)CALL (1)PUT ";
                    cin >> sCallPut;

                    cout << "履約價: ";
                    cin >> strStrikePrice;
                }
                g_nCode = pOrderLib->OverSeaCorrectPrice(g_strUserId, false, strNo, strExchangeNo, strStockNo, strYearMonth, strPrice, strOrderNumbertor, strOrderDenominator, sTradeType, sSpecialTradeType, sNewClose, sCallPut, strStrikePrice, nSeqOrBook, nMarket);
            }
            pCenterLib->PrintCodeMessage("Order", "CorrectPrice", g_nCode);
            break;
        case 3:
            nTradeType = 0;//委託條件ROD
            //輸入下單資料
            cout << "\n------------------\n";
            cout << "請輸入以下資料:\n\n";
            if (nMarket == 0 || nMarket == 1 || nMarket == 2)
            {
                cout << "刪單類別： (0)序號刪單 (1)書號刪單 (2)商品代號刪單 ";
                cin >> nSeqOrBookOrNum;
                cout << "序號/書號/商品代號： ";
                cin >> strNo;
                g_nCode = pOrderLib->CancelOrder(g_strUserId, false, strNo, nTradeType, nSeqOrBookOrNum, nMarket);
            }
            else if (nMarket == 3 || nMarket == 4)//海期、海選
            {
                cout << "刪單類別： (0)序號刪單 (1)書號刪單 ";
                cin >> nSeqOrBookOrNum;
                cout << "序號/書號： ";
                cin >> strNo;
                g_nCode = pOrderLib->OverSeaCancelOrder(g_strUserId, false, strNo, nTradeType, nSeqOrBookOrNum, nMarket);
            }
            pCenterLib->PrintCodeMessage("Order", "CancelOrder", g_nCode);
            break;
        default:
            break;
    }
}
void Quote(int num,int nMarket)
{
    int nYesno = 0, nChoice = 0;
    short sMarketNo = 0,sPageNo = 0,sType=0,nCode = 0;
    string strStockNos = "";
    switch (nMarket)
    {
        case 0://證券/期貨/選擇權
            switch (num)
            {
                case 0://報價連線
                    g_nCode = pQuoteLib->EnterMonitorLong();
                    pCenterLib->PrintCodeMessage("Quote", "EnterMonitor", g_nCode);
                    break;
                case 1://連線狀態
                    g_nCode = pQuoteLib->IsConnected();
                    cout<<"Quote-Isconnected (0:未連線,1:連線,2:下載中): "<<g_nCode;
                    break;
                case 2://報價斷線
                    g_nCode = pQuoteLib->LeaveMonitor();
                    pCenterLib->PrintCodeMessage("Quote", "LeaveMonitor", g_nCode);
                    break;
                case 3://訂閱指定商品即時報價
                    if (pQuoteLib->IsConnected()!= 1)
                    {
                        cout << "尚未連線" << endl;
                        break;
                    }
                    else
                    {
                        sPageNo = 1;
                        cout << "請輸入商品代號(一筆以上請用", "做區隔): ";
                        cin >> strStockNos;
                        g_nCode = pQuoteLib->RequestStocks(&sPageNo, strStockNos);
                        pCenterLib->PrintCodeMessage("Quote", "RequestStocks", g_nCode);
                        break;
                    }
                case 4://訂閱要求傳送成交明細以及五檔
                    cout << "零股: (0)是 (1)否 ";
                    cin >> nYesno;
                    if (nYesno==0)
                    {
                        cout << "盤中零股盤別: (5)上市 (6)上櫃 ";
                        cin >> sMarketNo;
                        sPageNo = -1;
                        cout << "請輸入商品代號(一筆以上請用", "做區隔): ";
                        cin >> strStockNos;
                        g_nCode = pQuoteLib->RequestTicksWithMarketNo(&sPageNo,sMarketNo,strStockNos);
                        pCenterLib->PrintCodeMessage("Quote", "RequestTicksWithMarketNo", g_nCode);
                    }
                    else if(nYesno==1)
                    {
                        sPageNo = -1;
                        cout << "請輸入商品代號 (一筆以上請用,做區隔): ";
                        cin >> strStockNos;
                        g_nCode = pQuoteLib->RequestTicks(&sPageNo, strStockNos);
                        pCenterLib->PrintCodeMessage("Quote", "RequestTicks", g_nCode);
                    }
                    break;
                case 5://stocklist 查詢國內商品清單
                    cout << "請輸入市場代碼 (0)上市 (1)上櫃 (2)期貨 (3)選擇權 (4)興櫃 (5)上市盤中零股 (6)上櫃盤中零股： ";
                    cin >> sMarketNo;
                    g_nCode = pQuoteLib->RequestStockList(sMarketNo);
                    pCenterLib->PrintCodeMessage("Quote", "RequestStockList", g_nCode);
                    break;
                case 6:
                    cout << "\n離開回報功能\n";
                    break;
                default:
                    break;
            }
            break;
        case 1://海外期貨
            switch (num)
            {
                case 0://報價連線
                    g_nCode = pOSQuoteLib->EnterMonitorLong();
                    pCenterLib->PrintCodeMessage("OSQoute", "EnterMonitor", g_nCode);
                    break;
                case 1://連線狀態
                    g_nCode = pOSQuoteLib->IsConnected(); 
                    cout << "OSQuote-Isconnected(0:未連線,1:連線,2:下載中):" << g_nCode;
                    break;
                case 2://報價斷線
                    g_nCode = pOSQuoteLib->LeaveMonitor();
                    pCenterLib->PrintCodeMessage("OSQuote", "LeaveMonitor", g_nCode);
                    break;
                case 3://訂閱指定商品即時報價，由OnNotifyQuote取得通知
                    sPageNo = -1;
                    cout << "請輸入商品代號 (一筆以上請用,做區隔): ";
                    cin >> strStockNos;
                    g_nCode = pOSQuoteLib->RequestStocks(&sPageNo, strStockNos);
                    pCenterLib->PrintCodeMessage("OSQuote", "RequestStocks", g_nCode);
                    break;
                case 4://訂閱傳送成交明細及五檔
                    //先確認有無連線
                    sPageNo = -1;
                    cout << "請輸入商品代號 (一筆以上請用,做區隔): ";
                    cin >> strStockNos;
                    g_nCode = pOSQuoteLib->RequestTicks(&sPageNo, strStockNos);
                    pCenterLib->PrintCodeMessage("OSQuote", "RequestTicks", g_nCode);
                    break;
                case 5://查詢商品清單
                    cout << "(0)商品資訊 (1)詳細商品資訊(含下單代碼): ";
                    cin >> nChoice;
                    if (nChoice == 0)
                    {
                        g_nCode = pOSQuoteLib->RequestOverseaProducts();
                        pCenterLib->PrintCodeMessage("OSQuote", "RequestOverseaProducts", g_nCode);
                    }
                    else if (nChoice == 1)
                    {
                        short sType = 1;
                        g_nCode = pOSQuoteLib->GetOverseaProductDetail(sType);
                        pCenterLib->PrintCodeMessage("OSQuote", "GetOverseaProductDetail", g_nCode);
                    }
                    break;
                case 6:
                    cout << "\n離開回報功能\n";
                    break;
                default:
                    break;
            }
            break;
        case 2://海外選擇權
            switch (num)
            {
                case 0://報價連線
                    g_nCode = pOOQuoteLib->EnterMonitorLong();
                    pCenterLib->PrintCodeMessage("OOQoute", "EnterMonitor", g_nCode);
                    break;
                case 1://連線狀態
                    g_nCode = pOOQuoteLib->IsConnected();
                    cout << "OOQuote-Isconnected(0:未連線,1:連線,2:下載中):" << g_nCode;
                    break;
                case 2://報價斷線
                    g_nCode = pOOQuoteLib->LeaveMonitor();
                    pCenterLib->PrintCodeMessage("OOQuote", "LeaveMonitor", g_nCode);
                    break;
                case 3://訂閱報價
                    sPageNo = -1;
                    cout << "請輸入商品代號(一筆以上請用","做區隔): ";
                    cin >> strStockNos;
                    g_nCode = pOOQuoteLib->RequestStocks(&sPageNo, strStockNos);
                    pCenterLib->PrintCodeMessage("OOQuote", "RequestStocks", g_nCode);
                    break;
                case 4://訂閱ticks及五檔tickets&5-訂閱要求傳送成交明細以及五檔
                    //確認連線
                    sPageNo = -1;
                    cout << "請輸入商品代號(一筆以上請用","做區隔): ";
                    cin >> strStockNos;
                    g_nCode = pOOQuoteLib->RequestTicks(&sPageNo, strStockNos);
                    pCenterLib->PrintCodeMessage("OOQuote", "RequestTicks", g_nCode);
                    break;
                case 5://要求商品資訊
                    g_nCode = pOOQuoteLib->RequestProducts();
                    pCenterLib->PrintCodeMessage("OOQuote", "RequestProducts", g_nCode);
                    break;
                case 6:
                    cout << "\n離開回報功能\n";
                    break;
                default:
                    break;
            }
            break;
        default:
            break;
    }
}
void Reply(int num, string g_strUserId)
{
    switch (num)
    {
        case 0://連線
            g_nCode = pReplyLib->ConnectById(g_strUserId);//0表示成功
            pCenterLib->PrintCodeMessage("Reply", "ConnectByID", g_nCode);
            break;
        case 1://連線狀態
            g_nCode = pReplyLib->IsConnecttedByID(g_strUserId);//0斷線；1連線中；2下載中。
            pCenterLib->PrintCodeMessage("Reply", "IsConnecttedByID", g_nCode);
            break;
        case 2://斷線
            g_nCode = pReplyLib->SolaceCloseByID(g_strUserId);
            pCenterLib->PrintCodeMessage("Reply", "SolaceCloseByID", g_nCode);
            break;
        default:
            break;
    }
}
void thread_Main()
{
    bool boolwhile = true, boolcontinue = true;
    int nfunction = 0, nMarket = 0;
    while (boolwhile)
    {
        boolcontinue = true;
        cout << "\n------------------\n";
        cout << "請選取以下功能\n\n";
        cout << "功能列 (0)下單 (1)回報 (2)報價 (3)退出: ";
        cin >> nfunction;
        if (nfunction == 0)
        {
            while (boolcontinue)
            {
                int num = 0;
                cout << "\n------------------\n";
                cout << "請選取以下功能: \n";
                cout << "(下單前請先執行報價連線)\n\n";
                cout << "(0)委託下單 (1)改量 (2)改價 (3)刪單 (4)離開: ";
                cin >> num;
                if (num < 4)
                {
                    cout << "\n市場: (0)證券 (1)期貨 (2)選擇權 (3)海外期貨 (4)海外選擇權 ";
                    cin >> nMarket;
                    Order(g_strUserId, nMarket, num);
                }
                else if (num == 4)
                {
                    cout << "\n離開下單功能\n";
                    boolcontinue = false;
                }
                else
                {
                    cout << "\n輸入錯誤，請重新輸入\n";
                }
            }
        }
        else if (nfunction == 1)
        {
            while (boolcontinue)
            {
                int num = 0;
                cout << "\n------------------\n";
                cout << "請選取以下功能\n\n";
                cout << "(0)回報連線 (1)連線狀態 (2)回報斷線 (3)離開: ";
                cin >> num;
                if (num == 0 || num == 1 || num == 2)
                {
                    Reply(num, g_strUserId);
                }
                else if (num==3)
                {
                    cout << "\n離開回報功能\n";
                    boolcontinue = false;
                }
                else
                {
                    cout << "\n輸入錯誤，請重新輸入\n";
                }
            }
        }
        else if (nfunction == 2 )
        {
            while (boolcontinue)
            {
                int num = 0;
                cout << "\n------------------\n";
                cout << "請選取以下功能\n\n";
                cout << "(0)報價連線 (1)連線狀態 (2)報價斷線 (3)報價 (4)訂閱Ticks及5檔 (5)StockList (6)離開: ";
                cin >> num;
                if (num < 6)
                {
                    cout << "\n類別: (0)證券/期貨/選擇權 (1)海外期貨 (2)海外選擇權 ";
                    cin >> nMarket;
                }
                if(num >= 0 && num < 6)
                {
                    Quote(num, nMarket);
                }
                else if (num == 6)
                {
                    cout << "\n離開報價功能\n";
                    boolcontinue = false;
                }
                else
                {
                    cout << "\n輸入錯誤，請重新輸入\n";
                }
            }
        }
        else if (nfunction == 3)
        {
            boolwhile = false;
            cout << "\n離開程式\n"; 
        }
        else
        {
            cout << "\n輸入錯誤，請重新輸入\n";
        }
    }
    delete pCenterLib;
    delete pReplyLib;
    delete pOrderLib;
    delete pQuoteLib;
    delete pOOQuoteLib;
    delete pOSQuoteLib;
    CoUninitialize();
    exit(0);
}
int main()
{
    CoInitialize(NULL);
    pCenterLib = new CCenter;
    pOrderLib = new COrder;
    pReplyLib = new CReply;
    pQuoteLib = new CQuote;
    pOSQuoteLib = new COSQuote;
    pOOQuoteLib = new COOQuote;

    string pwd;
    cout << "\n--------登入證券下單系統--------";
    cout << "\n請輸入身分證: ";
    cin>>g_strUserId;
    cout << "\n請輸入密碼： ";
    
    //連線設置主控台，改變寫入模式
    HANDLE hStdin = GetStdHandle(STD_INPUT_HANDLE);
    DWORD mode = 0;
    GetConsoleMode(hStdin, &mode);
    SetConsoleMode(hStdin, mode & (~ENABLE_ECHO_INPUT));
    //輸入密碼
    cin >> pwd;
    cout << endl;
    g_nCode = pCenterLib->Login(g_strUserId.c_str(), pwd.c_str());//login
    pCenterLib->PrintCodeMessage("Center", "Login", g_nCode);
    if (g_nCode != 0)
    {
        return 0;
    }
    SetConsoleMode(hStdin, mode);//寫入模式回歸預設模式
    thread tMain(thread_Main);
    if (tMain.joinable())
    {
        tMain.detach();
    }
    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) // Get SendMessage loop 
    {
        DispatchMessageW(&msg);
    }
    return 0;
}
