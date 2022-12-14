#include "COrder.h"
COrder::COrder()
{
	m_pSKOrderLib.CreateInstance(__uuidof(SKCOMLib::SKOrderLib));
	m_pSKOrderLibEventHandler = new ISKOrderLibEventHandler(*this, m_pSKOrderLib, &COrder::OnEventFiringObjectInvoke);
	vec_strFullAccount_TS = {};
	vec_strFullAccount_TF = {};
	vec_strFullAccount_OF = {};
}
COrder::~COrder()
{
	if (m_pSKOrderLibEventHandler != NULL)
	{
		m_pSKOrderLibEventHandler->ShutdownConnectionPoint();
		m_pSKOrderLibEventHandler->Release();
		m_pSKOrderLibEventHandler = NULL;
	}
	if (m_pSKOrderLib)
	{
		m_pSKOrderLib->Release();
	}
}
long COrder::Initialize()
{
	return  m_pSKOrderLib->SKOrderLib_Initialize();
}
HRESULT COrder::OnEventFiringObjectInvoke
(
	ISKOrderLibEventHandler* pEventHandler,
	DISPID dispidMember,
	REFIID riid,
	LCID lcid,
	WORD wFlags,
	DISPPARAMS* pdispparams,
	VARIANT* pvarResult,
	EXCEPINFO* pexcepinfo,
	UINT* puArgErr
)
{
	VARIANT varlValue;
	VariantInit(&varlValue);
	VariantClear(&varlValue);

	switch (dispidMember)
	{
		case 1:
		{
			varlValue = (pdispparams->rgvarg)[1];
			_bstr_t bstrLoginID = V_BSTR(&varlValue);
			varlValue = (pdispparams->rgvarg)[0];
			_bstr_t bstrData = V_BSTR(&varlValue);
			OnAccount(string(bstrLoginID), string(bstrData));
			break;
		}
		case 2:
		{
			varlValue = (pdispparams->rgvarg)[2];
			LONG nThreadID = V_I4(&varlValue);
			varlValue = (pdispparams->rgvarg)[1];
			LONG nCode = V_I4(&varlValue);
			varlValue = (pdispparams->rgvarg)[0];
			_bstr_t bstrMessage = V_BSTR(&varlValue);
			OnAsyncOrder(nThreadID, nCode, string(bstrMessage));
			break;
		}
	}
	return S_OK;
}
long COrder::SentStockOrder(string strLogInID,bool bAsyncOrder, string strStockNo, short sPrime, short sPeriod, short sFlag, long nTradeType, long nSpecialTradeType, short sBuySell, long nQty, string strPrice)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	BSTR bstrMessage;//�|�^�ǥ��ѩΦ��\�T��

	//�T�{�b��
	if (vec_strFullAccount_TS.size() != 0)
	{
		strFullAccount = vec_strFullAccount_TS[0];
		bstrFullAccount = strFullAccount.c_str();
	}
	else
	{
		cout << "SenrStockOrder Error: No Stock Account";
		return -1;
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrStockNo= strStockNo.c_str();
	bstr_t bstrPrice = strPrice.c_str();
	
	SKCOMLib::STOCKORDER pOrder;
	pOrder.bstrFullAccount = bstrFullAccount;
	pOrder.bstrStockNo = bstrStockNo;
	pOrder.sPrime = sPrime;
	pOrder.sPeriod = sPeriod;
	pOrder.sFlag = sFlag;
	pOrder.nTradeType = nTradeType;
	pOrder.nSpecialTradeType = nSpecialTradeType;
	pOrder.sBuySell = sBuySell;
	pOrder.nQty = nQty;
	pOrder.bstrPrice = bstrPrice;
	
	long m_nCode=m_pSKOrderLib->SendStockOrder(bstrLogInID, VARIANT_BOOL(bAsyncOrder), &pOrder, &bstrMessage);
	cout << "SentStockOrder : " << string(_bstr_t(bstrMessage)) << endl;
	::SysFreeString(bstrMessage);
	return m_nCode;
}
long COrder::SentFutureOrder(string strLogInID, bool bAsyncOrder, string strStockNo, short sTradeType, short sDayTrade, short sNewClose, short sBuySell, long nQty, string strPrice,short sReserved)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	BSTR bstrMessage;

	//�T�{�b��
	if (vec_strFullAccount_TF.size() != 0)
	{
		strFullAccount = vec_strFullAccount_TF[0];
		bstrFullAccount = strFullAccount.c_str();
	}
	else
	{
		cout << "SentFutureOrder Error: No Future Account";
		return -1;
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrStockNo = strStockNo.c_str();
	bstr_t bstrPrice = strPrice.c_str();

	SKCOMLib::FUTUREORDER pOrder;
	pOrder.bstrFullAccount = bstrFullAccount;
	pOrder.bstrStockNo = bstrStockNo;
	pOrder.sNewClose = sNewClose;
	pOrder.sDayTrade = sDayTrade;
	pOrder.sTradeType = sTradeType;
	pOrder.sBuySell = sBuySell;
	pOrder.nQty = nQty;
	pOrder.bstrPrice = bstrPrice;
	pOrder.sReserved = sReserved;
	long m_nCode = m_pSKOrderLib->SendFutureOrder(bstrLogInID, VARIANT_BOOL(bAsyncOrder), &pOrder, &bstrMessage);
	cout << "SentFutureOrder : " << string(_bstr_t(bstrMessage)) << endl;
	::SysFreeString(bstrMessage);
	return m_nCode;
}
long COrder::SentOptionOrder(string strLogInID, bool bAsyncOrder, string strStockNo, short sTradeType, short sDayTrade, short sNewClose, short sBuySell, long nQty, string strPrice, short sReserved)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	BSTR bstrMessage;

	//�T�{�b��
	if (vec_strFullAccount_TF.size() != 0)
	{
		strFullAccount = vec_strFullAccount_TF[0];
		bstrFullAccount = strFullAccount.c_str();
	}
	else
	{
		cout << "SentOptionOrder Error: No Option Account";
		return -1;
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrStockNo = strStockNo.c_str();
	bstr_t bstrPrice = strPrice.c_str();

	SKCOMLib::FUTUREORDER pOrder;
	pOrder.bstrFullAccount = bstrFullAccount;
	pOrder.bstrStockNo = bstrStockNo;
	pOrder.sNewClose = sNewClose;
	pOrder.sDayTrade = sDayTrade;
	pOrder.sTradeType = sTradeType;
	pOrder.sBuySell = sBuySell;
	pOrder.nQty = nQty;
	pOrder.bstrPrice = bstrPrice;
	pOrder.sReserved = sReserved;
	long m_nCode = m_pSKOrderLib->SendOptionOrder(bstrLogInID, VARIANT_BOOL(bAsyncOrder), &pOrder, &bstrMessage);
	cout << "SentOptionOrder : " << string(_bstr_t(bstrMessage)) << endl;
	::SysFreeString(bstrMessage);
	return m_nCode;
}
long COrder::SentOverSeaFutureOrder(string strLogInID, bool bAsyncOrder, string strExchangeNo, string strStockNo, string  strYearMonth, short sTradeType, string  strOrder, string strOrderNumbertor, string strTrigger, string strTriggerNumbertor, short sDayTrade, short sNewClose, short sBuySell, int nQty, short sSpecialTradeType)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	BSTR bstrMessage;

	//�T�{�b��
	if (vec_strFullAccount_OF.size() != 0)
	{
		strFullAccount = vec_strFullAccount_OF[0];
		bstrFullAccount = strFullAccount.c_str();
	}
	else
	{
		cout << "SentOverSeaFutureOrder Error: No OverSeaFuture Account";
		return -1;
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrExchangeNo = strExchangeNo.c_str();
	bstr_t bstrStockNo = strStockNo.c_str();
	bstr_t bstrYearMonth = strYearMonth.c_str();
	bstr_t bstrOrder = strOrder.c_str();
	bstr_t bstrOrderNumbertor = strOrderNumbertor.c_str();
	bstr_t bstrTrigger = strTrigger.c_str();
	bstr_t bstrTriggerNumbertor = strTriggerNumbertor.c_str();

	SKCOMLib::OVERSEAFUTUREORDER pOrder;
	pOrder.bstrFullAccount = bstrFullAccount;
	pOrder.bstrExchangeNo = bstrExchangeNo;
	pOrder.bstrStockNo = bstrStockNo;
	pOrder.sNewClose = sNewClose;
	pOrder.sDayTrade = sDayTrade;
	pOrder.sTradeType = sTradeType;
	pOrder.sBuySell = sBuySell;
	pOrder.nQty = nQty;
	pOrder.bstrOrder = bstrOrder;
	pOrder.bstrYearMonth = bstrYearMonth;
	pOrder.sSpecialTradeType = sSpecialTradeType;
	pOrder.bstrOrderNumerator = bstrOrderNumbertor;
	pOrder.bstrTrigger = bstrTrigger;
	pOrder.bstrTriggerNumerator = bstrTriggerNumbertor;

	long m_nCode = m_pSKOrderLib->SendOverseaFutureOrder(bstrLogInID, VARIANT_BOOL(bAsyncOrder), &pOrder, &bstrMessage);
	cout << "SentOverSeaFutureOrder : " << string(_bstr_t(bstrMessage)) << endl;
	::SysFreeString(bstrMessage);
	return m_nCode;
}
long COrder::SentOverSeaOptionOrder(string strLogInID, bool bAsyncOrder, string strExchangeNo, string strStockNo, string  strYearMonth, short sTradeType, string  strOrder, string strOrderNumbertor, short sDayTrade, short sNewClose, short sCallPut, short sBuySell, string strStrikePrice, int nQty, short sSpecialTradeType)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	BSTR bstrMessage;

	//�T�{�b��
	if (vec_strFullAccount_OF.size() != 0)
	{
		strFullAccount = vec_strFullAccount_OF[0];
		bstrFullAccount = strFullAccount.c_str();
	}
	else
	{
		cout << "SentOverSeaOptionOrder Error: No OverSeaOption Account";
		return -1;
	}
	//string�ഫ��bstr
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrExchangeNo = strExchangeNo.c_str();
	bstr_t bstrStockNo = strStockNo.c_str();
	bstr_t bstrYearMonth = strYearMonth.c_str();
	bstr_t bstrOrder = strOrder.c_str();
	bstr_t bstrOrderNumbertor = strOrderNumbertor.c_str();
	bstr_t bstrStrikePrice = strStrikePrice.c_str();

	SKCOMLib::OVERSEAFUTUREORDER pOrder;
	pOrder.bstrFullAccount = bstrFullAccount;
	pOrder.bstrExchangeNo = bstrExchangeNo;
	pOrder.bstrStockNo = bstrStockNo;
	pOrder.sNewClose = sNewClose;
	pOrder.sDayTrade = sDayTrade;
	pOrder.sTradeType = sTradeType;
	pOrder.nQty = nQty;
	pOrder.bstrOrder = bstrOrder;
	pOrder.bstrYearMonth = bstrYearMonth;
	pOrder.sSpecialTradeType = sSpecialTradeType;
	pOrder.bstrOrderNumerator = bstrOrderNumbertor;
	pOrder.bstrStrikePrice = bstrStrikePrice;
	pOrder.sCallPut = sCallPut;
	pOrder.sBuySell = sBuySell;

	long m_nCode = m_pSKOrderLib->SendOverseaOptionOrder(bstrLogInID, VARIANT_BOOL(bAsyncOrder), &pOrder, &bstrMessage);
	cout << "SentOverSeaOptionOrder : " << string(_bstr_t(bstrMessage)) << endl;
	::SysFreeString(bstrMessage);
	return m_nCode;
}
long COrder::DecreaseOrder(string strLogInID, bool bAsyncOrder,string strNo,long nQty,int nMarket)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	BSTR bstrMessage;
	//�T�{�b��
	if(nMarket==0)
	{
		if (vec_strFullAccount_TS.size() != 0)
		{
			strFullAccount = vec_strFullAccount_TS[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "DecreaseOrder Error: No Stock Account";
			return -1;
		}
	}
	else if (nMarket == 1|| nMarket == 2)
	{
		if (vec_strFullAccount_TF.size() != 0)
		{
			strFullAccount = vec_strFullAccount_TF[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "DecreaseOrder Error: No Future/Option Account";
			return -1;
		}
	}
	
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrNo = strNo.c_str();

	long m_nCode = m_pSKOrderLib->DecreaseOrderBySeqNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount, bstrNo, nQty,&bstrMessage);
	cout << "DecreaseOrderBySeqNo : " << string(_bstr_t(bstrMessage)) << endl;
	::SysFreeString(bstrMessage);
	return m_nCode;
}
long COrder::CorrectPrice(string strLogInID, bool bAsyncOrder, string strNo, string strPrice, long nTradeType, int nSeqOrBook, int nMarket)
{
	string strFullAccount = "",strMarketSymbol="";
	bstr_t bstrFullAccount;
	long m_nCode = 0;
	//�T�{�b��
	if (nMarket == 0)
	{
		if (vec_strFullAccount_TS.size() != 0)
		{
			strFullAccount = vec_strFullAccount_TS[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "CorrectPrice Error: No Stock Account";
			return -1;
		}
		strMarketSymbol = "TS";
	}
	else if (nMarket == 1 || nMarket == 2)
	{
		if (vec_strFullAccount_TF.size() != 0)
		{
			strFullAccount = vec_strFullAccount_TF[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "CorrectPrice Error: No Future/Option Account";
			return -1;
		}
		if (nMarket == 1)
		{
			strMarketSymbol = "TF";
		}
		else
		{
			strMarketSymbol = "TO";
		}
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrNo = strNo.c_str();
	bstr_t bstrPrice = strPrice.c_str();
	bstr_t bstrMarketSymbol = strMarketSymbol.c_str();
	BSTR bstrMessage;
	if (nSeqOrBook == 0)
	{
		m_nCode = m_pSKOrderLib->CorrectPriceBySeqNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount,bstrNo, bstrPrice, nTradeType,&bstrMessage);//�Ǹ�
		cout << "CorrectPriceBySeqNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
	else if(nSeqOrBook == 1)
	{
		m_nCode = m_pSKOrderLib->CorrectPriceByBookNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount,bstrMarketSymbol, bstrNo, bstrPrice, nTradeType, &bstrMessage);//�Ѹ�
		cout << "CorrectPriceByBookNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
}
long COrder::CancelOrder(string strLogInID, bool bAsyncOrder, string strNo, long nTradeType, int nSeqOrBookOrNum,int nMarket)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	long m_nCode = 0;
	//�T�{�b��
	if (nMarket == 0)
	{
		if (vec_strFullAccount_TS.size() != 0)
		{
			strFullAccount = vec_strFullAccount_TS[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "CancelOrder Error: No Stock Account";
			return -1;
		}
	}
	else if (nMarket == 1|| nMarket == 2)
	{
		if (vec_strFullAccount_TF.size() != 0)
		{
			strFullAccount = vec_strFullAccount_TF[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "CancelOrder Error: No Future/Option Account";
			return -1;
		}
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrNo = strNo.c_str();
	BSTR bstrMessage;
	if (nSeqOrBookOrNum == 0)
	{
		m_nCode = m_pSKOrderLib->CancelOrderBySeqNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount, bstrNo, &bstrMessage);//�Ǹ�
		cout << "CancelOrderBySeqNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
	else if (nSeqOrBookOrNum == 1)
	{
		m_nCode = m_pSKOrderLib->CancelOrderByBookNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount, bstrNo, &bstrMessage);//�Ѹ�
		cout << "CancelOrderByBookNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
	else if (nSeqOrBookOrNum == 2)
	{
		m_nCode = m_pSKOrderLib->CancelOrderByStockNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount, bstrNo, &bstrMessage);//�ӫ~�N��
		cout << "CancelOrderByStockNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
}
long COrder::OverSeaDecreaseOrder(string strLogInID, bool bAsyncOrder, string strNo, long nQty, int nMarket)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	BSTR bstrMessage;
	//�T�{�b��
	if (nMarket == 3 || nMarket == 4)
	{
		if (vec_strFullAccount_OF.size() != 0)
		{
			strFullAccount = vec_strFullAccount_OF[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "OverSeaDecreaseOrder Error: No OverSea Future/Option Account";
			return -1;
		}
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrNo = strNo.c_str();

	long m_nCode = m_pSKOrderLib->OverSeaDecreaseOrderBySeqNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount, bstrNo, nQty, &bstrMessage);
	cout << "OverSeaDecreaseOrderBySeqNo : " << string(_bstr_t(bstrMessage)) << endl;
	::SysFreeString(bstrMessage);
	return m_nCode;
}
long COrder::OverSeaCancelOrder(string strLogInID, bool bAsyncOrder, string strNo, long nTradeType, int nSeqOrBookOrNum, int nMarket)
{
	string strFullAccount = "";
	bstr_t bstrFullAccount;
	long m_nCode = 0;
	//�T�{�b��
	 if (nMarket == 3 || nMarket == 4)
	{
		if (vec_strFullAccount_OF.size() != 0)
		{
			strFullAccount = vec_strFullAccount_OF[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "OverSeaCancelOrder Error: No OverSea Future/Option Account";
			return -1;
		}
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrNo = strNo.c_str();
	BSTR bstrMessage;
	if (nSeqOrBookOrNum == 0)//�Ǹ��R��
	{
		m_nCode = m_pSKOrderLib->OverSeaCancelOrderBySeqNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount, bstrNo, &bstrMessage);//�Ǹ�
		cout << "OverSeaCancelOrderBySeqNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
	else if (nSeqOrBookOrNum == 1)//�Ѹ��R��
	{
		m_nCode = m_pSKOrderLib->OverSeaCancelOrderByBookNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), bstrFullAccount, bstrNo, &bstrMessage);//�Ѹ�
		cout << "OverSeaCancelOrderByBookNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
}
long COrder::OverSeaCorrectPrice(string strLogInID, bool bAsyncOrder,string strNo,string strExchangeNo,string strStockNo,string strYearMonth,
	string strPrice,string strOrderNumbertor,string strOrderDenominator,short sTradeType,short sSpecialTradeType,short sNewClose,short sCallPut,string strStrikePrice,int nSeqOrBook,int nMarket)
{
	string strFullAccount = "", strMarketSymbol = "",strSeqNo="";
	bstr_t bstrFullAccount;
	long m_nCode = 0;
	//�T�{�b��
	if (nMarket == 3 || nMarket == 4)
	{
		if (vec_strFullAccount_OF.size() != 0)
		{
			strFullAccount = vec_strFullAccount_OF[0];
			bstrFullAccount = strFullAccount.c_str();
		}
		else
		{
			cout << "OverSeaCorrectPrice Error: No Future/Option Account";
			return -1;
		}
		if (nMarket == 3)
		{
			strMarketSymbol = "OF";
		}
		else
		{
			strMarketSymbol = "OO";
		}
	}
	bstr_t bstrLogInID = strLogInID.c_str();
	bstr_t bstrNo = strNo.c_str();
	bstr_t bstrExchange = strExchangeNo.c_str();
	bstr_t bstrStockNo = strStockNo.c_str();
	bstr_t bstrYearMonth = strYearMonth.c_str();
	bstr_t bstrPrice = strPrice.c_str();
	bstr_t bstrOrderNumbertor = strOrderNumbertor.c_str();
	bstr_t bstrStrikePrice = strStrikePrice.c_str();
	bstr_t bstrOrderDenominator = strOrderDenominator.c_str();
	bstr_t bstrSeqNo = strSeqNo.c_str();

	SKCOMLib::OVERSEAFUTUREORDERFORGW pOrder;
	pOrder.bstrFullAccount = bstrFullAccount;
	pOrder.bstrBookNo = bstrNo;
	pOrder.bstrSeqNo = bstrSeqNo;//�Y�ϥήѸ��d�ߡA�Ǹ��٬O�n��쪫���(��="")
	pOrder.bstrExchangeNo = bstrExchange;
	pOrder.bstrStockNo = bstrStockNo;
	pOrder.bstrYearMonth = bstrYearMonth;
	pOrder.bstrOrderPrice = bstrPrice;
	pOrder.bstrOrderNumerator = bstrOrderNumbertor;
	pOrder.nTradeType = sTradeType;
	pOrder.nSpecialTradeType = sSpecialTradeType;
	pOrder.nNewClose = sNewClose;
	pOrder.bstrOrderDenominator = bstrOrderDenominator;
	BSTR bstrMessage;
	if (nMarket == 3)//�������
	{
		m_nCode = m_pSKOrderLib->OverSeaCorrectPriceByBookNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder), &pOrder, &bstrMessage);//�Ǹ�
		cout << "OverSeaCorrectPriceByBookNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
	else if (nMarket == 4)//������
	{
		pOrder.nCallPut = sCallPut;
		pOrder.bstrStrikePrice = bstrStrikePrice;
		m_nCode = m_pSKOrderLib->OverSeaOptionCorrectPriceByBookNo(bstrLogInID, VARIANT_BOOL(bAsyncOrder),&pOrder, &bstrMessage);//�Ѹ�
		cout << "OverSeaOptionCorrectPriceByBookNo : " << string(_bstr_t(bstrMessage)) << endl;
		::SysFreeString(bstrMessage);
		return m_nCode;
	}
}
//-------�\��禡-------
long COrder::GetUserAccount()//���^����b��
{
	return  m_pSKOrderLib->GetUserAccount();
}
long COrder::ReadCertByID(string strLogInID)//Ū���������
{
	bstr_t bstrLogInID = strLogInID.c_str();
	return m_pSKOrderLib->ReadCertByID(bstrLogInID);
}
long COrder::SetMaxQty(long nMarkettype,long nQuantity)//�C��e�U�q����
{
	return m_pSKOrderLib->SetMaxQty(nMarkettype, nQuantity);
}
void COrder::OnAsyncOrder(long nThreadID, long nCode, string bstrMessage)
{
	cout << "�iOnAsyncOrder�jThread ID " << nThreadID << " , nCode " << nCode << ",bstrMessage" << bstrMessage;
}
void COrder::OnAccount(string strLogInID,string strAccountData)//�^�Ǳb����T
{
	cout << "�iOnAccount�jstrLogInID " << strLogInID << "�istrAccount�j " << strAccountData<<"\n";
	
	const char* pcCut1 = ",";
	char* pcResult;
	int num = 0;
	string strArray[100] ;
	char cArray[100];

	strcpy(cArray, strAccountData.c_str());
	pcResult = strtok(cArray, pcCut1);
	while (pcResult!=NULL)
	{
		strArray[num] = pcResult;
		pcResult = strtok(NULL, pcCut1);
		num++;
	}
	if (num==7 && strArray[0] == "TS")//�P�_�Ҩ�b��
	{
		string strFullAccount = strArray[1] + strArray[3];
		vec_strFullAccount_TS.push_back(strFullAccount);//�s�JvectorTS��
		cout << "�iOnAccount�jTSFullAccount: " << strFullAccount << "\n";//�L�X�Ҩ�b���W 11�X
	}
	else if (num == 7 && strArray[0] == "TF")//��+��
	{
		string strFullAccount = strArray[1] + strArray[3];
		vec_strFullAccount_TF.push_back(strFullAccount);
		cout << "�iOnAccount�jTFFullAccount: " << strFullAccount << "\n";
	}
	else if (num == 7 && strArray[0] == "OF")//���~���f�B����v
	{
		string strFullAccount = strArray[1] + strArray[3];
		vec_strFullAccount_OF.push_back(strFullAccount);
		cout << "�iOnAccount�jOFFullAccount: " << strFullAccount << "\n";
	}
}
long COrder::LoadOSCommodity()
{
	short nCode = m_pSKOrderLib->SKOrderLib_LoadOSCommodity();//�U�������ӫ~��
	cout << "0:��l�Ʀ��\�A�N�X��: " << nCode << endl;
	return nCode;
}
long COrder::LoadOOCommodity()
{
	short nCode = m_pSKOrderLib->SKOrderLib_LoadOOCommodity();//�U�������ӫ~��
	cout << "0:��l�Ʀ��\�A�N�X��: " << nCode << endl;
	return nCode;
}
