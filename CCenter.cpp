#include "CCenter.h"

CCenter::CCenter()
{
	m_pSKCenterLib.CreateInstance(__uuidof(SKCOMLib::SKCenterLib));
    m_pSKCenterLibEventHandler = new ISKCenterLibEventHandler(*this, m_pSKCenterLib, &CCenter::OnEventFiringObjectInvoke);
}
CCenter::~CCenter()
{
    if (m_pSKCenterLibEventHandler!=NULL)
    {
        m_pSKCenterLibEventHandler->ShutdownConnectionPoint();
        m_pSKCenterLibEventHandler->Release();
        m_pSKCenterLibEventHandler = NULL;
    }
    if (m_pSKCenterLib!=NULL)
    {
        m_pSKCenterLib->Release();
    }
}
HRESULT CCenter::OnEventFiringObjectInvoke
(
    ISKCenterLibEventHandler* pEventHandler,
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
    case 1: // Event1 event.
        varlValue = (pdispparams->rgvarg)[0];
        OnTimer(V_I4(&varlValue));
        break;
    }

    return S_OK;
}

long CCenter::Login(string strUserID, string strPassword)
{
	const char* bstrUserID = strUserID.c_str();
	const char* bstrPassword = strPassword.c_str();
	return m_pSKCenterLib->SKCenterLib_Login(bstrUserID, bstrPassword);
}
BSTR CCenter::GetReturnCodeMessage(long nCode)//取得訊息代碼
{
	return m_pSKCenterLib->SKCenterLib_GetReturnCodeMessage(nCode);
}
BSTR CCenter::GetLastLogInfo()//取得最後一筆log內容
{
    return m_pSKCenterLib->SKCenterLib_GetLastLogInfo();
}
void CCenter::OnTimer(long nTime)//顯示當下時間
{
	long nHour = nTime / 10000;
	long nMin = (nTime % 10000) / 100;
	long nSec = nTime % 100;
	cout << "\n【OnTime】 " << nHour << " : " << nMin << " : " << nSec;
}
void CCenter::PrintCodeMessage(string strFunction, string strFunctionName, long nCode)//印出錯誤訊息
{
    if (nCode == 0)
    {
        string strMessage = (LPCSTR)(m_pSKCenterLib->SKCenterLib_GetReturnCodeMessage(nCode));
        cout << "【" << strFunction.c_str()<<"】" <<"【" << strFunctionName.c_str() << "】" << "【" << strMessage.c_str() << "】\n";
    }
    else
    {
        string strMessage = (LPCSTR)(m_pSKCenterLib->SKCenterLib_GetReturnCodeMessage(nCode));
        string strMessage2 = (LPCSTR)(m_pSKCenterLib->SKCenterLib_GetLastLogInfo());
        cout << "【" << strFunction.c_str() << "】" << "【" << strFunctionName.c_str() << "】" << "【" << strMessage.c_str() << "】" << "【" << strMessage2.c_str() << "】\n";
    }
}
