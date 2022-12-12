// PunchDemo_Excell.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "resource.h"

HINSTANCE hInst;                                // instance actuelle
CHAR szTitle[MAX_LOADSTRING];                  // Texte de la barre de titre
INITCOMMONCONTROLSEX iccex; 
WNDCLASS wc;
NOTIFYICONDATAA nf;
SYSTEMTIME st;
DWORD version=MAKEWORD(21,1);
HWND hWnd;
int x=0;

char jours[7][10] = {"dimanche", "lundi","mardi","mercredi","jeudi","vendredi","samedi"};
char mois[12][10] = {"janvier", "fevrier","mars", "avril", "mai", "juin","juillet","aout","septembre", "octobre", "novembre", "decembre"};
char buffer[MAX_PATH];
char tmp[10];

BOOL CALLBACK    ProcedurePrincipale(HWND, UINT, WPARAM, LPARAM);
BOOL CALLBACK    ProcedurePropos(HWND, UINT, WPARAM, LPARAM);
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...);
void GetDate();
char* GetTime();
int MsgBox(HWND hDlg,char* lpszText,char* lpszCaption, DWORD dwStyle,int lpszIcon);
void EcrireDatas(char* texte,int ligne,int colone);
void PeuplerLaListe();

struct DatasXLS{
	char date[0xC8];
	char heure[0x08];
	char nom[0x30];
	char scratchDir[MAX_PATH];
} xls;

int APIENTRY WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,LPSTR     lpCmdLine,int       nCmdShow){
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	iccex.dwICC=ICC_WIN95_CLASSES;
	iccex.dwSize=sizeof(iccex);
	InitCommonControlsEx(&iccex);
	GetCurrentDirectory(0xFF,xls.scratchDir);
	memset(&wc,0,sizeof(wc));
	wc.hCursor=LoadCursor(hInstance, (LPCTSTR)IDC_CURSOR1);
	wc.lpfnWndProc = DefDlgProc;
	wc.cbWndExtra = DLGWINDOWEXTRA;
	wc.hInstance = hInstance;
	wc.lpszClassName ="PunchDemo_Excell";
	wc.hbrBackground=(HBRUSH)CreateSolidBrush(RGB(0xCC,0xCC,0xCC));
	wc.hIcon=LoadIcon(hInstance,(LPCTSTR)IDI_ICON2);
	wc.style = CS_VREDRAW  | CS_HREDRAW | CS_SAVEBITS | CS_DBLCLKS;
	RegisterClass(&wc);
	return DialogBox(hInstance, (LPCTSTR)IDD_DIALOG1, NULL, (DLGPROC)ProcedurePrincipale);
}

BOOL CALLBACK ProcedurePrincipale(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	hWnd =hDlg;
	switch (message) {
		case WM_INITDIALOG:{
//			AnimateWindow(hDlg, 2000, AW_ACTIVATE);
			GetDate();
			SetTimer(hDlg,IDM_TIMER1,1000,(TIMERPROC) NULL);
			RemoveMenu(GetSystemMenu(hDlg, FALSE), SC_CLOSE, MF_STRING);
			RemoveMenu(GetSystemMenu(hDlg, FALSE), SC_MOVE, MF_STRING);
			AppendMenu(GetSystemMenu(hDlg,FALSE),MF_STRING, IDC_BUTTON1,"A propos de ce programme..");
			AppendMenu(GetSystemMenu(hDlg,FALSE),MF_STRING,0xE140,"Quel Windows s'execute?");
			AppendMenu(GetSystemMenu(hDlg,FALSE),MF_STRING,2,"Fermer ce programme");
			SendMessage(GetDlgItem(hDlg,IDCANCEL), WM_SETFONT, (WPARAM)GetStockObject(0x1E), MAKELPARAM(TRUE, 0));
			SendMessage(GetDlgItem(hDlg,IDOK), WM_SETFONT, (WPARAM)GetStockObject(0x1F), MAKELPARAM(TRUE, 0));
			nf.cbSize=sizeof(nf);
			nf.hIcon=wc.hIcon;
			nf.hWnd=hDlg;
			strcpy(nf.szTip,"PunchDemo_Excelm\0");
			nf.uCallbackMessage=WM_TRAY_ICONE;
			nf.uFlags=NIF_ICON|NIF_MESSAGE|NIF_TIP;
			Shell_NotifyIcon(NIM_ADD,&nf);
			HWND imghWnd = CreateWindowEx(0, "STATIC", NULL, WS_VISIBLE|WS_CHILD|SS_ICON,1, 1, 10, 10,hDlg , (HMENU)45000, wc.hInstance, NULL);
			SendMessage(imghWnd, STM_SETIMAGE, IMAGE_ICON, (LPARAM)LoadIcon(wc.hInstance,(LPCTSTR)IDI_ICON1));
			SetWindowText(hDlg,szTitle);
			GetLocalTime(&st);
			CreateStatusWindow(WS_CHILD|WS_VISIBLE,__argv[0],hDlg,6000);
			sprintf(buffer,"Nous sommes %s, le %2d %s %4d",jours[st.wDayOfWeek],st.wDay,mois[st.wMonth-1],st.wYear); //creation du string de date
			SetDlgItemText(hDlg,IDC_DATE,buffer);
			PeuplerLaListe();
#ifdef _WIN64
			SetClassLong(hDlg, GCLP_HICON, (long)LoadIcon(wc.hInstance, MAKEINTRESOURCE(IDI_ICON1)));
#else
			SetClassLong(hDlg, GCL_HICON, (long)LoadIcon(wc.hInstance, MAKEINTRESOURCE(IDI_ICON1)));
#endif // WIN64
			SendMessage(GetDlgItem(hDlg,IDOK), WM_SETFONT, (WPARAM)GetStockObject(2), 1L);
			SendMessage(GetDlgItem(hDlg,IDCANCEL), WM_SETFONT, (WPARAM)GetStockObject(10), 1L);

					   }return TRUE;
	case WM_SYSCOMMAND:{
		switch (LOWORD(wParam)) {
			case 0xE140:ShellAbout(hDlg,wc.lpszClassName,"Demo Punch utilisant Excell pour stocker les donnees\n© Patrice Waechter-Ebling 2022",wc.hIcon);
				break;
			case IDC_BUTTON1:
				DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG2), hWnd, ProcedurePropos);
				break;
			case IDCANCEL:{
					Shell_NotifyIcon(NIM_DELETE,&nf); 
					KillTimer(hDlg,IDM_TIMER1);
					PostQuitMessage(1);
						  }return 0x01;
				}		
					   }break;	
	case WM_COMMAND:{
		switch (LOWORD(wParam)) {
			case IDC_BUTTON1:
				DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG2), hWnd, ProcedurePropos);
				break;
			case IDC_BUTTON2:{
				SendMessage(GetDlgItem(hDlg,IDC_COMBO1),CB_GETLBTEXT,(WPARAM)SendMessage(GetDlgItem(hDlg,IDC_COMBO1),CB_GETCURSEL ,0,0),(LPARAM)xls.nom);
				x++;
				EcrireDatas(xls.nom,x,1);
				EcrireDatas(GetTime(),x,2);
							 }break;
			case IDC_BUTTON3:{
				SendMessage(GetDlgItem(hDlg,IDC_COMBO1),CB_GETLBTEXT,(WPARAM)SendMessage(GetDlgItem(hDlg,IDC_COMBO1),CB_GETCURSEL ,0,0),(LPARAM)xls.nom);
				x++;
				EcrireDatas(xls.nom,x,1);
				EcrireDatas(GetTime(),x,3);
							 }break;

			
			case IDCANCEL:{
					Shell_NotifyIcon(NIM_DELETE,&nf); 
					KillTimer(hDlg,IDM_TIMER1);
					PostQuitMessage(1);
						  }return 0x01;
			}
							 }break;
		case WM_TIMER:{
			GetLocalTime(&st);
			sprintf(buffer,"Nous sommes %s, le %2d %s %4d",jours[st.wDayOfWeek],st.wDay,mois[st.wMonth-1],st.wYear);
			SetDlgItemText(hDlg,IDC_DATE,buffer);
			sprintf(buffer,"%.2d:%.2d:%.2d",st.wHour,st.wMinute,st.wSecond);
			SetDlgItemText(hDlg,IDC_HEURE,buffer);
					  }break;
		case WM_CTLCOLORDLG:		return (long)wc.hbrBackground;
		case WM_CTLCOLORSTATIC:{
			SetTextColor( (HDC)wParam, RGB(0, 0,255) );
			SetBkMode( (HDC)wParam, TRANSPARENT ); 
						   }return (LONG)wc.hbrBackground; //retourne la couleur choisie dans la classe
		case WM_CTLCOLORLISTBOX: {
			SetTextColor( (HDC)wParam, RGB(255,255, 255) );
			SetBkMode( (HDC)wParam, TRANSPARENT ); 
							 }return (LONG)(HBRUSH)CreateSolidBrush(RGB(0,0,0));
		case WM_CTLCOLORBTN:{
			SetBkMode( (HDC)wParam, TRANSPARENT );
			switch(LOWORD(wParam)){
			case IDOK:{SetTextColor((HDC)wParam, RGB(0, 255, 0));}break;
			}
				if (GetDlgItem(hWnd, IDCANCEL) == reinterpret_cast<HWND>(lParam)) SetTextColor((HDC)wParam, RGB(255, 0, 0)); 
				if (GetDlgItem(hWnd, 6000) == reinterpret_cast<HWND>(lParam)) SetTextColor((HDC)wParam, RGB( 0, 0,255)); 
						}return (LONG)(HBRUSH)wc.hbrBackground;
		case WM_CTLCOLORMSGBOX:{
			SetTextColor( (HDC)wParam, RGB(0,255, 255) );
						   }return (long)(HBRUSH)CreateSolidBrush(RGB(255,0,128));
		case WM_CLOSE: {
			Shell_NotifyIcon(NIM_DELETE,&nf);
			KillTimer(hDlg,IDM_TIMER1);
			PostQuitMessage(1);
				   }return  0x01;
	}return FALSE;
 
}
BOOL CALLBACK ProcedurePropos(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam){
	switch (message)	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;
	case WM_CTLCOLORDLG:		return (long)CreateSolidBrush(RGB(0xCC, 0xCC, 0x50));
	case WM_CTLCOLORSTATIC: {
		SetTextColor((HDC)wParam, RGB(0, 0, 255));
		SetBkMode((HDC)wParam, TRANSPARENT);
	}return (LONG)CreateSolidBrush(RGB(0xCC, 0xCC, 0x50));
	case WM_CTLCOLORBTN: {
		SetBkMode((HDC)wParam, TRANSPARENT);
		switch (LOWORD(wParam)) {
		case IDOK: {SetTextColor((HDC)wParam, RGB(0, 255, 0)); }break;
		}
		if (GetDlgItem(hWnd, IDCANCEL) == reinterpret_cast<HWND>(lParam)) SetTextColor((HDC)wParam, RGB(255, 0, 0));
	}return (LONG)(HBRUSH)wc.hbrBackground;
	case WM_CTLCOLORMSGBOX: {
		SetTextColor((HDC)wParam, RGB(0, 255, 255));
	}return (long)(HBRUSH)CreateSolidBrush(RGB(255, 0, 128));

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...) {
  va_list marker;
  va_start(marker, cArgs);
  if(!pDisp) {
      MsgBox(hWnd, "Aucun IDispatch envoyé vers AutoWrap()", "Erreur", 0x00,IDI_ICON3);
      _exit(0);
  }
  DISPPARAMS dp = { NULL, NULL, 0, 0 };
  DISPID dispidNamed = DISPID_PROPERTYPUT;
  DISPID dispID;
  HRESULT hr;
  char buf[200];
  char szName[200];
  WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);
  hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
  if(FAILED(hr)) {
      sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") en erreur w/err 0x%08lx", szName, hr);
      MsgBox(hWnd, buf, "AutoWrap()", 0x00,IDI_ICON3);
      _exit(0);
      return hr;
  }
  VARIANT *pArgs = new VARIANT[cArgs+1];
  for(int i=0; i<cArgs; i++) {
      pArgs[i] = va_arg(marker, VARIANT);
  }
  dp.cArgs = cArgs;
  dp.rgvarg = pArgs;
  if(autoType & DISPATCH_PROPERTYPUT) {
      dp.cNamedArgs = 1;
      dp.rgdispidNamedArgs = &dispidNamed;
  }
  hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
  if(FAILED(hr)) {
      sprintf(buf, "IDispatch::Invocation (\"%s\"=%08lx) éronnée w/err 0x%08lx", szName, dispID, hr);
      MsgBox(hWnd, buf, "AutoWrap()", 0x00,IDI_ICON3);
      _exit(0);
      return hr;
  }
  va_end(marker);
  delete [] pArgs;
  return hr;
}
void GetDate(){
	GetLocalTime(&st);
	wsprintf(xls.date,"%4d%.2d%.2d",st.wYear,st.wMonth,st.wDay);
}
char* GetTime(){
	GetLocalTime(&st);
	wsprintf(tmp,"%.2d%:2d%:2d",st.wHour,st.wMinute,st.wSecond);
	return tmp;
}
int MsgBox(HWND hDlg,char* lpszText,char* lpszCaption, DWORD dwStyle,int lpszIcon){
	MSGBOXPARAMS lpmbp;
	lpmbp.hInstance=wc.hInstance;
	lpmbp.cbSize=sizeof(MSGBOXPARAMS);
	lpmbp.hwndOwner=hDlg;
	lpmbp.dwLanguageId=MAKELANGID(0x0800,0x0800); //par defaut celui du systeme
	lpmbp.lpszText=lpszText;
	if(lpszCaption!=NULL){
		lpmbp.lpszCaption=lpszCaption;
	}else{
		lpmbp.lpszCaption=szTitle;
	}
	lpmbp.dwStyle=dwStyle|0x00000080L;
	lpmbp.lpszIcon=(LPCTSTR)lpszIcon;
	lpmbp.lpfnMsgBoxCallback=0;
	return  MessageBoxIndirect(&lpmbp);
}
void EcrireDatas(char* texte,int ligne,int colone){
	VARIANT root[64] = {0};
	VARIANT parm[64] = {0};
	VARIANT rVal = {0};
	int level=0;
	OleInitialize(NULL);
	VARIANT app = {0}; 
	CLSID clsid;
	CLSIDFromProgID(L"Excel.Application", &clsid);
	HRESULT hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER|CLSCTX_INPROC_SERVER, IID_IDispatch, (void **)&rVal.pdispVal);
	if(FAILED(hr)) {
	char buf[256];
		sprintf(buf, "CoCreateInstance() pour la requete \"Excel.Application\" a échoué. Err=%08lx", hr);
		MsgBox(hWnd, buf, "Erreur d'instance", 0x00,IDI_ICON3);
		_exit(0);
	}
	rVal.vt = VT_DISPATCH;
	VariantCopy(&app, &rVal);
	VariantClear(&rVal);
	rVal.vt = VT_I4;
	rVal.lVal = 1;
	VariantCopy(&root[++level], &app);
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, root[level].pdispVal, L"visible", 1, rVal);
	VariantClear(&root[level--]);
	VariantClear(&rVal);
	VariantCopy(&root[++level], &app);
	AutoWrap(DISPATCH_PROPERTYGET|DISPATCH_METHOD, &root[level+1], root[level++].pdispVal, L"workbooks", 0);
	AutoWrap(DISPATCH_METHOD, NULL, root[level].pdispVal, L"add", 0);
	VariantClear(&root[level--]);
	VariantClear(&root[level--]);
	rVal.vt = VT_BSTR;
	rVal.bstrVal = ::SysAllocString((OLECHAR*)texte);
	VariantCopy(&root[++level], &app);
	AutoWrap(DISPATCH_PROPERTYGET|DISPATCH_METHOD, &root[level+1], root[level++].pdispVal, L"activesheet", 0);
	parm[0].vt = VT_I4; parm[0].lVal = ligne;
	parm[1].vt = VT_I4; parm[1].lVal = colone;
	AutoWrap(DISPATCH_PROPERTYGET|DISPATCH_METHOD, &root[level+1], root[level++].pdispVal, L"cells", 2, parm[1], parm[0]);
	VariantClear(&parm[0]);
	VariantClear(&parm[1]);
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, root[level].pdispVal, L"value", 1, rVal);
	VariantClear(&root[level--]);
	VariantClear(&root[level--]);
	VariantClear(&root[level--]);
	VariantClear(&rVal);
	VariantClear(&app);
	OleUninitialize();
}
void PeuplerLaListe(){
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Eleni Gjoka");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Edouard Witkowski");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Edgart Krieg");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Dominik Mohns");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Denis Steinmetz");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Claus Becker");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Claude Huffschmitt");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Claude Chartran");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Clarissa Selzer");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Christian Isaac Meneses Vela");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Uwe Bergmann");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Benjamin Magnussen");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Azzedine Harrouche");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Corine Allel ");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Ali Keskin");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Alexandre Comtois");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Albert Leroux");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Albert Kropp");
	SendMessage(GetDlgItem(hWnd,IDC_COMBO1),CB_ADDSTRING, 0,(LPARAM)"Alain Riehl");
}