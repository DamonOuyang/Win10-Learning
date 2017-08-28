#include <windows.h>
#include <atlstr.h>
#include <sphelper.h>
#include <sapi.h>
#include <comutil.h>
#include <string.h>


void openGame();
void closeGame(); // �����������÷���

#pragma comment(lib,"sapi.lib")
#pragma comment(lib, "comsupp.lib") 

#define GID_CMD_GR 333333
#define WM_RECOEVENT WM_USER+1

 LRESULT CALLBACK WndProc(HWND,UINT,WPARAM,LPARAM); // �۲���ģʽ + �ص���ʽ -> ʶ������

 char 	 szAppName[] = "Dimon";  // �õ� com �ӿ� ����Ԥ��һЩ�ӿ�
 BOOL b_initSR;
 BOOL b_Cmd_Grammar;
 CComPtr<ISpRecoContext>m_cpRecoCtxt;  //����ʶ�����ӿ�
 CComPtr<ISpRecoGrammar>m_cpCmdGramma; //ʶ���﷨
 CComPtr<ISpRecognizer>m_cpRecoEngine; //����ʶ������
 int speak(wchar_t *str);

 int WINAPI WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,PSTR szCmdLine,int iCmdShow)
 {
	 HWND        hwnd;
	 MSG         msg;
	 WNDCLASS    wndclass;

	 wndclass.cbClsExtra          =0;
	 wndclass.cbWndExtra          =0;
	 wndclass.hbrBackground       =(HBRUSH)GetStockObject(WHITE_BRUSH);
	 wndclass.hCursor             =LoadCursor(NULL,IDC_ARROW);
	 wndclass.hIcon               =LoadIcon(NULL,IDI_APPLICATION);
	 wndclass.hInstance           =hInstance;
	 wndclass.lpfnWndProc         =WndProc;
	 wndclass.lpszClassName       =szAppName;
	 wndclass.lpszMenuName        =NULL;
	 wndclass.style               =CS_HREDRAW|CS_VREDRAW;

	 if(!RegisterClass(&wndclass)) 
	 {
		 MessageBox(NULL,TEXT("This program requires Windows NT!"),szAppName,MB_ICONERROR);
		 return 0;
	 }
	 //speak(L"���ˣ�����"); // ˵����
	 //speak(L"Dimon ����");
	 //speak(L"Alen ���");
	 speak(L"���� ���");
	 hwnd=CreateWindow(szAppName,    // ����һ������
		               TEXT("C/C++����ʶ��̳�"),
					   WS_OVERLAPPEDWINDOW,
					   CW_USEDEFAULT,
					   CW_USEDEFAULT,
					   CW_USEDEFAULT,
					   CW_USEDEFAULT,
					   NULL,
					   NULL,
					   hInstance,
					   NULL);

	 ShowWindow(hwnd,iCmdShow);
	 UpdateWindow(hwnd);
	 
	 while(GetMessage(&msg,NULL,0,0))  // ��ȡ��Ϣѭ��
	 {
		 TranslateMessage(&msg);
		 DispatchMessage(&msg);
	 }
	
	 return msg.wParam;
 }

 LRESULT CALLBACK WndProc(HWND hwnd,UINT message,WPARAM wParam,LPARAM lParam)
 {
	 HDC           hdc;
	 PAINTSTRUCT   ps;

	 switch(message)
	 {
	 case WM_CREATE:  // ������ʱ��
		 {
			 //��ʼ��COM�˿�
			 ::CoInitializeEx(NULL,COINIT_APARTMENTTHREADED);
			 //����ʶ������COMʵ��Ϊ������
			 HRESULT hr=m_cpRecoEngine.CoCreateInstance(CLSID_SpSharedRecognizer);
			 //����ʶ�������Ľӿ�
			 if(SUCCEEDED(hr))
			 {
				 hr=m_cpRecoEngine->CreateRecoContext(&m_cpRecoCtxt);
			 }
			 else MessageBox(hwnd,TEXT("error1"),TEXT("error"),S_OK);
			 //����ʶ����Ϣ,ʹ�����ʱ�̼���������Ϣ
			 if(SUCCEEDED(hr))
			 {
				 hr=m_cpRecoCtxt->SetNotifyWindowMessage(hwnd,WM_RECOEVENT,0,0);
			 }
			 else MessageBox(hwnd,TEXT("error2"),TEXT("error"),S_OK);
			 //�������Ǹ���Ȥ���¼�
			 if(SUCCEEDED(hr))
			 {
				 ULONGLONG ullMyEvents=SPFEI(SPEI_SOUND_START)|SPFEI(SPEI_RECOGNITION)|SPFEI(SPEI_SOUND_END);
				 hr=m_cpRecoCtxt->SetInterest(ullMyEvents,ullMyEvents);
			 }
			 else MessageBox(hwnd,TEXT("error3"),TEXT("error"),S_OK);
			 //�����﷨����
			 b_Cmd_Grammar=TRUE;
			 if(FAILED(hr))
			 {
				 MessageBox(hwnd,TEXT("error4"),TEXT("error"),S_OK);
			 }
			 //��¼�������뵽.xml
			 hr=m_cpRecoCtxt->CreateGrammar(GID_CMD_GR,&m_cpCmdGramma);
			 //WCHAR wszXMLFile[20]=L"xx.xml";
			 //MultiByteToWideChar(CP_ACP,0,(LPCSTR)"xx.xml",-1,wszXMLFile,256);
			 //hr = m_cpCmdGramma->LoadCmdFromFile(wszXMLFile, SPLO_DYNAMIC);
			 //if (FAILED(hr))
			 //{
				// MessageBox(hwnd, TEXT("error5"), TEXT("error"), S_OK);
			 //}
			 WCHAR wszXMLFile[20] = L"er.xml";
			 MultiByteToWideChar(CP_ACP, 0, (LPCSTR)"er.xml", -1, wszXMLFile, 256);
			 hr = m_cpCmdGramma->LoadCmdFromFile(wszXMLFile, SPLO_DYNAMIC);
			 if (FAILED(hr))
			 {
				 MessageBox(hwnd, TEXT("error5"), TEXT("error"), S_OK);
			 }
			 b_initSR=TRUE;
			 //�ڿ�ʼʶ��ʱ�������﷨����ʶ��
		     hr=m_cpCmdGramma->SetRuleState(NULL,NULL,SPRS_ACTIVE);
	    	 return 0;
		 }
	 case WM_RECOEVENT: // ʶ��֮�󣬽���һϵ�в���
		 {
			 RECT rect;
             GetClientRect(hwnd,&rect);
             hdc=GetDC(hwnd);
			 USES_CONVERSION;
			 CSpEvent event;
			 while(event.GetFrom(m_cpRecoCtxt)==S_OK)
			 {
			     switch(event.eEventId)
			     {
			     case SPEI_RECOGNITION:
    				 {
            			 static const WCHAR wszUnrecognized[]=L"<Unrecognized>";
		            	 CSpDynamicString dstrText;
			    		 //ȡ��ʶ����
				    	 if(FAILED(event.RecoResult()->GetText(SP_GETWHOLEPHRASE,SP_GETWHOLEPHRASE,TRUE,&dstrText,NULL)))
					     {
						     dstrText=wszUnrecognized;
    					 }
        	    		 BSTR SRout;
	        	    	 dstrText.CopyToBSTR(&SRout);
						 char* lpszText2 = _com_util::ConvertBSTRToString(SRout);

			    		 if(b_Cmd_Grammar)
				    	 {
							 DrawText(hdc, TEXT(lpszText2), -1, &rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER);
							 if (strcmp("���絶",lpszText2)==0)
	    				     {    
								 keybd_event('A', 0, 0, 0);//����
								 keybd_event('A', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('D', 0, 0, 0);//����
								 keybd_event('D', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('W', 0, 0, 0);//����
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
					         }
							 if (strcmp("��紵", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('A', 0, 0, 0);//����
								 keybd_event('A', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('D', 0, 0, 0);//����
								 keybd_event('D', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('S', 0, 0, 0);//����
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
								 
							 }
							 if (strcmp("��ն", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('S', 0, 0, 0);//����
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('S', 0, 0, 0);//����
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('W', 0, 0, 0);//����
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

							 }

							 if (strcmp("�Ʒ�", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('W', 0, 0, 0);//����
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('W', 0, 0, 0);//����
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
							 }

							 if (strcmp("��Ѱ", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('W', 0, 0, 0);//����
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('S', 0, 0, 0);//����
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
							 }

							 if (strcmp("����", lpszText2) == 0)
							 {
								 //   \ang
								 keybd_event(VK_OEM_102, 0, 0, 0);//����
								 keybd_event(VK_OEM_102, 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('A', 0, 0, 0);//����
								 keybd_event('A', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('N', 0, 0, 0);//����
								 keybd_event('N', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event('G', 0, 0, 0);//����6igrk������ֲ
								 keybd_event('G', 0, KEYEVENTF_KEYUP, 0);//����
								 Sleep(10);
								 keybd_event(VK_RETURN, 0, 0, 0);//����6igrk������ֲ
								 keybd_event(VK_RETURN,  0, KEYEVENTF_KEYUP, 0);//����							
							 }
							 if (strcmp("��Ծ", lpszText2) == 0)
							 {

								 keybd_event(VK_SPACE, 0, 0, 0);//����
								 keybd_event(VK_SPACE, 0, KEYEVENTF_KEYUP, 0);//����

							 }
							 if (strcmp("ſ��", lpszText2) == 0)
							 {


							 }
							 if (strcmp("ǰ��", lpszText2) == 0)
							 {
								 keybd_event('W', 0, 0, 0);//����
	//							 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//����

							 }
							 if (strcmp("����", lpszText2) == 0)
							 {
								 keybd_event('S', 0, 0, 0);//����
	//							 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//����

							 }
							 if (strcmp("������˭", lpszText2) == 0)
							 {
								 speak(L"�����Ǻ�ճ�����Ӣ�ۣ����겻�������");
							 }
							 if (strcmp("����˭", lpszText2) == 0)
							 {
								 speak(L"�������������ڷܱ뺷�Ĵ��ʹ��д������ʶ�����");
							 }

							 if (strcmp("���Ǳ���", lpszText2) == 0)
							 {
								 speak(L"�ҵĴ����ߴ��ʹ������ľ����������");
			
							 }
							 if (strcmp("���Ǵ���", lpszText2) == 0)
							 {
								 speak(L"�ҹ�Ȼ�ܴ��������ҵĴ����ߴ��ʹ������ľ����������");

							 }
							 if (strcmp("�������", lpszText2) == 0)
							 {
								 speak(L"�����������");
							 }
							 if (strcmp("��", lpszText2) == 0)
							 {
								 speak(L"�װ������ˣ���");
								 openGame();
							 }
							 if (strcmp("�ر�", lpszText2) == 0)
							 {
								 speak(L"�װ������ˣ���");
								 closeGame();
							 }
							 if (strcmp("Ӣ��", lpszText2) == 0)
							 {
								 speak(L"СӢ�ܣ������");
							 }
							 if (strcmp("���", lpszText2) == 0)
							 {
								 speak(L"�װ������ˣ���,������֣�");
							 }
    					 }    
        			 }
	    		 }
			 }
			 return TRUE;
		 }
	 case WM_PAINT:
		 hdc=BeginPaint(hwnd,&ps); 
		 EndPaint(hwnd,&ps);
		 return 0;
	 case WM_DESTROY:
		 PostQuitMessage(0);
		 return 0;
	 }
	 return DefWindowProc(hwnd,message,wParam,lParam);
 }

#pragma comment(lib, "ole32.lib") //CoInitialize CoCreateInstance��Ҫ����ole32.dll   
int speak(wchar_t *str)
{
	 ISpVoice * pVoice = NULL;
	 ::CoInitialize(NULL);
	 //��ȡISpVoice�ӿڣ�   
	 long hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void **)&pVoice);
	 hr = pVoice->Speak(str, 0, NULL);
	 pVoice->Release();
	 pVoice = NULL;
	 //ǧ��Ҫ���ǣ�   
	 ::CoUninitialize();
	 return TRUE;
 }

void openGame()
{
	ShellExecuteA(0, "open", "\"D:\\Program Files (x86)\\Meteor\\config.exe\"", 0, 0, 1);
}

void closeGame()
{
	system("taskkill /f /im config.exe");
}