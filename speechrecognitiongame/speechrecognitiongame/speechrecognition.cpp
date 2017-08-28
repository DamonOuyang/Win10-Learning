#include <windows.h>
#include <atlstr.h>
#include <sphelper.h>
#include <sapi.h>
#include <comutil.h>
#include <string.h>


void openGame();
void closeGame(); // 语音输入后调用方法

#pragma comment(lib,"sapi.lib")
#pragma comment(lib, "comsupp.lib") 

#define GID_CMD_GR 333333
#define WM_RECOEVENT WM_USER+1

 LRESULT CALLBACK WndProc(HWND,UINT,WPARAM,LPARAM); // 观察者模式 + 回调方式 -> 识别语音

 char 	 szAppName[] = "Dimon";  // 用到 com 接口 程序预留一些接口
 BOOL b_initSR;
 BOOL b_Cmd_Grammar;
 CComPtr<ISpRecoContext>m_cpRecoCtxt;  //语音识别程序接口
 CComPtr<ISpRecoGrammar>m_cpCmdGramma; //识别语法
 CComPtr<ISpRecognizer>m_cpRecoEngine; //语音识别引擎
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
	 //speak(L"主人，您好"); // 说出来
	 //speak(L"Dimon 您好");
	 //speak(L"Alen 你好");
	 speak(L"张朕 你好");
	 hwnd=CreateWindow(szAppName,    // 创建一个窗口
		               TEXT("C/C++语音识别教程"),
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
	 
	 while(GetMessage(&msg,NULL,0,0))  // 获取消息循环
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
	 case WM_CREATE:  // 创建的时候
		 {
			 //初始化COM端口
			 ::CoInitializeEx(NULL,COINIT_APARTMENTTHREADED);
			 //创建识别引擎COM实例为共享型
			 HRESULT hr=m_cpRecoEngine.CoCreateInstance(CLSID_SpSharedRecognizer);
			 //创建识别上下文接口
			 if(SUCCEEDED(hr))
			 {
				 hr=m_cpRecoEngine->CreateRecoContext(&m_cpRecoCtxt);
			 }
			 else MessageBox(hwnd,TEXT("error1"),TEXT("error"),S_OK);
			 //设置识别消息,使计算机时刻监听语音消息
			 if(SUCCEEDED(hr))
			 {
				 hr=m_cpRecoCtxt->SetNotifyWindowMessage(hwnd,WM_RECOEVENT,0,0);
			 }
			 else MessageBox(hwnd,TEXT("error2"),TEXT("error"),S_OK);
			 //设置我们感兴趣的事件
			 if(SUCCEEDED(hr))
			 {
				 ULONGLONG ullMyEvents=SPFEI(SPEI_SOUND_START)|SPFEI(SPEI_RECOGNITION)|SPFEI(SPEI_SOUND_END);
				 hr=m_cpRecoCtxt->SetInterest(ullMyEvents,ullMyEvents);
			 }
			 else MessageBox(hwnd,TEXT("error3"),TEXT("error"),S_OK);
			 //创建语法规则
			 b_Cmd_Grammar=TRUE;
			 if(FAILED(hr))
			 {
				 MessageBox(hwnd,TEXT("error4"),TEXT("error"),S_OK);
			 }
			 //记录语音输入到.xml
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
			 //在开始识别时，激活语法进行识别
		     hr=m_cpCmdGramma->SetRuleState(NULL,NULL,SPRS_ACTIVE);
	    	 return 0;
		 }
	 case WM_RECOEVENT: // 识别之后，进行一系列操作
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
			    		 //取得识别结果
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
							 if (strcmp("旋风刀",lpszText2)==0)
	    				     {    
								 keybd_event('A', 0, 0, 0);//按下
								 keybd_event('A', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('D', 0, 0, 0);//按下
								 keybd_event('D', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('W', 0, 0, 0);//按下
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
					         }
							 if (strcmp("大风吹", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('A', 0, 0, 0);//按下
								 keybd_event('A', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('D', 0, 0, 0);//按下
								 keybd_event('D', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('S', 0, 0, 0);//按下
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
								 
							 }
							 if (strcmp("雷斩", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('S', 0, 0, 0);//按下
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('S', 0, 0, 0);//按下
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('W', 0, 0, 0);//按下
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

							 }

							 if (strcmp("破防", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('W', 0, 0, 0);//按下
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('W', 0, 0, 0);//按下
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
							 }

							 if (strcmp("难寻", lpszText2) == 0)
							 {
								 //ADS 
								 keybd_event('W', 0, 0, 0);//按下
								 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('S', 0, 0, 0);//按下
								 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
								 mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
							 }

							 if (strcmp("充气", lpszText2) == 0)
							 {
								 //   \ang
								 keybd_event(VK_OEM_102, 0, 0, 0);//按下
								 keybd_event(VK_OEM_102, 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('A', 0, 0, 0);//按下
								 keybd_event('A', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('N', 0, 0, 0);//按下
								 keybd_event('N', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event('G', 0, 0, 0);//按下6igrk；）移植
								 keybd_event('G', 0, KEYEVENTF_KEYUP, 0);//弹起
								 Sleep(10);
								 keybd_event(VK_RETURN, 0, 0, 0);//按下6igrk；）移植
								 keybd_event(VK_RETURN,  0, KEYEVENTF_KEYUP, 0);//弹起							
							 }
							 if (strcmp("跳跃", lpszText2) == 0)
							 {

								 keybd_event(VK_SPACE, 0, 0, 0);//按下
								 keybd_event(VK_SPACE, 0, KEYEVENTF_KEYUP, 0);//弹起

							 }
							 if (strcmp("趴下", lpszText2) == 0)
							 {


							 }
							 if (strcmp("前进", lpszText2) == 0)
							 {
								 keybd_event('W', 0, 0, 0);//按下
	//							 keybd_event('W', 0, KEYEVENTF_KEYUP, 0);//弹起

							 }
							 if (strcmp("后退", lpszText2) == 0)
							 {
								 keybd_event('S', 0, 0, 0);//按下
	//							 keybd_event('S', 0, KEYEVENTF_KEYUP, 0);//弹起

							 }
							 if (strcmp("大猛是谁", lpszText2) == 0)
							 {
								 speak(L"大猛是横空出世的英雄，百年不遇的天才");
							 }
							 if (strcmp("你是谁", lpszText2) == 0)
							 {
								 speak(L"我是你们勇猛勤奋彪悍的大猛大哥写的语音识别程序");
							 }

							 if (strcmp("你是笨蛋", lpszText2) == 0)
							 {
								 speak(L"我的创造者大猛大哥聪明的惊天地泣鬼神");
			
							 }
							 if (strcmp("你是蠢猪", lpszText2) == 0)
							 {
								 speak(L"我固然很蠢，但是我的创造者大猛大哥聪明的惊天地泣鬼神");

							 }
							 if (strcmp("新年快乐", lpszText2) == 0)
							 {
								 speak(L"那来个红包呗");
							 }
							 if (strcmp("打开", lpszText2) == 0)
							 {
								 speak(L"亲爱的主人，好");
								 openGame();
							 }
							 if (strcmp("关闭", lpszText2) == 0)
							 {
								 speak(L"亲爱的主人，好");
								 closeGame();
							 }
							 if (strcmp("英杰", lpszText2) == 0)
							 {
								 speak(L"小英杰，你好吗？");
							 }
							 if (strcmp("你好", lpszText2) == 0)
							 {
								 speak(L"亲爱的主人，好,新年快乐！");
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

#pragma comment(lib, "ole32.lib") //CoInitialize CoCreateInstance需要调用ole32.dll   
int speak(wchar_t *str)
{
	 ISpVoice * pVoice = NULL;
	 ::CoInitialize(NULL);
	 //获取ISpVoice接口：   
	 long hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void **)&pVoice);
	 hr = pVoice->Speak(str, 0, NULL);
	 pVoice->Release();
	 pVoice = NULL;
	 //千万不要忘记：   
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