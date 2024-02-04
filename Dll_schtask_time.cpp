#define _WIN32_DCOM
#include <windows.h>
#include <iostream>
#include <comdef.h>
#include <taskschd.h>
#include "tlhelp32.h"
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
using namespace std;

int main() {
    // 获取标准输出句柄
    HANDLE han = GetStdHandle(STD_OUTPUT_HANDLE);

    // 使用 COM 初始化
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

    // 设置 COM 安全性
    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);

    // 任务名称
    string wszTaskName = "SystemUpdater";

    // 创建 Task Scheduler 实例
    ITaskService* pService = NULL;
    hr = CoCreateInstance(CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        (void**)&pService);

    // 连接 Task Scheduler 服务
    hr = pService->Connect(_variant_t(), _variant_t(),
        _variant_t(), _variant_t());

    // 获取根文件夹
    ITaskFolder* pRootFolder = NULL;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);

    // 删除同名任务
    pRootFolder->DeleteTask(_bstr_t(wszTaskName.c_str()), 0);

    // 创建任务定义
    ITaskDefinition* pTask = NULL;
    hr = pService->NewTask(0, &pTask);

    // 释放 Task Scheduler 接口
    pService->Release();

    // 设置任务注册信息
    IRegistrationInfo* pRegInfo = NULL;
    hr = pTask->get_RegistrationInfo(&pRegInfo);
    hr = pRegInfo->put_Author(_bstr_t("SystemAdmin"));
    pRegInfo->Release();

    // 获取任务设置
    ITaskSettings* pSettings = NULL;
    hr = pTask->get_Settings(&pSettings);
    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    pSettings->Release();

    // ...

    // 设置触发器（每日触发）
    ITriggerCollection* pTriggerCollection = NULL;
    hr = pTask->get_Triggers(&pTriggerCollection);
    ITrigger* pTrigger = NULL;
    hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);
    pTriggerCollection->Release();

    IDailyTrigger* pDailyTrigger = NULL;
    hr = pTrigger->QueryInterface(IID_IDailyTrigger, (void**)&pDailyTrigger);
    pTrigger->Release();

    // 设置触发器 ID
    hr = pDailyTrigger->put_Id(_bstr_t(L"UpdateTrigger"));

    // 设置触发器开始时间
    hr = pDailyTrigger->put_StartBoundary(_bstr_t(L"2022-01-01T12:00:00"));

    // 设置触发器间隔（每天执行）
    hr = pDailyTrigger->put_DaysInterval(1);
    pDailyTrigger->Release();

    // ...

    // 设置任务动作
    IActionCollection* pActionCollection = NULL;
    hr = pTask->get_Actions(&pActionCollection);
    IAction* pAction = NULL;
    hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    pActionCollection->Release();

    IExecAction* pExecAction = NULL;
    hr = pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);
    pAction->Release();

    // 替换为你的可执行文件路径
    LPWSTR filename = const_cast<wchar_t*>(L"UpdaterExecutable");
    hr = pExecAction->put_Path(filename);
    pExecAction->Release();

    // 注册任务定义
    IRegisteredTask* pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(L"SystemUpdater"),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),
        _variant_t(),
        TASK_LOGON_INTERACTIVE_TOKEN,
        _variant_t(L""),
        &pRegisteredTask);

    // 处理注册失败情况
    if (FAILED(hr)) {
        WriteConsole(han, L"\nError updating the task", 26, new DWORD, 0);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    // 输出成功信息
    WriteConsole(han, L"Update task successful", 24, new DWORD, 0);

    // 释放资源
    pRootFolder->Release();
    pTask->Release();
    pRegisteredTask->Release();
    CoUninitialize();
    return 0;
}
