'***************************************************************************
'*
'* MODULE NAME:     Protected VBA project Picklock(PVP)
'* AUTHOR & DATE:   tt.t
'*                  23 April 2007
'* E-Mail:          ttui(ＡＴ)163.com, sohu邮箱垃圾邮件太多已经不用了
'*
'* Usage:           运行FrmHookMain窗口,点补丁,然后双击工程窗口中有密码保护的模块
'*                  应该能够直接打开了:)
'*
'*
'* DESCRIPTION:     在写中文字符串转换为拼音函数(HzToPy)过程中,第一次发现VBA功能的强大.
'*                  于是这次尝试将其他语言中比较好写的API HOOK移植成VBA代码,
'*                  正好顺便把VBA密码保护去掉,喜欢加密码的朋友不要生气啊:)
'*                  总的来说VBA的写法和其他语言区别不大,但VBA毕竟不太方便,代码必须放在标准模块中.
'*                  再有就是对指针的支持实在有限,于是最后选择了一种写起来最简单的API hook方法,
'*                  就是所谓的陷阱法.如果你不太清楚什么是API HOOK,请求助于google.
'*
'* Theory：         这里就不说API hook的方法了,都是传统方法没什么可说的,这里只
'*                  简单说下VBA模块密码破解.其实这些我也不是很了解,毕竟知道加密过程
'*                  用处不大,这个问题上我比较关心结果:)
'*                  判断有密码以及提示输入密码都是VBE6.dll干得好事.如果有密码,
'*                  VBE6.dll会调用DialogBoxParamA显示VB6INTL.dll资源中的第4070号
'*                  对话框(就是那个输入密码的窗口),若DialogBoxParamA返回值非0,
'*                  则VBE会认为密码正确,然后乖乖展开加密模块的资源.很显然其中存在很大
'*                  漏洞,就像给日记本加上了锁,但里面全是活页,我们不需要打开锁,只要从侧面
'*                  取出活页就可以了.这个从侧面取活页的过程就是hook住DialogBoxParamA函数,
'*                  若程序调用DialogBoxParamA装入4070号对话框,我们就直接返回1,让
'*                  VBE以为密码正确.
'*
'* PS:              PVP是在一个叫Advanced VBA Password Recovery (AVPR)的软件启发下
'*                  作出来的,AVPR提供了一个VBA Backdoor功能就是跳过密码直接查看工程资源.
'*                  它的原理和PVP一样,但用了通用性比较差的方法,适用系统比较有限,而PVP的方法
'*                  理论上适用于所有采用第4070号对话框录入密码的Office系统.
'*                  经测试PVP适用于Office 2002, 2003, 2007,其他版本尚未测试,但估计依然有效.
'*                  在2000和XP系统上测试通过,但条件限制没有在Vista系统上测试,听说Vista有些机制
'*                  可能影响API hook,暂时没机会测试就先这样吧~
'*
'*                  *64位操作系统下面的API hook代码肯定运行出错,就不要测试了
'*
'*
'***************************************************************************Option Explicit
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Long, Source As Long, ByVal Length As Long)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long   
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long   
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
ByVal lpProcName As String) As Long  
Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
ByVal pTemplateName As Long, ByVal hWndParent As Long, _
ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer   
Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim pFunc As Long
Dim Flag As Boolean

Private Function GetPtr(ByVal Value As Long) As Long
'获得函数的地址
    GetPtr = Value
End Function

Public Sub RecoverBytes()
'若已经hook,则恢复原API开头的6字节,也就是恢复原来函数的功能
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function Hook() As Boolean
Dim TmpBytes(0 To 5) As Byte
Dim p As Long
Dim OriginProtect As Long   
Hook = False  
'VBE6.dll调用DialogBoxParamA显示VB6INTL.dll资源中的第4070号对话框(就是输入密码的窗口)
    '若DialogBoxParamA返回值非0,则VBE会认为密码正确,所以我们要hook DialogBoxParamA函数
    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
'标准api hook过程之一: 修改内存属性,使其可写
    If VirtualProtect(ByVal pFunc, 6, &H40, OriginProtect) <> 0 Then
'标准api hook过程之二: 判断是否已经hook,看看API的第一个字节是否为&H68,
        '若是则说明已经Hook
        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
If TmpBytes(0) <> &H68 Then
'标准api hook过程之三: 保存原函数开头字节,这里是6个字节,以备后面恢复
            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6
'用AddressOf获取MyDialogBoxParam的地址
            '因为语法不允许写成p = AddressOf MyDialogBoxParam,这里我们写一个函数
            'GetPtr,作用仅仅是返回AddressOf MyDialogBoxParam的值,从而实现将
            'MyDialogBoxParam的地址付给p的目的
            p = GetPtr(AddressOf MyDialogBoxParam)
'标准api hook过程之四: 组装API入口的新代码
            'HookBytes 组成如下汇编
            'push MyDialogBoxParam的地址
            'ret
            '作用是跳转到MyDialogBoxParam函数
            HookBytes(0) = &H68
MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
HookBytes(5) = &HC3
'标准api hook过程之五: 用HookBytes的内容改写API前6个字节
            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
'设置hook成功标志
            Flag = True
Hook = True
End If
End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As Long, _
ByVal pTemplateName As Long, ByVal hWndParent As Long, _
ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
If pTemplateName = 4070 Then
'有程序调用DialogBoxParamA装入4070号对话框,这里我们直接返回1,让
        'VBE以为密码正确了
        MyDialogBoxParam = 1
Else
'有程序调用DialogBoxParamA,但装入的不是4070号对话框,这里我们调用
        'RecoverBytes函数恢复原来函数的功能,在进行原来的函数
        RecoverBytes
MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
hWndParent, lpDialogFunc, dwInitParam)
'原来的函数执行完毕,再次hook
        Hook
End If
End Function