/**
 * 此方法是根据页面调起WPS的SDK做的调用方法封装
 * 可参照此定义
 * @param {*} funcs     这是在WPS加载项内部定义的方法，采用JSON格式（先方法名，再参数）
 */
function _WpsStartUp(funcs) {
    var info = {};

    info.funcs = funcs;

    WpsStartUp.StartUp(WpsStartUp.ClientType.wps, // 组件类型
        "WpsOAAssist", // 插件名，与wps客户端加载的加载的插件名对应
        "dispatcher", // 插件方法入口，与wps客户端加载的加载的插件代码对应，详细见插件代码
        info, // 传递给插件的数据
        function (result) { // 调用回调，status为0为成功，其他是错误
            if (result.status)
                alert(result.message)
            else
                console.log(result.response)
        })
}

/**
 * 这是页面中针对代码显示的变量定义，开发者无需关心
 */
var _wps = {}


// 从这里开始是新建文档
function newDoc() {
    var uploadPath = 'http://127.0.0.1:3888/Upload'
    
    _WpsStartUp([{
        "NewDoc": {
            "uploadFieldName": {
                "docKey": "123"
            },
            "uploadPath": uploadPath,
            "newFileName": "1234"
        }
    }]) // NewDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['newDoc'] = {
    action: newDoc,
    detail: "\n\
  说明：\n\
    点击按钮，打开WPS文字后,新建一个空白doc文档\n\
\n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行插件中的js函数NewDoc,新建一个空白doc\n\
    funcs参数说明:\n\
        NewDoc方法对应于OA助手dispatcher支持的方法名\n\
"
}

// 从这里开始是在线打开文档
function onlineEditDoc() {
    // var filePath = prompt("请输入打开文件路径（本地或是url）：", GetDemoPath("样章.docx"))
    // var uploadPath = prompt("请输入文档上传路径:", GetUploadPath())
    // var uploadFieldName = prompt("请输入文档上传到业务系统时自定义字段：", "自定义字段")

    var filePath = 'http://127.0.0.1:3888/file/11111_2020071733333.docx'
    var uploadPath = 'http://127.0.0.1:3888/Upload'

    _WpsStartUp([{
        "OnlineEditDoc": {
            "docId": "666", // 文档ID
            "attachId": "attachId",
            "uploadPath": uploadPath, // 保存文档上传路径
            "fileName": filePath,
            "uploadFieldName": "123",
            "buttonGroups": "btnSaveAsFile,btnImportDoc", //屏蔽功能按钮
            "userName": "liuxq",
            "revisionCtrl": {
                "bOpenRevision": true,
                "bShowRevision": true
            },
            "notifyUrl": "http://127.0.0.1:3388/notify"
        }
    }]) // onlineEditDoc方法对应于OA助手dispatcher支持的方法名
}

_wps['onlineEditDoc'] = {
    action: onlineEditDoc,
    detail: "\n\
  说明：\n\
    点击按钮，输入要打开的文档路径，输入文档上传路径，如果传的不是有效的服务端地址，将无法使用保存上传功能。\n\
    打开WPS文字后,将根据文档路径在线打开对应的文档，保存将自动上传指定服务器地址\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
        onlineEditDoc方法对应于OA助手dispatcher支持的方法名\n\
            docId 文档ID\n\
            uploadPath 保存文档上传路径\n\
            fileName 打开的文档路径\n\
            uploadFieldName 文档上传到业务系统时自定义字段\ n\
            buttonGroups 屏蔽的OA助手功能按钮\n\
            userName 传给wps要显示的OA用户名\n\
"
}

function insertRedHeader() {
    // var filePath = prompt("请输入打开文件路径，如果为空则对活动文档套红：", GetDemoPath("样章.docx"))
    // var templateURL = prompt("请输入红头模板路径（本地或是url）:", GetDemoPath("红头文件.docx"))

    var filePath = 'http://127.0.0.1:3888/file/sinosure.docx'
    var uploadPath = 'http://127.0.0.1:3888/Upload'
    var templateURL = 'http://127.0.0.1:3888/file/红头文件.doc'

    if (filePath != '' && filePath != null) {
        _WpsStartUp([{
            "OnlineEditDoc": {
                "docId": "123", // 文档ID
                "fileName": filePath,
                "uploadPath": uploadPath, // 保存文档上传路径
                "insertFileUrl": templateURL,
                "bkInsertFile": 'Content', //红头模板中填充正文的位置书签名
                "buttonGroups": "btnSaveAsFile,btnImportDoc,btnPageSetup,btnInsertDate,btnSelectBookmark", //屏蔽功能按钮
                "redFileElement": {"briefingyear": 11, "copyReceiverDeptName": "办公室", "finishTime": "2020-09-30", "reportReceiverDeptName": "信息技术部"}
            }
        }])
    } else {
        _WpsStartUp([{
            "InsertRedHead": {
                "insertFileUrl": templateURL,
                "bkInsertFile": 'Content', //红头模板中填充正文的位置书签名
            }
        }])
    }
}

_wps['insertRedHeader'] = {
    action: insertRedHeader,
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开指定文档，然后使用指定红头模板对该文档进行套红头\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        docId 文档ID\n\
        fileName 打开的文档路径\n\
        insertFileUrl 指定的红头模板\n\
        bkInsertFile 红头模板中正文的位置书签名\n\
    InsertRedHead方法对应于OA助手dispatcher支持的方法名\n\
        insertFileUrl 指定的红头模板\n\
        bkInsertFile 红头模板中正文的位置书签名\n\
"
}

function protectOpen() {

    var filePath = 'http://127.0.0.1:3888/file/sinosure.docx'
    var uploadPath = 'http://127.0.0.1:3888/Upload'

    _WpsStartUp([{
        "OpenDoc": {
            "docId": "123", // 文档ID
            "uploadPath": uploadPath, // 保存文档上传路径
            "fileName": filePath,
            "buttonGroups": "btnSaveAsFile,btnSaveToServer,btnImportDoc,btnPageSetup,btnInsertDate,btnSelectBookmark", //屏蔽功能按钮
            "openType": { //文档打开方式
                // 文档保护类型，-1：不启用保护模式，0：只允许对现有内容进行修订，
                // 1：只允许添加批注，2：只允许修改窗体域(禁止拷贝功能)，3：只读
                "protectType": 3
            }
        }
    }])
}


_wps['protectOpen'] = {
    action: protectOpen,
    code: _WpsStartUp.toString() + "\n\n" + protectOpen.toString(),
    detail: "\n\
  说明：\n\
    点击按钮，输入参数后，打开WPS文字后，打开使用保护模式指定文档\n\
    \n\
  方法使用：\n\
    页面点击按钮，通过wps客户端协议来启动WPS，调用oaassist插件，执行传输数据中的指令\n\
    funcs参数信息说明:\n\
    OpenDoc方法对应于OA助手dispatcher支持的方法名\n\
        uploadPath 保存文档上传路径\n\
        docId 文档ID\n\
        fileName 打开的文档路径\n\
        openType 文档打开方式控制参数 protectType：1：不启用保护模式，0：只允许对现有内容进行修订，\n\
        \t\t1：只允许添加批注，2：只允许修改窗体域(禁止拷贝功能)，3：只读 password为密码\n\
"
}

/** 
 * 这是HTML页面上的按钮赋予事件的实现，开发者无需关心，使用自己习惯的方式做开发即可
 */
window.onload = function () { 

    var btn1 = document.getElementById('newDoc');
    var btn2 = document.getElementById('onlineEditDoc');
    var btn3 = document.getElementById('insertRedHeader');
    var btn4 = document.getElementById('protectOpen');
    var onBtnAction1 = _wps['newDoc'].action;
    var onBtnAction2 = _wps['onlineEditDoc'].action;
    var onBtnAction3 = _wps['insertRedHeader'].action;
    var onBtnAction4 = _wps['protectOpen'].action;

    btn1.onclick = function () {
        var xhr = new WpsStartUp.CreateXHR();
        xhr.onload = function () {
            onBtnAction1()
        }
        xhr.onerror = function () {
            alert("请确认本地服务端(StartupServer.js)是启动状态")
            return
        }
        xhr.open('get', 'http://127.0.0.1:3888/FileList', true)
        xhr.send()        
    }

    btn2.onclick = function () {
        var xhr = new WpsStartUp.CreateXHR();
        xhr.onload = function () {
            onBtnAction2()
        }
        xhr.onerror = function () {
            alert("请确认本地服务端(StartupServer.js)是启动状态")
            return
        }
        xhr.open('get', 'http://127.0.0.1:3888/FileList', true)
        xhr.send()
    }

    btn3.onclick = function () {
        var xhr = new WpsStartUp.CreateXHR();
        xhr.onload = function () {
            onBtnAction3()
        }
        xhr.onerror = function () {
            alert("请确认本地服务端(StartupServer.js)是启动状态")
            return
        }
        xhr.open('get', 'http://127.0.0.1:3888/FileList', true)
        xhr.send()
    }

    btn4.onclick = function () {
        var xhr = new WpsStartUp.CreateXHR();
        xhr.onload = function () {
            onBtnAction4()
        }
        xhr.onerror = function () {
            alert("请确认本地服务端(StartupServer.js)是启动状态")
            return
        }
        xhr.open('get', 'http://127.0.0.1:3888/FileList', true)
        xhr.send()
    }
}