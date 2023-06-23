"""
    TsPost For Windows (FileUploader) v2.4.0 [Release]
    Designed by KQY & XTS-X The Student  
    Windows Upload Section
    2023-05-07
    Written By KQY
"""

"""
    NetWork Version Introduce:
    {
        "Name": "TSPost For Windows [2.4.0 Release]",
        "Type": "Release",
        "Force": true, 
        "Version": "latest",
        "NewVersion": "2.4.0",
        "DownloadCmd": "https://www.tsginkgo.cn/file-uploader/api/download/windows?edition=release",
        "MoreUrl": "https://www.tsginkgo.cn/file-uploader/about/",
        "UpdateDetails": [
            {"type": "新增", "text": "添加了每次发送Post请求时的ApiKey验证"},
            {"type": "新增", "text": "添加了查看更新内容"},
            {"type": "新增", "text": "添加了可直接从程序中下载并安装新版本"},
            {"type": "修正", "text": "修正了一些逻辑错误"},
            {"type": "修正", "text": "修正了强制更新不生效的错误"},
            {"type": "修正", "text": "修正了版本检测部分的错误"},
            {"type": "改善", "text": "改善了检测ApiKey的过程"},
            {"type": "改善", "text": "大幅改善了版本检测的冗余部分"},
            {"type": "改善", "text": "大幅改善整体逻辑与排布，减小了文件大小"}
        ]
    }

    [新增] 添加了每次发送Post请求时的ApiKey验证
    [新增] 添加了查看更新内容
    [新增] 添加了可直接从程序中下载并安装新版本
    [修正] 修正了一些逻辑错误
    [修正] 修正了强制更新不生效的错误
    [修正] 修正了版本检测部分的错误
    [改善] 改善了检测ApiKey的过程
    [改善] 大幅改善了版本检测的冗余部分
    [改善] 大幅改善整体逻辑与排布，减小了文件大小
"""

# Import Packages Needed
    # Main Packages
from requests import get, post;
from os import path, listdir, makedirs;
from configparser import ConfigParser;
from json import loads;
from webbrowser import open as webopen;
from windnd import hook_dropfiles;
from time import time;
from subprocess import Popen;
from tkinter.ttk import Progressbar;
from threading import Thread;
    # GUI Packages
import tkinter as tk;
from tkinter import simpledialog;
from tkinter.messagebox import showinfo, askquestion;
from tkinter.filedialog import askopenfilename, askdirectory;

# Init The Variables
global NowVersion, VerSionType, MaxFileSize, Domain, RootDict, DownloadDict;
global VersionCheckUrl, LearnMoreUrl, KeyCheckUrl, FileUploadUrl;
global IniConfig, IniConfigPath; # For Release
NowVersion = "2.4.0";
VerSionType = "Release";
Domain = "https://www.tsginkgo.cn";
VersionCheckUrl = f"{Domain}/file-uploader/api/programs/windows/NewVersion.json";
LearnMoreUrl = f"{Domain}/file-uploader/about/";
KeyCheckUrl = f"{Domain}/file-uploader/api/upload/public/keyCheck/";
FileUploadUrl = f"{Domain}/file-uploader/api/upload/public/";
MaxFileSize = 52428800;     # 50Mb Max
RootDict = path.dirname(__file__);
DownloadDict = f"{RootDict}/download";

if not path.exists(DownloadDict):
    makedirs(DownloadDict);

# Read The Settings
IniConfig = ConfigParser();
IniConfigPath = r'{}\Settings\ms.ini'.format(RootDict);
IniConfig.read(IniConfigPath);

def VersionCheck(Type="IN"):
    """
        SoftWare Version Check Main Function
    """
    try:
        Req = get(VersionCheckUrl).text;
        # print(Req)
        Req = loads(Req);
        
        for items in Req:
            if ((items["Type"]==VerSionType) and (items["Version"]=="latest")):
                NewVersion = items["NewVersion"];
                DownloadCmd = items["DownloadCmd"];
                MoreUrl = items["MoreUrl"];
                Force = items["Force"];
                NewVersionFileName = items["Name"];
                UpdateDetails = "\n相较上版本：";
                for details in items["UpdateDetails"]:
                    UpdateDetails += f"""\n\t[{details["type"]}] {details["text"]}""";
            
                if (not VersionCompare(NowVersion, NewVersion)):
                    # print("Need Update");
                    print("-"*64+f"\n你的 TSPost For Windows [{VerSionType}] 需要更新！\n当前版本:{NowVersion}，最新版本:{NewVersion}");
                    if (Force):
                        print(f"新版本为强制更新版本，你必须更新到最新版本才能继续使用\n你可以通过访问'{DownloadCmd}'来获得更新\n在{MoreUrl}了解更多"+"-"*64);
                        showinfo('检测到更新——强制更新', f'当前版本:{NowVersion}，最新版本:{NewVersion}\n新版本为强制更新版本，你必须更新到最新版本才能继续使用\n你可以通过访问{DownloadCmd}来获得更新\n在{MoreUrl}了解更多{UpdateDetails}');
                        IfDownload = askquestion("立即下载新版本", "是否立即下载新版本并安装");
                        if (IfDownload=="yes"):
                            try:
                                DownloadNewVersion(NewVersionFileName);
                            except Exception as error:
                                showinfo("error", f"自动下载失败，请手动下载\n{error}");
                        return False;
                    else:
                        print(f"新版本为不强制更新版本，你可以继续使用此旧版本\n你可以通过访问'{DownloadCmd}'来获得更新\n在{MoreUrl}了解更多"+"-"*64);
                        showinfo('检测到更新——不强制更新', f'当前版本:{NowVersion}，最新版本:{NewVersion}\n新版本为不强制更新版本，你可以继续使用此旧版本\n你可以通过访问{DownloadCmd}来获得更新\n在{MoreUrl}了解更多{UpdateDetails}');
                        IfDownload = askquestion("立即下载新版本", "是否立即下载新版本并安装");
                        if (IfDownload=="yes"):
                            try:
                                DownloadNewVersion(NewVersionFileName);
                            except Exception as error:
                                showinfo("error", f"自动下载失败，请手动下载\n{error}");
                        return True;
                else:
                    print("-"*64);
                    print(f"TSPost For Windows [{VerSionType}] {NowVersion}");
                    print("-"*64);
                    if (Type=="IN"):
                        showinfo('更新检测——最新版', f"TSPost For Windows [{VerSionType}] {NowVersion}\n你安装的版本为最新版！{UpdateDetails}");
                    return True;
    except Exception as error:
        print(error);
        print("-"*64);
        print(f"TSPost For Windows [{VerSionType}] {NowVersion}   OFFLINE");
        print("-"*64);
        showinfo('OFFLINE', f'TSPost For Windows {NowVersion}   OFFLINE\nDetail:{error}');
        return True;  

def VersionCompare(Now_Version, New_Version):
    """
        SoftWare Version Compare (Only For 3points Version) e.g."2.2.1"
    """
    NowMajor, NowMinor, NowPatch = map(int, Now_Version.split('.'));
    NewMajor, NewMinor, NewPatch = map(int, New_Version.split('.'));
    if ((NewMajor, NewMinor, NewPatch) > (NowMajor, NowMinor, NowPatch)):
        # 新版本比现有版本高，需要更新
        return False;
    else:
        # 新版本比现有版本低或相同，不需要更新
        return True;

def DownloadExe(url, FileName):
    """
        Download New Version Setup (.exe)
    """
    response = get(url, stream=True, allow_redirects=True);
    TotalSize = int(response.headers.get("content-length", 0));
    BlockSize = 1024;  # 分片下载
    try:
        ProgressWindow = tk.Tk();
    except:
        ProgressWindow = tk.Toplevel(window);
    ProgressWindow.title("更新");
    ProgressWindow.geometry("400x150");
    ProgressBar = Progressbar(ProgressWindow, orient="horizontal", length=200, mode="determinate");
    ProgressBar.pack(pady=10)
    
    SpeedLabel = tk.Label(ProgressWindow, text="下载速度：0 KB/s\n0KB/0KB");
    SpeedLabel.pack();
    ProgressLabel = tk.Label(ProgressWindow, text="下载进度：0%");
    ProgressLabel.pack();
    DownloadedSize = 0;
    StartTime = None;
    with open(f"{DownloadDict}/{FileName} Setup.exe", "wb") as f:
        for data in response.iter_content(BlockSize):
            if StartTime is None:
                StartTime = time();
            DownloadedSize += len(data);
            Progress = DownloadedSize * 100 // TotalSize;
            ProgressBar["value"] = Progress;
            ProgressLabel["text"] = f"下载进度：{Progress:.1f}%";
            ElapsedTime = time() - StartTime;
            try:
                Speed = DownloadedSize / ElapsedTime / 1024;
            except:
                Speed = 0;
            SpeedLabel["text"] = f"{FileName}\n下载速度：{Speed:.1f} KB/s\n{DownloadedSize/1024**2:.1f}MB/{TotalSize/1024**2:.1f}MB";
            ProgressWindow.update();
            f.write(data);
    ProgressWindow.destroy();
    Popen(f"{DownloadDict}/{FileName} Setup.exe");

def DownloadNewVersion(FileName):
    """
        Start Download New Version Setup (.exe)
    """
    url = f"https://www.tsginkgo.cn/file-uploader/api/download/windows/files/{FileName} Setup.exe";
    url.replace(" ", "%20");
    # print(url)
    t = Thread(target=DownloadExe, args=(url, FileName, ));
    t.start();

def OpenAbout():
    """
        Open Our Website To Find More Details
    """
    webopen(LearnMoreUrl);

def SelectFile():
    """
        Select File To Upload
    """
    try:
        FilePath = askopenfilename(title='选择文件',filetypes=[('任意文件','*')])
        DraggedFiles(FilePath, "search");
    except Exception as e:
        # print(e)
        # showinfo('error', e);
        pass;

def SelectDict():
    """
        Select Dict To Upload
    """
    try:
        DictPath = askdirectory(title='选择文件夹');
        DraggedFiles(DictPath, "search");
    except Exception as e:
        # showinfo('error', e);
        pass;

def SearchFileInDict(BasePath): # Def Search Lambda
    """
        Output Files When Drag Dict
    """
    for item in listdir(BasePath):   # Show All Iiems In Dict
        TempPath = path.join(BasePath, item);    # Add LevelOnePath To LevelTwoPath
        if path.isfile(TempPath):    # If File 
            DictFile.append(TempPath);   # -> Append To The List
        else:   # If Dict
            SearchFileInDict(TempPath);   # Rerun The Lambda
    return DictFile; # Return DictFile List

def DraggedFiles(Files, Type='drag'):    # Def Dragged Lambda
    """
        Drag File(& Dict) To Upload
    """
    if (Type=='drag'):
        Path = "".join((item.decode('gbk') for item in Files));  # Use Gbk To Correct File Name
    else:
        Path = Files;
    if (path.isdir(Path)): # If Dict
        global DictFile; # -> Set DictFile List Global
        DictFile = [];   # -> Init DictFile List
        SearchFileInDict(Path);   # -> Seacrh Files In Dict
        Num = 0; # Init File Num
        for Item in DictFile:
            if (path.getsize(Item)>MaxFileSize):
                showinfo("OK", f'文件{Item}超过最大大小限制50MB！');
            else:
                Result = FileUploadPost(Path);  # Run
                print(Result);   # Print Result
                Num += 1;    # Add File Num
        showinfo("OK", f'{Num}个文件\n状态:{Result}');    # Display Info
    else:   # If File
        if (path.getsize(Path)>MaxFileSize):
            showinfo("OK", f'文件{Path}超过最大大小限制50MB！');
        else:
            Result = FileUploadPost(Path);  # Run Directly
            print(Result);   # Print Result
            showinfo("OK", f'文件{Path}\n状态:{Result}');    # Display Info

def GetSettingData():
    """
        Get The Setting Data From Settings/ms.ini
    """
    try:
        UploaderKey = IniConfig.get("release", "apikey");
        WindowWidth = IniConfig.get("settings", "window_width");
        WindowHeight = IniConfig.get("settings", "window_height");
        return {"UploaderKey": UploaderKey, "WindowWidth": WindowWidth, "WindowHeight":WindowHeight};
    except Exception as error:
        try:
            UploaderKey = IniConfig.get("release", "apikey");
            return {"UploaderKey": UploaderKey, "WindowWidth": "400", "WindowHeight": "400", "ReadConfigError": error};
        except Exception as error:
            return {"UploaderKey": "", "WindowWidth": "400", "WindowHeight": "400", "ReadConfigError": error};
    
def SetAPIKey():
    """
        Update The Api Key
    """
    try:
        NowKey = GetSettingData()["UploaderKey"];
        # APIKey = tk.simpledialog.askstring(title='设置APIKEY', initialvalue=NowKey, prompt='输入您的APIKEY\n当前KEY:{}'.format(NowKey));
        APIKey = simpledialog.askstring('设置APIKEY', '输入您的APIKEY', initialvalue=NowKey);
        IniConfig.set("release", "apikey", APIKey);
        with open(IniConfigPath, "w") as f:
            IniConfig.write(f);
        showinfo('Success', '您的API KEY已设置成功');
    except Exception as e:
        # print(e)
        SetError = e;
        # pass;
    ApiCheck = CheckAPIKey();
    if (ApiCheck["code"] == 0):
        showinfo("success", "API KEY校验成功！");
        usershow.config(text="欢迎您，{}".format(ApiCheck["Name"]));
    else:
        showinfo("error", f"你的API KEY有误！\n{SetError}");
        usershow.config(text="欢迎您，来访者");

def CheckAPIKey():
    """
        Check The Api Key From The Server
    """
    key = GetSettingData()["UploaderKey"];
    data = {'UploaderKey': key};
    res = post(KeyCheckUrl, data=data).text;
    return loads(res);

def FileUploadPost(FilePath=None, Description=None):
    """
        Post File(s) To The Server Main Function
    """
    UploaderKey = GetSettingData()["UploaderKey"];
    if ((not FilePath) or (not UploaderKey)):
        return {"code": -1, "error": '请现在选项>配置>API KEY中配置你的API KEY'};

    if (path.getsize(FilePath)>MaxFileSize):
        return {"code": -1, "error": f'文件{FilePath}超过最大大小限制50MB！'};

    try:
        data = {'UploaderKey': UploaderKey, 'Description': Description, 'PostFrom': r'{"Name":"TsPost For Windows", "Type":"'+VerSionType+'", "Version":"'+NowVersion+'", "OfficialKey":"CFC0CE1DF6A5916DF3B11497C399B4F3"}'};
        files = {'file': open(FilePath, 'rb')};

        res = post(FileUploadUrl, data=data, files=files).text; # requests.post
        try:
            return loads(res);  # 将服务器返回的数据转换成json格式
        except:
            return res;
    except Exception as error:
        return {"code": -1, "error": str(error)};

if (__name__ == "__main__"):
    # Tk GUI Init
    window = tk.Tk();
    window.title(f'TSPost for Windows {NowVersion} [{VerSionType}]');
    InitData = GetSettingData();
    try:
        window.geometry(f"{InitData['WindowWidth']}x{InitData['WindowHeight']}");
    except:
        window.geometry('400x100');
    
    # Set the window icon if it exists
    icon_path = path.join(RootDict, 'Settings', 'logo.ico');
    if (path.exists(icon_path)):
        window.iconbitmap(icon_path);

    # Create and pack labels
    usershow = tk.Label(window, width=40, text='欢迎您，来访者');
    usershow.pack();
    tk.Label(window, width=40, text='可通过打开或拖拽的方式进行上传').pack();
    tk.Label(window, width=40, text=f'Release Edition {NowVersion}').pack();

    # Create the menu bar
    menubar = tk.Menu(window);

    # Create the "Options" menu
    filemenu = tk.Menu(menubar, tearoff=0);
    filemenu.add_command(label="设置ApiKey", command=SetAPIKey);
    filemenu.add_separator();
    filemenu.add_command(label='检查更新', command=VersionCheck);
    filemenu.add_command(label='关于', command=OpenAbout);
    filemenu.add_command(label='退出', command=window.quit);
    menubar.add_cascade(label='选项', menu=filemenu);

    # Create the "Upload" menu
    uploadmenu = tk.Menu(menubar, tearoff=0);
    uploadmenu.add_command(label="文件", command=SelectFile);
    uploadmenu.add_command(label="文件夹", command=SelectDict);
    menubar.add_cascade(label='上传', menu=uploadmenu);

    # Add the menu bar to the window
    window.config(menu=menubar);

    # Dragging section
    hook_dropfiles(window, func=DraggedFiles);

    # Version check
    Vc = VersionCheck(Type="OUT");
    if (not Vc):
        window.quit;
        exit();
    
    ApiCheck = CheckAPIKey();
    if ApiCheck["code"] == 0:
        # showinfo("success", "API KEY校验成功！");
        usershow.config(text="欢迎您，{}".format(ApiCheck["Name"]));
    
    # Start the main loop
    window.mainloop();