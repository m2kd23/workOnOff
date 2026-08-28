using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

public class DetectedItem
{
    public string Type { get; set; } // "File" or "Folder"
    public string FilePath { get; set; }
    public string FileName { get; set; }
    public string Extension { get; set; }
    public string PreferredApplication { get; set; }
    public bool Exists { get; set; }
    public bool IsSelected { get; set; }

    public DetectedItem()
    {
        IsSelected = true;
    }
}

public static class FileDetector
{
    // COM 관련 P/Invoke
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out System.Runtime.InteropServices.ComTypes.IRunningObjectTable pprot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out System.Runtime.InteropServices.ComTypes.IBindCtx ppbc);

    public static List<DetectedItem> DetectAllOpenedItems()
    {
        Logger.Write("열린 업무파일 탐색 시작");
        List<DetectedItem> result = new List<DetectedItem>();
        HashSet<string> uniquePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 1. MS Office 및 한글 파일 탐색 (COM)
        DetectOfficeAndHwp(result, uniquePaths);

        // 2. Windows 탐색기 폴더 탐색 (Shell API)
        DetectExplorerFolders(result, uniquePaths);

        // 3. 기타 파일 형식 (PDF, TXT, CSV 등) 탐색 (MainWindowTitle + Recent Files + Common Folders)
        DetectOtherFiles(result, uniquePaths);

        Logger.Write(string.Format("탐색 완료: 총 {0}개 항목 발견", result.Count));
        return result;
    }

    private static void DetectOfficeAndHwp(List<DetectedItem> result, HashSet<string> uniquePaths)
    {
        // Excel 탐색
        try
        {
            dynamic excel = Marshal.GetActiveObject("Excel.Application");
            if (excel != null)
            {
                Logger.Write("Excel Application COM 연결 성공");
                foreach (dynamic wb in excel.Workbooks)
                {
                    try
                    {
                        string path = wb.FullName;
                        if (!string.IsNullOrEmpty(path) && Path.IsPathRooted(path))
                        {
                            AddFileItem(path, "excel.exe", result, uniquePaths);
                        }
                    }
                    catch {}
                }
            }
        }
        catch (COMException) { /* Excel이 실행 중이 아님 */ }
        catch (Exception ex) { Logger.Write("Excel 탐색 중 오류: " + ex.Message, "ERROR"); }

        // Word 탐색
        try
        {
            dynamic word = Marshal.GetActiveObject("Word.Application");
            if (word != null)
            {
                Logger.Write("Word Application COM 연결 성공");
                foreach (dynamic doc in word.Documents)
                {
                    try
                    {
                        string path = doc.FullName;
                        if (!string.IsNullOrEmpty(path) && Path.IsPathRooted(path))
                        {
                            AddFileItem(path, "winword.exe", result, uniquePaths);
                        }
                    }
                    catch {}
                }
            }
        }
        catch (COMException) { /* Word가 실행 중이 아님 */ }
        catch (Exception ex) { Logger.Write("Word 탐색 중 오류: " + ex.Message, "ERROR"); }

        // PowerPoint 탐색
        try
        {
            dynamic ppt = Marshal.GetActiveObject("PowerPoint.Application");
            if (ppt != null)
            {
                Logger.Write("PowerPoint Application COM 연결 성공");
                foreach (dynamic pres in ppt.Presentations)
                {
                    try
                    {
                        string path = Path.Combine(pres.Path, pres.Name);
                        if (!string.IsNullOrEmpty(path) && Path.IsPathRooted(path))
                        {
                            AddFileItem(path, "powerpnt.exe", result, uniquePaths);
                        }
                    }
                    catch {}
                }
            }
        }
        catch (COMException) { /* PowerPoint가 실행 중이 아님 */ }
        catch (Exception ex) { Logger.Write("PowerPoint 탐색 중 오류: " + ex.Message, "ERROR"); }

        // 한글 (HWP) 탐색
        try
        {
            dynamic hwp = Marshal.GetActiveObject("HWPFrame.HwpObject");
            if (hwp != null)
            {
                Logger.Write("한글(HWP) COM 연결 성공");
                dynamic docs = hwp.XHwpDocuments;
                int count = docs.Count;
                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        dynamic doc = docs.Item(i);
                        string path = doc.FullName;
                        if (!string.IsNullOrEmpty(path) && Path.IsPathRooted(path))
                        {
                            AddFileItem(path, "hwp.exe", result, uniquePaths);
                        }
                    }
                    catch {}
                }
            }
        }
        catch (COMException) { /* 한글이 실행 중이 아님 */ }
        catch (Exception ex) { Logger.Write("한글(HWP) 탐색 중 오류: " + ex.Message, "ERROR"); }
    }

    private static void DetectExplorerFolders(List<DetectedItem> result, HashSet<string> uniquePaths)
    {
        try
        {
            Type shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null) return;

            dynamic shell = Activator.CreateInstance(shellType);
            dynamic windows = shell.Windows();
            if (windows == null) return;

            int count = windows.Count;
            for (int i = 0; i < count; i++)
            {
                try
                {
                    dynamic window = windows.Item(i);
                    if (window == null) continue;

                    string fullName = window.FullName;
                    if (fullName != null && fullName.ToLower().EndsWith("explorer.exe"))
                    {
                        string locationUrl = window.LocationURL;
                        if (!string.IsNullOrEmpty(locationUrl))
                        {
                            Uri uri = new Uri(locationUrl);
                            if (uri.IsFile)
                            {
                                string localPath = uri.LocalPath;
                                if (Directory.Exists(localPath))
                                {
                                    if (!uniquePaths.Contains(localPath))
                                    {
                                        uniquePaths.Add(localPath);
                                        result.Add(new DetectedItem
                                        {
                                            Type = "Folder",
                                            FilePath = localPath,
                                            FileName = Path.GetFileName(localPath),
                                            Extension = "",
                                            PreferredApplication = "explorer.exe",
                                            Exists = true
                                        });
                                        Logger.Write("탐색기 폴더 감지: " + localPath);
                                    }
                                }
                            }
                        }
                    }
                }
                catch {}
            }
        }
        catch (Exception ex)
        {
            Logger.Write("탐색기 폴더 탐색 중 오류: " + ex.Message, "ERROR");
        }
    }

    private static void DetectOtherFiles(List<DetectedItem> result, HashSet<string> uniquePaths)
    {
        try
        {
            // 감시 대상 프로세스 정의
            string[] targetProcessNames = { "notepad", "AcroRd32", "Acrobat", "msedge", "chrome" };
            Process[] processes = Process.GetProcesses();

            foreach (Process proc in processes)
            {
                try
                {
                    bool isTarget = false;
                    foreach (string target in targetProcessNames)
                    {
                        if (proc.ProcessName.Equals(target, StringComparison.OrdinalIgnoreCase))
                        {
                            isTarget = true;
                            break;
                        }
                    }
                    if (!isTarget) continue;

                    string title = proc.MainWindowTitle;
                    if (string.IsNullOrEmpty(title)) continue;

                    string fileName = ExtractFileNameFromTitle(title);
                    if (string.IsNullOrEmpty(fileName)) continue;

                    // 이미 다른 곳에서 탐지된 파일인지 확인 (파일명 기준 중복 검사)
                    // 예: "보고서.xlsx"가 Excel COM으로 이미 들어가 있다면 생략
                    bool alreadyExists = false;
                    foreach (var path in uniquePaths)
                    {
                        if (Path.GetFileName(path).Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            alreadyExists = true;
                            break;
                        }
                    }
                    if (alreadyExists) continue;

                    // 실제 파일 경로 찾기
                    string fullPath = FindActualFilePath(fileName);
                    if (!string.IsNullOrEmpty(fullPath))
                    {
                        if (!uniquePaths.Contains(fullPath))
                        {
                            string procExe = proc.ProcessName.ToLower() + ".exe";
                            AddFileItem(fullPath, procExe, result, uniquePaths);
                            Logger.Write(string.Format("기타 파일 감지 (프로세스 {0}): {1}", proc.ProcessName, fullPath));
                        }
                    }
                    else
                    {
                        // 경로를 찾지 못했더라도 파일명은 있으므로 목록에 가상으로 노출 (Exists = false로 둠)
                        string virtualPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), fileName);
                        if (!uniquePaths.Contains(virtualPath))
                        {
                            uniquePaths.Add(virtualPath);
                            result.Add(new DetectedItem
                            {
                                Type = "File",
                                FilePath = virtualPath, // 임시 가상 경로
                                FileName = fileName,
                                Extension = Path.GetExtension(fileName).ToLower(),
                                PreferredApplication = proc.ProcessName.ToLower() + ".exe",
                                Exists = false
                            });
                            Logger.Write(string.Format("기타 파일 감지(경로 미확인): {0}", fileName), "WARN");
                        }
                    }
                }
                catch {}
            }
        }
        catch (Exception ex)
        {
            Logger.Write("기타 파일 탐색 중 오류: " + ex.Message, "ERROR");
        }
    }

    private static void AddFileItem(string path, string preferredApp, List<DetectedItem> result, HashSet<string> uniquePaths)
    {
        if (uniquePaths.Contains(path)) return;
        uniquePaths.Add(path);

        result.Add(new DetectedItem
        {
            Type = "File",
            FilePath = path,
            FileName = Path.GetFileName(path),
            Extension = Path.GetExtension(path).ToLower(),
            PreferredApplication = preferredApp,
            Exists = File.Exists(path)
        });
    }

    private static string ExtractFileNameFromTitle(string title)
    {
        if (string.IsNullOrEmpty(title)) return null;

        // 파일 이름 및 확장자 패턴 (공백, 한글, 특수문자 대응)
        // 파일명은 \ / : * ? " < > | 문자를 포함할 수 없음
        Regex regex = new Regex(@"([^\\/:*?""<>|\r\n]+)\.(pdf|txt|csv|hwp|hwpx|xlsx|xls|docx|doc|pptx|ppt)", RegexOptions.IgnoreCase);
        Match match = regex.Match(title);
        if (match.Success)
        {
            // 엣지나 크롬의 경우 창 제목에 브라우저 이름 등이 포함되어 올 수 있으므로 트림 처리
            string val = match.Value.Trim();
            // 종종 "파일명.pdf - Adobe Acrobat Reader" 에서 "파일명.pdf" 만 온전하게 검출되도록 함
            return val;
        }
        return null;
    }

    private static string FindActualFilePath(string fileName)
    {
        // 1. 최근 문서 (Recent Files) 조회
        string path = FindPathFromRecent(fileName);
        if (!string.IsNullOrEmpty(path)) return path;

        // 2. 주요 폴더 (바탕화면, 다운로드, 문서) 조회
        path = FindPathInCommonFolders(fileName);
        if (!string.IsNullOrEmpty(path)) return path;

        return null;
    }

    private static string FindPathFromRecent(string fileName)
    {
        try
        {
            string recentPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Windows\Recent");
            if (!Directory.Exists(recentPath)) return null;

            string queryName = Path.GetFileNameWithoutExtension(fileName).ToLower();
            string ext = Path.GetExtension(fileName).ToLower();

            string[] files = Directory.GetFiles(recentPath, "*.lnk");
            foreach (string lnk in files)
            {
                try
                {
                    string name = Path.GetFileNameWithoutExtension(lnk).ToLower();
                    if (name.Contains(queryName))
                    {
                        string target = ResolveShortcut(lnk);
                        if (!string.IsNullOrEmpty(target))
                        {
                            if (Path.GetFileName(target).Equals(fileName, StringComparison.OrdinalIgnoreCase))
                            {
                                return target;
                            }
                        }
                    }
                }
                catch {}
            }
        }
        catch {}
        return null;
    }

    private static string ResolveShortcut(string lnkPath)
    {
        try
        {
            Type shellType = Type.GetTypeFromProgID("Wscript.Shell");
            if (shellType == null) return null;

            dynamic shell = Activator.CreateInstance(shellType);
            dynamic shortcut = shell.CreateShortcut(lnkPath);
            string target = shortcut.TargetPath;

            if (!string.IsNullOrEmpty(target) && (File.Exists(target) || Directory.Exists(target)))
            {
                return target;
            }
        }
        catch {}
        return null;
    }

    private static string FindPathInCommonFolders(string fileName)
    {
        Environment.SpecialFolder[] folders = new Environment.SpecialFolder[]
        {
            Environment.SpecialFolder.DesktopDirectory,
            Environment.SpecialFolder.MyDocuments,
            Environment.SpecialFolder.UserProfile
        };

        foreach (var folder in folders)
        {
            try
            {
                string basePath = Environment.GetFolderPath(folder);
                if (!Directory.Exists(basePath)) continue;

                if (folder == Environment.SpecialFolder.UserProfile)
                {
                    string downloads = Path.Combine(basePath, "Downloads");
                    if (Directory.Exists(downloads))
                    {
                        string target = SearchFileInDir(downloads, fileName);
                        if (target != null) return target;
                    }
                }
                else
                {
                    string target = SearchFileInDir(basePath, fileName);
                    if (target != null) return target;
                }
            }
            catch {}
        }
        return null;
    }

    private static string SearchFileInDir(string dir, string fileName)
    {
        try
        {
            string[] matches = Directory.GetFiles(dir, fileName, SearchOption.TopDirectoryOnly);
            if (matches.Length > 0) return matches[0];

            string[] subdirs = Directory.GetDirectories(dir);
            foreach (string subdir in subdirs)
            {
                try
                {
                    DirectoryInfo di = new DirectoryInfo(subdir);
                    if ((di.Attributes & FileAttributes.Hidden) != 0) continue;

                    string[] subMatches = Directory.GetFiles(subdir, fileName, SearchOption.TopDirectoryOnly);
                    if (subMatches.Length > 0) return subMatches[0];
                }
                catch {}
            }
        }
        catch {}
        return null;
    }
}
