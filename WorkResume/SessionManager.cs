using System;
using System.IO;
using System.Collections.Generic;
using System.Web.Script.Serialization;

public class SessionItem
{
    public string Type { get; set; } // "File" or "Folder"
    public string FilePath { get; set; }
    public string FileName { get; set; }
    public string Extension { get; set; }
    public string PreferredApplication { get; set; }
    public bool Exists { get; set; }
    public bool IsSelected { get; set; }
}

public class WorkSession
{
    public string Id { get; set; }
    public string Title { get; set; }
    public string Description { get; set; }
    public string SavedAt { get; set; }
    public List<SessionItem> Items { get; set; }

    public WorkSession()
    {
        Items = new List<SessionItem>();
    }
}

public static class SessionManager
{
    private static readonly string SessionDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        @"WorkResume"
    );

    private static readonly string SessionFilePath = Path.Combine(SessionDirectory, "sessions.json");

    static SessionManager()
    {
        try
        {
            if (!Directory.Exists(SessionDirectory))
            {
                Directory.CreateDirectory(SessionDirectory);
            }
        }
        catch {}
    }

    public static List<WorkSession> LoadSessions()
    {
        try
        {
            if (!File.Exists(SessionFilePath))
            {
                Logger.Write("저장된 세션 파일이 없습니다.");
                return new List<WorkSession>();
            }

            string json = File.ReadAllText(SessionFilePath);
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            
            // JSON 문자열을 객체 리스트로 역직렬화
            List<WorkSession> sessions = serializer.Deserialize<List<WorkSession>>(json);
            if (sessions == null) sessions = new List<WorkSession>();

            // 세션 로드 시 각 파일들의 최신 존재 여부를 체크하여 실시간 업데이트
            foreach (var session in sessions)
            {
                foreach (var item in session.Items)
                {
                    if (item.Type == "File")
                    {
                        item.Exists = File.Exists(item.FilePath);
                    }
                    else if (item.Type == "Folder")
                    {
                        item.Exists = Directory.Exists(item.FilePath);
                    }
                }
            }

            Logger.Write(string.Format("세션 로드 완료: 총 {0}개 세션", sessions.Count));
            return sessions;
        }
        catch (Exception ex)
        {
            Logger.Write("세션 로드 중 에러 발생: " + ex.Message, "ERROR");
            return new List<WorkSession>();
        }
    }

    public static bool SaveSessions(List<WorkSession> sessions)
    {
        try
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(sessions);
            
            File.WriteAllText(SessionFilePath, json);
            Logger.Write("세션 저장 완료");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Write("세션 저장 중 에러 발생: " + ex.Message, "ERROR");
            return false;
        }
    }

    public static string CreateSession(string title, string desc, List<DetectedItem> items)
    {
        try
        {
            List<WorkSession> sessions = LoadSessions();
            WorkSession newSession = new WorkSession
            {
                Id = Guid.NewGuid().ToString(),
                Title = string.IsNullOrEmpty(title) ? DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " 업무" : title,
                Description = desc,
                SavedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };

            foreach (var detected in items)
            {
                newSession.Items.Add(new SessionItem
                {
                    Type = detected.Type,
                    FilePath = detected.FilePath,
                    FileName = detected.FileName,
                    Extension = detected.Extension,
                    PreferredApplication = detected.PreferredApplication,
                    Exists = detected.Exists,
                    IsSelected = detected.IsSelected
                });
            }

            // 최신 세션이 목록 맨 위에 오도록 추가
            sessions.Insert(0, newSession);
            
            if (SaveSessions(sessions))
            {
                Logger.Write("새로운 세션 생성: " + newSession.Title);
                return newSession.Id;
            }
        }
        catch (Exception ex)
        {
            Logger.Write("세션 생성 중 에러 발생: " + ex.Message, "ERROR");
        }
        return null;
    }

    public static bool DeleteSession(string id)
    {
        try
        {
            List<WorkSession> sessions = LoadSessions();
            WorkSession target = sessions.Find(s => s.Id == id);
            if (target != null)
            {
                sessions.Remove(target);
                bool success = SaveSessions(sessions);
                if (success)
                {
                    Logger.Write("세션 삭제 완료: " + target.Title);
                }
                return success;
            }
        }
        catch (Exception ex)
        {
            Logger.Write("세션 삭제 중 에러 발생: " + ex.Message, "ERROR");
        }
        return false;
    }

    public static bool UpdateSession(string id, string title, string desc, List<DetectedItem> items)
    {
        try
        {
            List<WorkSession> sessions = LoadSessions();
            int idx = sessions.FindIndex(s => s.Id == id);
            if (idx >= 0)
            {
                WorkSession updated = new WorkSession
                {
                    Id = id,
                    Title = string.IsNullOrEmpty(title) ? DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " 업무" : title,
                    Description = desc,
                    SavedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                };

                foreach (var detected in items)
                {
                    updated.Items.Add(new SessionItem
                    {
                        Type = detected.Type,
                        FilePath = detected.FilePath,
                        FileName = detected.FileName,
                        Extension = detected.Extension,
                        PreferredApplication = detected.PreferredApplication,
                        Exists = detected.Exists,
                        IsSelected = detected.IsSelected
                    });
                }

                sessions[idx] = updated;
                // 업데이트 후 최신순으로 맨 위로 보낼 수도 있으나, 위치를 그대로 유지하거나 최신 순서로 재배열 가능.
                // 여기서는 사용성상 최신 정렬을 유지하기 위해 맨 앞으로 끌어올림
                sessions.RemoveAt(idx);
                sessions.Insert(0, updated);

                bool success = SaveSessions(sessions);
                if (success)
                {
                    Logger.Write("세션 업데이트 완료: " + updated.Title);
                }
                return success;
            }
        }
        catch (Exception ex)
        {
            Logger.Write("세션 업데이트 중 에러 발생: " + ex.Message, "ERROR");
        }
        return false;
    }
}
