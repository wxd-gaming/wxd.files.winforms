using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace wxd.files.winforms
{
    /// <summary>
    /// 笔记持久化管理（单例 + 后台写盘）
    /// 存储到程序所在目录下的 notes.json，方便随程序迁移
    /// </summary>
    public static class NoteStore
    {
        private static readonly object _lock = new();
        private static List<NoteItem> _notes = new();
        private static bool _loaded;

        /// <summary>笔记文件路径：程序目录\notes.json</summary>
        private static string SavePath => Path.Combine(AppContext.BaseDirectory, "notes.json");

        /// <summary>确保数据已从磁盘加载</summary>
        private static void EnsureLoaded()
        {
            if (_loaded) return;
            lock (_lock)
            {
                if (_loaded) return;
                try
                {
                    if (File.Exists(SavePath))
                    {
                        var json = File.ReadAllText(SavePath);
                        _notes = JsonSerializer.Deserialize<List<NoteItem>>(json) ?? new List<NoteItem>();
                    }
                }
                catch
                {
                    _notes = new List<NoteItem>();
                }
                _loaded = true;
            }
        }

        /// <summary>获取全部笔记（最新在前，返回副本）</summary>
        public static List<NoteItem> GetAll()
        {
            EnsureLoaded();
            lock (_lock) return new List<NoteItem>(_notes);
        }

        /// <summary>根据 Id 获取笔记（不存在返回 null）</summary>
        public static NoteItem? GetById(string id)
        {
            EnsureLoaded();
            lock (_lock) return _notes.Find(n => n.Id == id);
        }

        /// <summary>新增笔记并返回，写入磁盘在后台执行</summary>
        public static NoteItem Add()
        {
            EnsureLoaded();
            var note = new NoteItem
            {
                Title = "新建笔记",
                RtfContent = "",
                PlainText = "",
                CreatedAt = DateTime.Now,
                UpdatedAt = DateTime.Now
            };
            lock (_lock) _notes.Insert(0, note);
            _ = Task.Run(Save);
            return note;
        }

        /// <summary>更新指定笔记（标题/RTF/纯文本），并刷新 UpdatedAt</summary>
        public static void Update(string id, string title, string rtfContent, string plainText)
        {
            EnsureLoaded();
            lock (_lock)
            {
                var n = _notes.Find(x => x.Id == id);
                if (n == null) return;
                n.Title = title;
                n.RtfContent = rtfContent;
                n.PlainText = plainText;
                n.UpdatedAt = DateTime.Now;
                // 将更新的笔记移到列表最前
                _notes.Remove(n);
                _notes.Insert(0, n);
            }
            _ = Task.Run(Save);
        }

        /// <summary>删除指定笔记</summary>
        public static void Delete(string id)
        {
            EnsureLoaded();
            lock (_lock) _notes.RemoveAll(n => n.Id == id);
            _ = Task.Run(Save);
        }

        /// <summary>持久化到磁盘</summary>
        private static void Save()
        {
            try
            {
                List<NoteItem> snap;
                lock (_lock) snap = new List<NoteItem>(_notes);
                var json = JsonSerializer.Serialize(snap, new JsonSerializerOptions { WriteIndented = false });
                AtomicFileWriter.TryWriteAllText(SavePath, json);
            }
            catch
            {
                // 忽略写入失败
            }
        }
    }
}
