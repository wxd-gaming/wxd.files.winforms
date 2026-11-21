using System;
using System.Windows.Forms;

namespace wxd.files.winforms
{
    /// <summary>YAML 页签：在同一份文本上做 YAML→JSON、JSON→YAML 和语法校验。</summary>
    internal sealed class YamlToolPanel : ToolPanelBase
    {
        private readonly TextBox _editor = MakeEditor();
        private readonly ThemeBordered _border;

        internal YamlToolPanel()
        {
            _border = new ThemeBordered(_editor) { Dock = DockStyle.Fill };

            AddButton("YAML → JSON", (_, _) => ToJson(), primary: true);
            AddButton("JSON → YAML", (_, _) => ToYaml());
            AddButton("校验 YAML", (_, _) => Validate());
            AddButton("复制全部", (_, _) => CopyToClipboard(_editor.Text));
            AddButton("清空", (_, _) =>
            {
                _editor.Clear();
                SetStatus("已清空");
            });

            FinishChrome(_border);
            SetStatus("粘贴 YAML 或 JSON，再点转换；普通标量会按 null / 布尔 / 数字 / 字符串推断");
        }

        protected override void ApplyThemeCore()
        {
            _border.ApplyTheme();
            ApplyEditorTheme(_editor);
        }

        private void ToJson()
        {
            var result = YamlToolEngine.YamlToJson(_editor.Text);
            Replace(result, "已转为 JSON");
        }

        private void ToYaml()
        {
            if (JsonToolEngine.ContainsComments(_editor.Text)
                && AlertForm.ShowConfirm(this,
                    "当前 JSON 包含注释，转成 YAML 会移除注释，是否继续？",
                    "工具箱", AlertType.Warning) != DialogResult.Yes)
                return;

            var result = YamlToolEngine.JsonToYaml(_editor.Text);
            Replace(result, "已转为 YAML");
        }

        private void Validate()
        {
            var result = YamlToolEngine.Validate(_editor.Text);
            SetStatus(result.Success ? result.Text : result.Error ?? "YAML 校验失败", error: !result.Success);
        }

        private void Replace(TextToolResult result, string action)
        {
            if (!result.Success)
            {
                SetStatus(result.Error ?? "转换失败", error: true);
                return;
            }

            _editor.Text = result.Text;
            _editor.SelectionStart = 0;
            _editor.SelectionLength = 0;
            SetStatus($"{action} · {result.Text.Length} 字符");
        }
    }
}
