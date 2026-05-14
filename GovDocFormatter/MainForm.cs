using System.Diagnostics;

namespace GovDocFormatter;

public sealed class MainForm : Form
{
    private const string AuthorName = "阿谪talk";
    private const string WechatQrCodeResourceName = "GovDocFormatter.Assets.wechat-qr.png";

    private readonly TextBox _inputPathTextBox = SingleLineTextBox();
    private readonly TextBox _outputPathTextBox = SingleLineTextBox();

    private readonly TextBox _serialNumberTextBox = SingleLineTextBox();
    private readonly TextBox _secretLevelTextBox = SingleLineTextBox();
    private readonly TextBox _urgencyTextBox = SingleLineTextBox();
    private readonly TextBox _dispatchMarkTextBox = SingleLineTextBox();
    private readonly TextBox _dispatchNumberTextBox = SingleLineTextBox();
    private readonly TextBox _signerNameTextBox = SingleLineTextBox();
    private readonly TextBox _titleTextBox = SingleLineTextBox();
    private readonly TextBox _mainRecipientTextBox = SingleLineTextBox();
    private readonly TextBox _issuerTextBox = SingleLineTextBox();
    private readonly TextBox _documentDateTextBox = SingleLineTextBox();
    private readonly TextBox _copyToTextBox = SingleLineTextBox();
    private readonly TextBox _printingOfficeTextBox = SingleLineTextBox();
    private readonly TextBox _printingDateTextBox = SingleLineTextBox();

    private readonly CheckBox _insertFrontMatterCheckBox = new()
    {
        Text = "按填写内容插入版头、标题、主送机关",
        Checked = true,
        AutoSize = true
    };

    private readonly CheckBox _insertEndingMatterCheckBox = new()
    {
        Text = "按填写内容插入落款、成文日期、版记",
        Checked = true,
        AutoSize = true
    };

    private readonly CheckBox _applyPageNumbersCheckBox = new()
    {
        Text = "添加单双页外侧页码",
        Checked = true,
        AutoSize = true
    };

    private readonly CheckBox _formatExistingContentCheckBox = new()
    {
        Text = "格式化现有正文和层级标题",
        Checked = true,
        AutoSize = true
    };

    private readonly CheckBox _autoDetectTitleCheckBox = new()
    {
        Text = "未填写标题时，将首个短段落识别为标题",
        Checked = true,
        AutoSize = true
    };

    private readonly CheckBox _upwardDocumentCheckBox = new()
    {
        Text = "上行文：发文字号左侧、签发人右侧",
        AutoSize = true
    };

    private readonly Button _formatButton = new()
    {
        Text = "开始排版",
        AutoSize = true,
        Height = 36
    };

    private readonly Button _openOutputFolderButton = new()
    {
        Text = "打开输出目录",
        AutoSize = true,
        Height = 36,
        Enabled = false
    };

    private readonly TextBox _logTextBox = new()
    {
        Dock = DockStyle.Fill,
        Multiline = true,
        ReadOnly = true,
        ScrollBars = ScrollBars.Vertical,
        BackColor = Color.White,
        BorderStyle = BorderStyle.FixedSingle
    };

    public MainForm()
    {
        Text = "公文 Word 自动排版";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(980, 720);
        Size = new Size(1080, 780);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Microsoft YaHei UI", 9F);

        BuildLayout();
        WireEvents();
        AppendLog("请选择 .docx 文件。程序会复制原文件并生成一个新的排版结果，不会覆盖原稿。");
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(16),
            BackColor = Color.White
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));

        root.Controls.Add(BuildFileGroup(), 0, 0);
        root.Controls.Add(BuildTabs(), 0, 1);
        root.Controls.Add(BuildActionBar(), 0, 2);
        root.Controls.Add(_logTextBox, 0, 3);

        Controls.Add(root);
    }

    private Control BuildFileGroup()
    {
        var group = new GroupBox
        {
            Text = "文件",
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(12)
        };

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 3,
            RowCount = 2
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var inputButton = new Button { Text = "选择...", AutoSize = true };
        inputButton.Click += (_, _) => SelectInputFile();

        var outputButton = new Button { Text = "另存为...", AutoSize = true };
        outputButton.Click += (_, _) => SelectOutputFile();

        grid.Controls.Add(new Label { Text = "源文件", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        grid.Controls.Add(_inputPathTextBox, 1, 0);
        grid.Controls.Add(inputButton, 2, 0);
        grid.Controls.Add(new Label { Text = "输出到", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        grid.Controls.Add(_outputPathTextBox, 1, 1);
        grid.Controls.Add(outputButton, 2, 1);

        group.Controls.Add(grid);
        return group;
    }

    private Control BuildTabs()
    {
        var tabs = new TabControl
        {
            Dock = DockStyle.Fill,
            Padding = new Point(12, 6)
        };

        var metadataTab = new TabPage("公文要素")
        {
            BackColor = Color.White,
            Padding = new Padding(8)
        };
        metadataTab.Controls.Add(BuildMetadataGrid());

        var optionsTab = new TabPage("排版选项")
        {
            BackColor = Color.White,
            Padding = new Padding(16)
        };
        optionsTab.Controls.Add(BuildOptionsPanel());

        tabs.TabPages.Add(metadataTab);
        tabs.TabPages.Add(optionsTab);
        return tabs;
    }

    private Control BuildMetadataGrid()
    {
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            ColumnCount = 4,
            RowCount = 7,
            Padding = new Padding(4)
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        AddFieldRow(grid, 0, "份号", _serialNumberTextBox, "密级", _secretLevelTextBox);
        AddFieldRow(grid, 1, "紧急程度", _urgencyTextBox, "发文机关", _dispatchMarkTextBox);
        AddFieldRow(grid, 2, "发文字号", _dispatchNumberTextBox, "签发人", _signerNameTextBox);
        AddFieldRow(grid, 3, "标题", _titleTextBox, "主送机关", _mainRecipientTextBox);
        AddFieldRow(grid, 4, "发文机关署名", _issuerTextBox, "成文日期", _documentDateTextBox);
        AddFieldRow(grid, 5, "抄送机关", _copyToTextBox, "印发机关", _printingOfficeTextBox);
        AddFieldRow(grid, 6, "印发日期", _printingDateTextBox, "", new Label());

        return grid;
    }

    private Control BuildOptionsPanel()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));

        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true
        };

        panel.Controls.Add(_insertFrontMatterCheckBox);
        panel.Controls.Add(_insertEndingMatterCheckBox);
        panel.Controls.Add(_applyPageNumbersCheckBox);
        panel.Controls.Add(_formatExistingContentCheckBox);
        panel.Controls.Add(_autoDetectTitleCheckBox);
        panel.Controls.Add(_upwardDocumentCheckBox);

        layout.Controls.Add(panel, 0, 0);
        layout.Controls.Add(BuildAuthorPanel(), 1, 0);
        return layout;
    }

    private static Control BuildAuthorPanel()
    {
        var group = new GroupBox
        {
            Text = "作者与联系",
            Dock = DockStyle.Fill,
            Padding = new Padding(12)
        };

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4
        };
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        panel.Controls.Add(new Label
        {
            Text = "作者： " + AuthorName,
            AutoSize = true,
            Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 10)
        }, 0, 0);

        panel.Controls.Add(new Label
        {
            Text = "微信二维码",
            AutoSize = true,
            ForeColor = Color.DimGray,
            Margin = new Padding(0, 0, 0, 8)
        }, 0, 1);

        var qrCodePictureBox = new PictureBox
        {
            Dock = DockStyle.Fill,
            Image = LoadWechatQrCodeImage(),
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            MinimumSize = new Size(180, 220)
        };
        panel.Controls.Add(qrCodePictureBox, 0, 2);

        panel.Controls.Add(new Label
        {
            Text = "扫码添加微信",
            AutoSize = true,
            ForeColor = Color.Gray,
            TextAlign = ContentAlignment.MiddleCenter,
            Anchor = AnchorStyles.None,
            Margin = new Padding(0, 8, 0, 0)
        }, 0, 3);

        group.Controls.Add(panel);
        return group;
    }

    private static Image? LoadWechatQrCodeImage()
    {
        using var stream = typeof(MainForm).Assembly.GetManifestResourceStream(WechatQrCodeResourceName);
        if (stream is null)
        {
            return null;
        }

        using var image = Image.FromStream(stream);
        return new Bitmap(image);
    }

    private Control BuildActionBar()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Padding = new Padding(0, 12, 0, 12)
        };
        panel.Controls.Add(_formatButton);
        panel.Controls.Add(_openOutputFolderButton);
        return panel;
    }

    private static void AddFieldRow(
        TableLayoutPanel grid,
        int row,
        string leftLabel,
        Control leftControl,
        string rightLabel,
        Control rightControl)
    {
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));

        grid.Controls.Add(new Label { Text = leftLabel, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);
        grid.Controls.Add(leftControl, 1, row);

        if (!string.IsNullOrWhiteSpace(rightLabel))
        {
            grid.Controls.Add(new Label { Text = rightLabel, AutoSize = true, Anchor = AnchorStyles.Left }, 2, row);
            grid.Controls.Add(rightControl, 3, row);
        }
    }

    private void WireEvents()
    {
        _formatButton.Click += async (_, _) => await FormatDocumentAsync();
        _openOutputFolderButton.Click += (_, _) => OpenOutputFolder();
    }

    private void SelectInputFile()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Word 文档 (*.docx)|*.docx",
            Title = "选择 Word 文档",
            CheckFileExists = true,
            Multiselect = false
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        _inputPathTextBox.Text = dialog.FileName;
        _outputPathTextBox.Text = BuildDefaultOutputPath(dialog.FileName);
        _openOutputFolderButton.Enabled = false;
        AppendLog("已选择源文件：" + dialog.FileName);
    }

    private void SelectOutputFile()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Word 文档 (*.docx)|*.docx",
            Title = "保存排版结果",
            AddExtension = true,
            DefaultExt = "docx",
            OverwritePrompt = true
        };

        if (!string.IsNullOrWhiteSpace(_outputPathTextBox.Text))
        {
            dialog.FileName = _outputPathTextBox.Text;
        }
        else if (!string.IsNullOrWhiteSpace(_inputPathTextBox.Text))
        {
            dialog.FileName = BuildDefaultOutputPath(_inputPathTextBox.Text);
        }

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _outputPathTextBox.Text = dialog.FileName;
            _openOutputFolderButton.Enabled = false;
        }
    }

    private async Task FormatDocumentAsync()
    {
        var inputPath = _inputPathTextBox.Text.Trim();
        var outputPath = _outputPathTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
        {
            MessageBox.Show(this, "请先选择存在的 .docx 源文件。", "缺少源文件", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            outputPath = BuildDefaultOutputPath(inputPath);
            _outputPathTextBox.Text = outputPath;
        }

        SetBusy(true);
        AppendLog("开始排版：" + Path.GetFileName(inputPath));

        try
        {
            var options = CollectOptions();
            var formatter = new GovDocumentFormatter();
            var result = await Task.Run(() => formatter.Format(inputPath, outputPath, options));

            AppendLog("完成：" + result.OutputPath);
            foreach (var message in result.Messages)
            {
                AppendLog(" - " + message);
            }

            _openOutputFolderButton.Enabled = true;
            MessageBox.Show(this, "排版完成。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            AppendLog("错误：" + ex.Message);
            MessageBox.Show(this, ex.Message, "排版失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private FormattingOptions CollectOptions()
    {
        return new FormattingOptions
        {
            SerialNumber = _serialNumberTextBox.Text.Trim(),
            SecretLevel = _secretLevelTextBox.Text.Trim(),
            Urgency = _urgencyTextBox.Text.Trim(),
            DispatchMark = _dispatchMarkTextBox.Text.Trim(),
            DispatchNumber = _dispatchNumberTextBox.Text.Trim(),
            SignerName = _signerNameTextBox.Text.Trim(),
            Title = _titleTextBox.Text.Trim(),
            MainRecipient = _mainRecipientTextBox.Text.Trim(),
            Issuer = _issuerTextBox.Text.Trim(),
            DocumentDate = _documentDateTextBox.Text.Trim(),
            CopyTo = _copyToTextBox.Text.Trim(),
            PrintingOffice = _printingOfficeTextBox.Text.Trim(),
            PrintingDate = _printingDateTextBox.Text.Trim(),
            InsertFrontMatter = _insertFrontMatterCheckBox.Checked,
            InsertEndingMatter = _insertEndingMatterCheckBox.Checked,
            ApplyPageNumbers = _applyPageNumbersCheckBox.Checked,
            FormatExistingContent = _formatExistingContentCheckBox.Checked,
            AutoDetectTitle = _autoDetectTitleCheckBox.Checked,
            UpwardDocument = _upwardDocumentCheckBox.Checked
        };
    }

    private void SetBusy(bool busy)
    {
        _formatButton.Enabled = !busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
    }

    private void OpenOutputFolder()
    {
        var outputPath = _outputPathTextBox.Text.Trim();
        var directory = Path.GetDirectoryName(outputPath);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = "/select,\"" + outputPath + "\"",
            UseShellExecute = true
        });
    }

    private void AppendLog(string message)
    {
        _logTextBox.AppendText(DateTime.Now.ToString("HH:mm:ss ") + message + Environment.NewLine);
    }

    private static TextBox SingleLineTextBox()
    {
        return new TextBox
        {
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            Width = 320
        };
    }

    private static string BuildDefaultOutputPath(string inputPath)
    {
        var directory = Path.GetDirectoryName(inputPath) ?? "";
        var fileName = Path.GetFileNameWithoutExtension(inputPath);
        return Path.Combine(directory, fileName + "_公文排版.docx");
    }
}
