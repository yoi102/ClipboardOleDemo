using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.Json;
using Vanara.PInvoke;
using static Vanara.PInvoke.Ole32;

namespace ClipboardOleDemo;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        OleInitialize(IntPtr.Zero).ThrowIfFailed();

        try
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        finally
        {
            OleUninitialize();
        }
    }
}

public sealed class MainForm : Form
{
    private readonly Button _btnFormats;
    private readonly Button _btnRaw;
    private readonly Button _btnDrawable;
    private readonly Button _btnCreateFromFile;
    private readonly Button _btnSave;
    private readonly Button _btnLoad;
    private readonly TextBox _txtInfo;
    private readonly Panel _panelPreview;

    private ClipboardDrawableOleObject? _drawable;

    public MainForm()
    {
        Text = "Vanara Clipboard/File OLE Persist Demo";
        Width = 1280;
        Height = 820;
        StartPosition = FormStartPosition.CenterScreen;

        _btnFormats = new Button
        {
            Text = "枚举格式",
            Left = 12,
            Top = 12,
            Width = 110,
            Height = 25
        };
        _btnFormats.Click += (_, __) => DumpFormats();

        _btnRaw = new Button
        {
            Text = "读取 Descriptor",
            Left = 132,
            Top = 12,
            Width = 130,
            Height = 25
        };
        _btnRaw.Click += (_, __) => ReadObjectDescriptor();

        _btnDrawable = new Button
        {
            Text = "从剪切板创建",
            Left = 272,
            Top = 12,
            Width = 140,
            Height = 25
        };
        _btnDrawable.Click += (_, __) => CreateDrawableFromClipboard();

        _btnCreateFromFile = new Button
        {
            Text = "从本地文件创建",
            Left = 422,
            Top = 12,
            Width = 150,
            Height = 25
        };
        _btnCreateFromFile.Click += (_, __) => CreateDrawableFromLocalFile();

        _btnSave = new Button
        {
            Text = "保存到文件",
            Left = 12,
            Top = 40,
            Width = 120,
            Height = 25
        };
        _btnSave.Click += (_, __) => SaveToFile();

        _btnLoad = new Button
        {
            Text = "从文件打开",
            Left = 142,
            Top = 40,
            Width = 120,
            Height = 25
        };
        _btnLoad.Click += (_, __) => LoadFromFile();

        _txtInfo = new TextBox
        {
            Left = 12,
            Top = 80,
            Width = 560,
            Height = 690,
            Multiline = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Font = new Font("Consolas", 10)
        };

        _panelPreview = new Panel
        {
            Left = 585,
            Top = 12,
            Width = 670,
            Height = 758,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White
        };
        _panelPreview.Paint += PanelPreview_Paint;
        _panelPreview.MouseDoubleClick += PanelPreview_DoubleClick;

        Controls.Add(_btnFormats);
        Controls.Add(_btnRaw);
        Controls.Add(_btnDrawable);
        Controls.Add(_btnCreateFromFile);
        Controls.Add(_btnSave);
        Controls.Add(_btnLoad);
        Controls.Add(_txtInfo);
        Controls.Add(_panelPreview);
        this.Resize += MainForm_Resize;
    }

    private void MainForm_Resize(object? sender, EventArgs e)
    {
        _panelPreview.Height = ClientSize.Height - _panelPreview.Top - 12;
        _panelPreview.Width = ClientSize.Width - _panelPreview.Left - 12;
        _panelPreview.Invalidate();


    }
    private void DumpFormats()
    {
        try
        {
            List<ClipboardFormatInfo> formats = VanaraClipboardOle.GetClipboardFormats();
            var sb = new StringBuilder();
            sb.AppendLine("Clipboard formats");
            sb.AppendLine(new string('-', 100));

            foreach (ClipboardFormatInfo f in formats)
            {
                sb.AppendLine($"cf={f.FormatId,-6} name={f.FormatName,-30} tymed={f.Tymed,-20} aspect={f.Aspect}");
            }

            _txtInfo.Text = sb.ToString();
        }
        catch (Exception ex)
        {
            _txtInfo.Text = ex.ToString();
        }
    }

    private void ReadObjectDescriptor()
    {
        try
        {
            if (!VanaraClipboardOle.TryGetObjectDescriptorFromClipboard(out ClipboardOleDescriptorInfo? desc))
            {
                _txtInfo.Text = "剪切板中没有 Object Descriptor，无法识别 OLE 类型。";
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("Object Descriptor");
            sb.AppendLine(new string('-', 100));
            sb.AppendLine($"FormatId         : {desc.FormatId}");
            sb.AppendLine($"FormatName       : {desc.FormatName}");
            sb.AppendLine($"Tymed            : {desc.Tymed}");
            sb.AppendLine($"OleKind          : {desc.OleKind}");
            sb.AppendLine($"ClassId          : {desc.ClassId}");
            sb.AppendLine($"ProgId           : {desc.ProgId}");
            sb.AppendLine($"FullUserTypeName : {desc.FullUserTypeName}");
            sb.AppendLine($"SourceOfCopy     : {desc.SourceOfCopy}");
            sb.AppendLine($"DrawAspect       : {desc.DrawAspect}");
            sb.AppendLine($"Status           : 0x{desc.Status:X8}");

            _txtInfo.Text = sb.ToString();
        }
        catch (Exception ex)
        {
            _txtInfo.Text = ex.ToString();
        }
    }

    private void CreateDrawableFromClipboard()
    {
        try
        {
            ClipboardOleDescriptorInfo? desc = null;
            VanaraClipboardOle.TryGetObjectDescriptorFromClipboard(out desc);

            _drawable?.Dispose();
            _drawable = VanaraClipboardOle.CreateDrawableObjectFromClipboard(_panelPreview);

            var sb = new StringBuilder();

            if (_drawable is null)
            {
                sb.AppendLine("从剪切板创建失败");
                if (desc != null)
                {
                    sb.AppendLine(new string('-', 100));
                    sb.AppendLine($"类型     : {desc.OleKind}");
                    sb.AppendLine($"ProgId   : {desc.ProgId}");
                    sb.AppendLine($"UserType : {desc.FullUserTypeName}");
                    sb.AppendLine($"Source   : {desc.SourceOfCopy}");
                }

                _txtInfo.Text = sb.ToString();
                return;
            }

            sb.AppendLine("从剪切板创建成功");
            sb.AppendLine(new string('-', 100));

            if (desc != null)
            {
                sb.AppendLine($"类型     : {desc.OleKind}");
                sb.AppendLine($"ProgId   : {desc.ProgId}");
                sb.AppendLine($"UserType : {desc.FullUserTypeName}");
                sb.AppendLine($"Source   : {desc.SourceOfCopy}");
                sb.AppendLine($"CLSID    : {desc.ClassId}");
                sb.AppendLine();
            }

            PersistedOlePackage package = _drawable.ExportPackage();
            sb.AppendLine($"StorageBytes    : {package.StorageBytes.Length}");
            sb.AppendLine($"Version         : {package.Version}");
            sb.AppendLine($"DisplayName     : {package.DisplayName}");
            sb.AppendLine($"SavedAtUtc      : {package.SavedAtUtc:O}");
            sb.AppendLine();
            sb.AppendLine("双击右侧预览区打开原生编辑器。关闭编辑器后，如果 OLE Server 调用了 SaveObject，修改会保存回当前 OLE 存储。再次双击应打开最新内容。保存到文件后，可在后续会话中重新恢复。");

            _txtInfo.Text = sb.ToString();
            _panelPreview.Invalidate();
        }
        catch (Exception ex)
        {
            _txtInfo.Text = ex.ToString();
        }
    }

    private void CreateDrawableFromLocalFile()
    {
        try
        {
            using var dialog = new OpenFileDialog
            {
                Title = "选择本地文件并创建 OLE",
                Filter = "All files (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            _drawable?.Dispose();
            _drawable = VanaraClipboardOle.CreateDrawableObjectFromFile(dialog.FileName, _panelPreview);

            var sb = new StringBuilder();
            sb.AppendLine("从本地文件创建成功");
            sb.AppendLine(new string('-', 100));
            sb.AppendLine($"FilePath        : {dialog.FileName}");
            sb.AppendLine($"FileName        : {Path.GetFileName(dialog.FileName)}");
            sb.AppendLine();

            PersistedOlePackage package = _drawable.ExportPackage(Path.GetFileName(dialog.FileName));
            sb.AppendLine($"StorageBytes    : {package.StorageBytes.Length}");
            sb.AppendLine($"Version         : {package.Version}");
            sb.AppendLine($"DisplayName     : {package.DisplayName}");
            sb.AppendLine($"SavedAtUtc      : {package.SavedAtUtc:O}");
            sb.AppendLine();
            sb.AppendLine("双击右侧预览区打开原生编辑器。关闭编辑器后，如果 OLE Server 调用了 SaveObject，修改会保存回当前 OLE 存储。");
            sb.AppendLine("注意：只有已注册 OLE Server 的文件类型（如 Excel/Word/PDF/Visio 等）才能正常创建；普通未知类型文件可能会失败。");

            _txtInfo.Text = sb.ToString();
            _panelPreview.Invalidate();
        }
        catch (Exception ex)
        {
            _txtInfo.Text = ex.ToString();
        }
    }

    private void SaveToFile()
    {
        try
        {
            if (_drawable == null)
            {
                MessageBox.Show(this, "当前还没有对象可保存。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Title = "保存 OLE 对象",
                Filter = "OLE Persist Package (*.olepkg.json)|*.olepkg.json|All files (*.*)|*.*",
                FileName = "clipboard-object.olepkg.json"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            PersistedOlePackage package = _drawable.ExportPackage();
            string json = JsonSerializer.Serialize(package);
            File.WriteAllText(dialog.FileName, json, Encoding.UTF8);

            ShowDrawableInfo($"已保存到文件: {dialog.FileName}");
        }
        catch (Exception ex)
        {
            _txtInfo.Text = ex.ToString();
        }
    }

    private void LoadFromFile()
    {
        try
        {
            using var dialog = new OpenFileDialog
            {
                Title = "打开 OLE 对象",
                Filter = "OLE Persist Package (*.olepkg.json)|*.olepkg.json|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            string json = File.ReadAllText(dialog.FileName, Encoding.UTF8);
            PersistedOlePackage? package = JsonSerializer.Deserialize<PersistedOlePackage>(json);
            if (package == null)
                throw new InvalidOperationException("文件反序列化失败。");

            _drawable?.Dispose();
            _drawable = VanaraClipboardOle.OpenFromPersistedPackage(package, _panelPreview);
            ShowDrawableInfo($"已从文件恢复: {dialog.FileName}", Path.GetFileName(dialog.FileName));
        }
        catch (Exception ex)
        {
            _txtInfo.Text = ex.ToString();
        }
    }

    private void ShowDrawableInfo(string title, string? displayName = null)
    {
        if (_drawable == null)
            return;

        PersistedOlePackage package = _drawable.ExportPackage(displayName ?? "Clipboard OLE Object");

        var sb = new StringBuilder();
        sb.AppendLine(title);
        sb.AppendLine(new string('-', 100));
        sb.AppendLine($"StorageBytes    : {package.StorageBytes.Length}");
        sb.AppendLine($"Version         : {package.Version}");
        sb.AppendLine($"DisplayName     : {package.DisplayName}");
        sb.AppendLine($"SavedAtUtc      : {package.SavedAtUtc:O}");
        sb.AppendLine();
        sb.AppendLine("双击右侧预览区打开原生编辑器。关闭编辑器后，如果 OLE Server 调用了 SaveObject，修改会保存回当前 OLE 存储。再次双击应打开最新内容。保存到文件后，可在后续会话中重新恢复。");

        _txtInfo.Text = sb.ToString();
        _panelPreview.Invalidate();
    }

    private void PanelPreview_DoubleClick(object? sender, MouseEventArgs e)
    {
        if (_drawable == null)
            return;

        try
        {
            _drawable.OpenEditor(
                Handle,
                containerAppName: "WinFormsApp4",
                documentName: "OLE Object");
        }
        catch (Exception ex)
        {
            _txtInfo.Text = ex.ToString();
        }
    }

    private void PanelPreview_Paint(object? sender, PaintEventArgs e)
    {
        e.Graphics.Clear(Color.White);

        if (_drawable == null)
        {
            using var pen = new Pen(Color.Silver);
            e.Graphics.DrawRectangle(
                pen,
                10,
                10,
                _panelPreview.ClientSize.Width - 20,
                _panelPreview.ClientSize.Height - 20);

            using var brush = new SolidBrush(Color.Gray);
            e.Graphics.DrawString("这里显示 OleDraw 结果", Font, brush, 20, 20);
            return;
        }

        var outerBox = new Rectangle(
            20,
            20,
            _panelPreview.ClientSize.Width - 40,
            _panelPreview.ClientSize.Height - 40);

        Rectangle drawRect = outerBox;
        Size? naturalSize = _drawable.GetNaturalPixelSize(e.Graphics.DpiX, e.Graphics.DpiY);
        if (naturalSize.HasValue)
        {
            drawRect.Width = naturalSize.Value.Width;
            drawRect.Height = naturalSize.Value.Height;
        }
        else
        {
            return;
        }

        _drawable.Draw(e.Graphics, drawRect);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _drawable?.Dispose();
        base.OnFormClosed(e);
    }
}

public sealed class ClipboardFormatInfo
{
    public ushort FormatId { get; init; }
    public string FormatName { get; init; } = "";
    public TYMED Tymed { get; init; }
    public DVASPECT Aspect { get; init; }
}

public sealed class ClipboardOleRawData
{
    public ushort FormatId { get; init; }
    public string FormatName { get; init; } = "";
    public TYMED Tymed { get; init; }
    public byte[] RawBytes { get; init; } = Array.Empty<byte>();
}

public sealed class ClipboardOleDescriptorInfo
{
    public ushort FormatId { get; init; }
    public string FormatName { get; init; } = "";
    public TYMED Tymed { get; init; }

    public Guid ClassId { get; init; }
    public string ProgId { get; init; } = "";
    public string FullUserTypeName { get; init; } = "";
    public string SourceOfCopy { get; init; } = "";

    public uint DrawAspect { get; init; }
    public uint Status { get; init; }

    public string OleKind { get; init; } = "Unknown";
}

public sealed class PersistedOlePackage
{
    public int Version { get; init; } = 1;
    public string DisplayName { get; init; } = "Clipboard OLE Object";
    public DateTime SavedAtUtc { get; init; } = DateTime.UtcNow;
    public byte[] StorageBytes { get; init; } = Array.Empty<byte>();
}

public sealed class ClipboardDrawableOleObject : IDisposable
{
    private object? _oleObject;
    private IStorage? _storage;
    private ILockBytes? _lockBytes;
    private readonly Control _hostControl;

    private OleClientSiteBridge? _clientSite;
    private OleAdviseSinkBridge? _adviseSink;
    private uint _oleAdviseCookie;
    private bool _viewAdviseAttached;
    private bool _isSaving;

    private const int OLEIVERB_PRIMARY = 0;
    private const int OLEIVERB_SHOW = -1;
    private const int OLEIVERB_OPEN = -2;

    internal ClipboardDrawableOleObject(object oleObject, IStorage storage, ILockBytes lockBytes, Control hostControl)
    {
        _oleObject = oleObject;
        _storage = storage;
        _lockBytes = lockBytes;
        _hostControl = hostControl ?? throw new ArgumentNullException(nameof(hostControl));

        AttachCallbacks();
    }

    public PersistedOlePackage ExportPackage(string displayName = "Clipboard OLE Object")
    {
        SaveToBackingStorage();

        return new PersistedOlePackage
        {
            Version = 1,
            DisplayName = displayName,
            SavedAtUtc = DateTime.UtcNow,
            StorageBytes = GetBackingStorageBytes()
        };
    }

    public void OpenEditor(
        IntPtr hwndParent,
        string containerAppName = "WinFormsApp4",
        string? documentName = "Embedded Object")
    {
        if (_oleObject == null)
            throw new ObjectDisposedException(nameof(ClipboardDrawableOleObject));
        if (_clientSite == null)
            throw new ObjectDisposedException(nameof(ClipboardDrawableOleObject));

        if (_oleObject is not IOleObject ole)
            throw new InvalidOperationException("当前对象不支持 IOleObject。");

        ole.SetHostNames(containerAppName, documentName);

        var rc = new RECT();
        HRESULT hr = ole.DoVerb(
            OLEIVERB_OPEN,
            default,
            _clientSite,
            0,
            hwndParent,
            rc);

        if (hr == HRESULT.OLEOBJ_S_INVALIDVERB || hr == HRESULT.E_NOTIMPL)
        {
            hr = ole.DoVerb(
                OLEIVERB_PRIMARY,
                default,
                _clientSite,
                0,
                hwndParent,
                rc);
        }

        if (hr == HRESULT.OLEOBJ_S_INVALIDVERB || hr == HRESULT.E_NOTIMPL)
        {
            hr = ole.DoVerb(
                OLEIVERB_SHOW,
                default,
                _clientSite,
                0,
                hwndParent,
                rc);
        }

        hr.ThrowIfFailed();
    }

    public Size? GetNaturalPixelSize(float dpiX, float dpiY)
    {
        if (_oleObject == null)
            return null;

        if (_oleObject is not IOleObject ole)
            return null;

        HRESULT hr = ole.GetExtent(DVASPECT.DVASPECT_CONTENT, out SIZE sz);
        if (hr != 0 || sz.cx <= 0 || sz.cy <= 0)
            return null;

        int pxW = (int)Math.Round(sz.cx * dpiX / 2540.0);
        int pxH = (int)Math.Round(sz.cy * dpiY / 2540.0);

        if (pxW <= 0 || pxH <= 0)
            return null;

        return new Size(pxW, pxH);
    }

    public void Draw(Graphics g, Rectangle bounds)
    {
        if (_oleObject == null)
            throw new ObjectDisposedException(nameof(ClipboardDrawableOleObject));

        IntPtr hdc = g.GetHdc();
        try
        {
            var rc = new RECT(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
            OleDraw(_oleObject, DVASPECT.DVASPECT_CONTENT, hdc, rc).ThrowIfFailed();
        }
        finally
        {
            g.ReleaseHdc(hdc);
        }
    }

    public byte[] GetBackingStorageBytes()
    {
        if (_lockBytes == null)
            throw new ObjectDisposedException(nameof(ClipboardDrawableOleObject));
        if (_storage != null)
            _storage.Commit(0);

        GetHGlobalFromILockBytes(_lockBytes, out var hGlobal).ThrowIfFailed();
        return VanaraClipboardOle.HGlobalToBytes(hGlobal);
    }

    internal void NotifyHostChanged()
    {
        if (_hostControl.IsDisposed)
            return;

        void RefreshCore()
        {
            if (_hostControl.IsDisposed)
                return;

            _hostControl.Invalidate();
            _hostControl.Update();
        }

        if (_hostControl.InvokeRequired)
            _hostControl.BeginInvoke((Action)RefreshCore);
        else
            RefreshCore();
    }

    internal HRESULT SaveForClientSite()
    {
        try
        {
            SaveToBackingStorage();
            return HRESULT.S_OK;
        }
        catch
        {
            return HRESULT.E_FAIL;
        }
    }

    private void SaveToBackingStorage()
    {
        if (_isSaving)
            return;

        if (_oleObject == null)
            throw new ObjectDisposedException(nameof(ClipboardDrawableOleObject));
        if (_storage == null)
            throw new ObjectDisposedException(nameof(ClipboardDrawableOleObject));

        if (_oleObject is not IPersistStorage persist)
            return;

        _isSaving = true;
        try
        {
            HRESULT dirtyHr = persist.IsDirty();
            if (dirtyHr == HRESULT.S_FALSE)
                return;

            dirtyHr.ThrowIfFailed();

            OleSave(persist, _storage, true).ThrowIfFailed();
            persist.SaveCompleted(null);
            _storage.Commit(0);

            NotifyHostChanged();
        }
        finally
        {
            _isSaving = false;
        }
    }

    private void AttachCallbacks()
    {
        if (_oleObject == null)
            return;

        _clientSite ??= new OleClientSiteBridge(this);
        _adviseSink ??= new OleAdviseSinkBridge(this);

        if (_oleObject is IOleObject ole)
        {
            ole.SetClientSite(_clientSite).ThrowIfFailed();

            HRESULT hrAdvise = ole.Advise(_adviseSink, out _oleAdviseCookie);
            if (hrAdvise.Failed)
                _oleAdviseCookie = 0;
        }

        if (_oleObject is IViewObject view)
        {
            HRESULT hr = view.SetAdvise(
                DVASPECT.DVASPECT_CONTENT,
                ADVF.ADVF_PRIMEFIRST,
                _adviseSink);

            _viewAdviseAttached = hr.Succeeded;
        }
    }

    public void Dispose()
    {
        if (_oleObject is IOleObject ole && _oleAdviseCookie != 0)
        {
            try { ole.Unadvise(_oleAdviseCookie); } catch { }
            _oleAdviseCookie = 0;
        }

        if (_oleObject is IViewObject view && _viewAdviseAttached)
        {
            try { view.SetAdvise(DVASPECT.DVASPECT_CONTENT, 0, null); } catch { }
            _viewAdviseAttached = false;
        }

        if (_oleObject != null && Marshal.IsComObject(_oleObject))
        {
            if (_oleObject is IOleObject oleForClose)
            {
                try { oleForClose.Close(OLECLOSE.OLECLOSE_NOSAVE); } catch { }
                try { oleForClose.SetClientSite(null); } catch { }
            }

            Marshal.ReleaseComObject(_oleObject);
            _oleObject = null;
        }

        if (_storage != null && Marshal.IsComObject(_storage))
        {
            Marshal.ReleaseComObject(_storage);
            _storage = null;
        }

        if (_lockBytes != null && Marshal.IsComObject(_lockBytes))
        {
            Marshal.ReleaseComObject(_lockBytes);
            _lockBytes = null;
        }

        _adviseSink = null;
        _clientSite = null;
    }
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class OleClientSiteBridge : IOleClientSite
{
    private readonly ClipboardDrawableOleObject _owner;

    public OleClientSiteBridge(ClipboardDrawableOleObject owner)
    {
        _owner = owner;
    }

    public HRESULT SaveObject() => _owner.SaveForClientSite();

    public HRESULT ShowObject()
    {
        _owner.NotifyHostChanged();
        return HRESULT.S_OK;
    }

    public HRESULT OnShowWindow(bool fShow)
    {
        _owner.NotifyHostChanged();
        return HRESULT.S_OK;
    }

    public HRESULT RequestNewObjectLayout()
    {
        _owner.NotifyHostChanged();
        return HRESULT.S_OK;
    }

    public HRESULT GetMoniker(OLEGETMONIKER dwAssign, OLEWHICHMK dwWhichMoniker, out IMoniker? ppmk)
    {
        ppmk = null;
        return HRESULT.E_NOTIMPL;
    }

    public HRESULT GetContainer(out IOleContainer? ppContainer)
    {
        ppContainer = null;
        return HRESULT.E_NOINTERFACE;
    }
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public sealed class OleAdviseSinkBridge : IAdviseSink
{
    private readonly ClipboardDrawableOleObject _owner;

    public OleAdviseSinkBridge(ClipboardDrawableOleObject owner)
    {
        _owner = owner;
    }

    public void OnDataChange(ref FORMATETC pFormatetc, ref STGMEDIUM pStgmed)
    {
        _owner.NotifyHostChanged();
    }

    public void OnSave()
    {
        _owner.NotifyHostChanged();
    }

    public void OnClose()
    {
        _owner.NotifyHostChanged();
        _owner.SaveForClientSite();
    }

    public void OnRename(IMoniker moniker)
    {
        _owner.NotifyHostChanged();
    }

    public void OnViewChange(int aspect, int index)
    {
        _owner.NotifyHostChanged();
    }
}

public static class VanaraClipboardOle
{
    private const int OLE_S_STATIC = 0x00040001;

    private static readonly Guid IID_IUnknown =
        new("00000000-0000-0000-C000-000000000046");

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int ProgIDFromCLSID(ref Guid clsid, out IntPtr lplpszProgID);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, EntryPoint = "OleCreateFromFile")]
    private static extern int OleCreateFromFileNative(
        [In] ref Guid rclsid,
        [MarshalAs(UnmanagedType.LPWStr)] string lpszFileName,
        [In] ref Guid riid,
        OLERENDER renderopt,
        [In] ref FORMATETC pFormatEtc,
        [MarshalAs(UnmanagedType.Interface)] IOleClientSite? pClientSite,
        [MarshalAs(UnmanagedType.Interface)] IStorage pStg,
        [MarshalAs(UnmanagedType.Interface)] out object ppvObj);

    [StructLayout(LayoutKind.Sequential)]
    private struct SIZEL_NATIVE
    {
        public int cx;
        public int cy;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINTL_NATIVE
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct OBJECTDESCRIPTOR_NATIVE
    {
        public uint cbSize;
        public Guid clsid;
        public uint dwDrawAspect;
        public SIZEL_NATIVE sizel;
        public POINTL_NATIVE pointl;
        public uint dwStatus;
        public uint dwFullUserTypeName;
        public uint dwSrcOfCopy;
    }

    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public static List<ClipboardFormatInfo> GetClipboardFormats()
    {
        OleGetClipboard(out System.Runtime.InteropServices.ComTypes.IDataObject dataObj).ThrowIfFailed();

        try
        {
            var list = new List<ClipboardFormatInfo>();
            IEnumFORMATETC enumFmt = dataObj.EnumFormatEtc(DATADIR.DATADIR_GET);

            var arr = new FORMATETC[1];
            var fetched = new int[1];

            while (enumFmt.Next(1, arr, fetched) == 0 && fetched[0] == 1)
            {
                ushort cf = unchecked((ushort)arr[0].cfFormat);
                list.Add(new ClipboardFormatInfo
                {
                    FormatId = cf,
                    FormatName = GetClipboardFormatDisplayName(cf),
                    Tymed = arr[0].tymed,
                    Aspect = arr[0].dwAspect
                });
            }

            return list;
        }
        finally
        {
            ReleaseComObjectSafe(dataObj);
        }
    }

    public static bool TryGetObjectDescriptorFromClipboard(out ClipboardOleDescriptorInfo info)
    {
        info = null!;

        OleGetClipboard(out System.Runtime.InteropServices.ComTypes.IDataObject dataObj).ThrowIfFailed();
        try
        {
            var cfObjectDescriptor = User32.RegisterClipboardFormat("Object Descriptor");

            if (!TryReadRawFormat(dataObj, (ushort)cfObjectDescriptor, "Object Descriptor", out ClipboardOleRawData raw))
                return false;

            info = ParseObjectDescriptor(raw);
            return true;
        }
        finally
        {
            ReleaseComObjectSafe(dataObj);
        }
    }

    public static bool TryGetPreferredRawOleDataFromClipboard(out ClipboardOleRawData raw)
    {
        raw = null!;

        OleGetClipboard(out System.Runtime.InteropServices.ComTypes.IDataObject dataObj).ThrowIfFailed();
        try
        {
            var cfEmbeddedObject = User32.RegisterClipboardFormat("Embedded Object");
            var cfEmbedSource = User32.RegisterClipboardFormat("Embed Source");

            if (TryReadRawFormat(dataObj, (ushort)cfEmbeddedObject, "Embedded Object", out raw))
                return true;

            if (TryReadRawFormat(dataObj, (ushort)cfEmbedSource, "Embed Source", out raw))
                return true;

            return false;
        }
        finally
        {
            ReleaseComObjectSafe(dataObj);
        }
    }

    public static ClipboardDrawableOleObject? CreateDrawableObjectFromClipboard(Control hostControl)
    {
        if (hostControl == null)
            throw new ArgumentNullException(nameof(hostControl));

        OleGetClipboard(out System.Runtime.InteropServices.ComTypes.IDataObject dataObj).ThrowIfFailed();

        try
        {
            HRESULT qhr = OleQueryCreateFromData(dataObj);
            if ((int)qhr == OLE_S_STATIC)
            {
                return null;
            }
            else
            {
                qhr.ThrowIfFailed();
            }

            CreateILockBytesOnHGlobal(IntPtr.Zero, true, out ILockBytes lockBytes).ThrowIfFailed();

            IStorage? storage = null;
            object? oleObj = null;

            try
            {
                StgCreateDocfileOnILockBytes(
                    lockBytes,
                    STGM.STGM_CREATE | STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_TRANSACTED,
                    0,
                    out storage).ThrowIfFailed();

                var renderFmt = new FORMATETC
                {
                    cfFormat = 0,
                    ptd = IntPtr.Zero,
                    dwAspect = DVASPECT.DVASPECT_CONTENT,
                    lindex = -1,
                    tymed = TYMED.TYMED_NULL
                };

                OleCreateFromData(
                    dataObj,
                    in IID_IUnknown,
                    OLERENDER.OLERENDER_DRAW,
                    in renderFmt,
                    null,
                    storage,
                    out oleObj).ThrowIfFailed();

                storage.Commit(0);
                return new ClipboardDrawableOleObject(oleObj, storage, lockBytes, hostControl);
            }
            catch
            {
                if (oleObj != null && Marshal.IsComObject(oleObj))
                    Marshal.ReleaseComObject(oleObj);

                if (storage != null && Marshal.IsComObject(storage))
                    Marshal.ReleaseComObject(storage);

                if (lockBytes != null && Marshal.IsComObject(lockBytes))
                    Marshal.ReleaseComObject(lockBytes);

                throw;
            }
        }
        finally
        {
            ReleaseComObjectSafe(dataObj);
        }
    }

    public static ClipboardDrawableOleObject CreateDrawableObjectFromFile(string filePath, Control hostControl)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentNullException(nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException("文件不存在。", filePath);

        if (hostControl == null)
            throw new ArgumentNullException(nameof(hostControl));

        CreateILockBytesOnHGlobal(IntPtr.Zero, true, out ILockBytes lockBytes).ThrowIfFailed();

        IStorage? storage = null;
        object? oleObj = null;

        try
        {
            StgCreateDocfileOnILockBytes(
                lockBytes,
                STGM.STGM_CREATE | STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_TRANSACTED,
                0,
                out storage).ThrowIfFailed();

            Guid clsid = Guid.Empty;
            Guid iid = IID_IUnknown;

            var renderFmt = new FORMATETC
            {
                cfFormat = 0,
                ptd = IntPtr.Zero,
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                tymed = TYMED.TYMED_NULL
            };

            HRESULT hr = (HRESULT)OleCreateFromFileNative(
                ref clsid,
                filePath,
                ref iid,
                OLERENDER.OLERENDER_DRAW,
                ref renderFmt,
                null,
                storage,
                out oleObj);

            hr.ThrowIfFailed();

            storage.Commit(0);
            return new ClipboardDrawableOleObject(oleObj, storage, lockBytes, hostControl);
        }
        catch
        {
            if (oleObj != null && Marshal.IsComObject(oleObj))
                Marshal.ReleaseComObject(oleObj);

            if (storage != null && Marshal.IsComObject(storage))
                Marshal.ReleaseComObject(storage);

            if (lockBytes != null && Marshal.IsComObject(lockBytes))
                Marshal.ReleaseComObject(lockBytes);

            throw;
        }
    }

    public static ClipboardDrawableOleObject OpenFromPersistedPackage(PersistedOlePackage package, Control hostControl)
    {
        if (package == null)
            throw new ArgumentNullException(nameof(package));
        if (hostControl == null)
            throw new ArgumentNullException(nameof(hostControl));
        if (package.StorageBytes == null || package.StorageBytes.Length == 0)
            throw new ArgumentException("StorageBytes 为空。", nameof(package));

        IntPtr hGlobal = IntPtr.Zero;
        ILockBytes? lockBytes = null;
        IStorage? storage = null;
        object? oleObj = null;

        try
        {
            hGlobal = BytesToHGlobal(package.StorageBytes);

            CreateILockBytesOnHGlobal(hGlobal, true, out lockBytes).ThrowIfFailed();
            hGlobal = IntPtr.Zero;

            StgOpenStorageOnILockBytes(
                lockBytes,
                null,
                STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_TRANSACTED,
                null,
                0,
                out storage).ThrowIfFailed();

            OleLoad(
                storage,
                in IID_IUnknown,
                null,
                out oleObj).ThrowIfFailed();

            return new ClipboardDrawableOleObject(oleObj, storage, lockBytes, hostControl);
        }
        catch
        {
            if (oleObj != null && Marshal.IsComObject(oleObj))
                Marshal.ReleaseComObject(oleObj);

            if (storage != null && Marshal.IsComObject(storage))
                Marshal.ReleaseComObject(storage);

            if (lockBytes != null && Marshal.IsComObject(lockBytes))
                Marshal.ReleaseComObject(lockBytes);

            if (hGlobal != IntPtr.Zero)
                Kernel32.GlobalFree(hGlobal);

            throw;
        }
    }

    private static ClipboardOleDescriptorInfo ParseObjectDescriptor(ClipboardOleRawData raw)
    {
        if (raw.RawBytes == null || raw.RawBytes.Length == 0)
            throw new InvalidOperationException("Object Descriptor 数据为空。");

        int size = Marshal.SizeOf<OBJECTDESCRIPTOR_NATIVE>();
        if (raw.RawBytes.Length < size)
            throw new InvalidOperationException("Object Descriptor 数据长度不足。");

        var handle = GCHandle.Alloc(raw.RawBytes, GCHandleType.Pinned);
        try
        {
            IntPtr ptr = handle.AddrOfPinnedObject();
            OBJECTDESCRIPTOR_NATIVE native = Marshal.PtrToStructure<OBJECTDESCRIPTOR_NATIVE>(ptr);

            string fullUserTypeName = ReadDescriptorString(raw.RawBytes, (int)native.dwFullUserTypeName);
            string srcOfCopy = ReadDescriptorString(raw.RawBytes, (int)native.dwSrcOfCopy);
            string progId = TryGetProgId(native.clsid);
            string oleKind = ClassifyOleKind(progId, fullUserTypeName, srcOfCopy);

            return new ClipboardOleDescriptorInfo
            {
                FormatId = raw.FormatId,
                FormatName = raw.FormatName,
                Tymed = raw.Tymed,
                ClassId = native.clsid,
                ProgId = progId,
                FullUserTypeName = fullUserTypeName,
                SourceOfCopy = srcOfCopy,
                DrawAspect = native.dwDrawAspect,
                Status = native.dwStatus,
                OleKind = oleKind
            };
        }
        finally
        {
            handle.Free();
        }
    }

    private static bool TryReadRawFormat(
        System.Runtime.InteropServices.ComTypes.IDataObject dataObj,
        ushort cfFormat,
        string formatName,
        out ClipboardOleRawData raw)
    {
        raw = null!;

        foreach (TYMED tymed in new[] { TYMED.TYMED_HGLOBAL, TYMED.TYMED_ISTREAM, TYMED.TYMED_ISTORAGE })
        {
            var fmt = new FORMATETC
            {
                cfFormat = unchecked((short)cfFormat),
                dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1,
                ptd = IntPtr.Zero,
                tymed = tymed
            };

            int q = dataObj.QueryGetData(ref fmt);
            if (q != 0)
                continue;

            dataObj.GetData(ref fmt, out STGMEDIUM stg);
            try
            {
                switch (stg.tymed)
                {
                    case TYMED.TYMED_HGLOBAL:
                        raw = new ClipboardOleRawData
                        {
                            FormatId = cfFormat,
                            FormatName = formatName,
                            Tymed = stg.tymed,
                            RawBytes = HGlobalToBytes(stg.unionmember)
                        };
                        return true;

                    case TYMED.TYMED_ISTREAM:
                        {
                            var stream = (IStream)Marshal.GetObjectForIUnknown(stg.unionmember);
                            try
                            {
                                raw = new ClipboardOleRawData
                                {
                                    FormatId = cfFormat,
                                    FormatName = formatName,
                                    Tymed = stg.tymed,
                                    RawBytes = ReadAllBytesAndRewind(stream)
                                };
                                return true;
                            }
                            finally
                            {
                                ReleaseComObjectSafe(stream);
                            }
                        }

                    case TYMED.TYMED_ISTORAGE:
                        {
                            var storage = (IStorage)Marshal.GetObjectForIUnknown(stg.unionmember);
                            try
                            {
                                raw = new ClipboardOleRawData
                                {
                                    FormatId = cfFormat,
                                    FormatName = formatName,
                                    Tymed = stg.tymed,
                                    RawBytes = StorageToBytes(storage)
                                };
                                return true;
                            }
                            finally
                            {
                                ReleaseComObjectSafe(storage);
                            }
                        }
                }
            }
            finally
            {
                ReleaseStgMedium(stg);
            }
        }

        return false;
    }

    private static byte[] StorageToBytes(IStorage sourceStorage)
    {
        CreateILockBytesOnHGlobal(IntPtr.Zero, true, out ILockBytes lockBytes).ThrowIfFailed();

        IStorage? destStorage = null;
        try
        {
            StgCreateDocfileOnILockBytes(
                lockBytes,
                STGM.STGM_CREATE | STGM.STGM_READWRITE | STGM.STGM_SHARE_EXCLUSIVE | STGM.STGM_TRANSACTED,
                0,
                out destStorage).ThrowIfFailed();

            sourceStorage.CopyTo(0, null, IntPtr.Zero, destStorage);
            destStorage.Commit(0);

            GetHGlobalFromILockBytes(lockBytes, out var hGlobal).ThrowIfFailed();
            return HGlobalToBytes(hGlobal);
        }
        finally
        {
            ReleaseComObjectSafe(destStorage);
            ReleaseComObjectSafe(lockBytes);
        }
    }

    private static byte[] ReadAllBytesAndRewind(IStream stream)
    {
        using var ms = new MemoryStream();
        byte[] buffer = new byte[8192];
        IntPtr pcbRead = Marshal.AllocCoTaskMem(sizeof(int));

        try
        {
            while (true)
            {
                stream.Read(buffer, buffer.Length, pcbRead);
                int read = Marshal.ReadInt32(pcbRead);
                if (read <= 0)
                    break;

                ms.Write(buffer, 0, read);
            }

            stream.Seek(0, 0, IntPtr.Zero);
            return ms.ToArray();
        }
        finally
        {
            Marshal.FreeCoTaskMem(pcbRead);
        }
    }

    private static string ReadDescriptorString(byte[] bytes, int offset)
    {
        if (offset <= 0 || offset >= bytes.Length)
            return "";

        string ansi = ReadNullTerminatedAnsi(bytes, offset);
        string unicode = ReadNullTerminatedUnicode(bytes, offset);

        if (string.IsNullOrWhiteSpace(unicode))
            return ansi;

        if (string.IsNullOrWhiteSpace(ansi))
            return unicode;

        int ansiScore = ScoreString(ansi);
        int unicodeScore = ScoreString(unicode);

        if (unicodeScore > ansiScore)
            return unicode;

        if (ansiScore > unicodeScore)
            return ansi;

        return unicode.Length >= ansi.Length ? unicode : ansi;
    }

    private static string ReadNullTerminatedAnsi(byte[] bytes, int offset)
    {
        int end = offset;
        while (end < bytes.Length && bytes[end] != 0)
            end++;

        int len = end - offset;
        if (len <= 0)
            return "";

        return Encoding.Default.GetString(bytes, offset, len).Trim();
    }

    private static string ReadNullTerminatedUnicode(byte[] bytes, int offset)
    {
        int end = offset;

        while (end + 1 < bytes.Length)
        {
            if (bytes[end] == 0 && bytes[end + 1] == 0)
                break;

            end += 2;
        }

        int len = end - offset;
        if (len <= 0 || (len % 2) != 0)
            return "";

        return Encoding.Unicode.GetString(bytes, offset, len).Trim();
    }

    private static int ScoreString(string s)
    {
        if (string.IsNullOrEmpty(s))
            return 0;

        int score = 0;
        foreach (char ch in s)
        {
            if (!char.IsControl(ch))
                score += 2;

            if (char.IsLetterOrDigit(ch) || char.IsWhiteSpace(ch) || ch is '.' or '\\' or ':' or '_' or '-' or '(' or ')' or '/')
                score += 2;

            if (ch == '\uFFFD')
                score -= 6;
        }

        return score;
    }

    private static string TryGetProgId(Guid clsid)
    {
        IntPtr p = IntPtr.Zero;
        try
        {
            int hr = ProgIDFromCLSID(ref clsid, out p);
            if (hr != 0 || p == IntPtr.Zero)
                return "";

            return Marshal.PtrToStringUni(p) ?? "";
        }
        finally
        {
            if (p != IntPtr.Zero)
                Marshal.FreeCoTaskMem(p);
        }
    }

    private static string ClassifyOleKind(string progId, string fullUserTypeName, string srcOfCopy)
    {
        string s = $"{progId}|{fullUserTypeName}|{srcOfCopy}".ToLowerInvariant();

        if (s.Contains("excel"))
            return "Excel";

        if (s.Contains("word"))
            return "Word";

        if (s.Contains("powerpoint") || s.Contains("ppt"))
            return "PowerPoint";

        if (s.Contains("visio"))
            return "Visio";

        if (s.Contains("acrobat") || s.Contains("pdf"))
            return "PDF";

        if (s.Contains("package"))
            return "Package";

        if (s.Contains("paint") || s.Contains("bitmap") || s.Contains("image") || s.Contains("png") || s.Contains("jpg") || s.Contains("jpeg"))
            return "Image";

        if (s.Contains("static"))
            return "Static";

        return "Other/Unknown";
    }

    internal static byte[] HGlobalToBytes(IntPtr hglobal)
    {
        int size = checked((int)Kernel32.GlobalSize(hglobal));
        IntPtr ptr = Kernel32.GlobalLock(hglobal);
        if (ptr == IntPtr.Zero)
            throw new InvalidOperationException("GlobalLock failed.");

        try
        {
            var bytes = new byte[size];
            Marshal.Copy(ptr, bytes, 0, size);
            return bytes;
        }
        finally
        {
            Kernel32.GlobalUnlock(hglobal);
        }
    }

    private static IntPtr BytesToHGlobal(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);

        var hGlobal = Kernel32.GlobalAlloc(Kernel32.GMEM.GHND, (UIntPtr)bytes.Length);
        if (hGlobal == IntPtr.Zero)
            throw new OutOfMemoryException("GlobalAlloc failed.");

        IntPtr ptr = Kernel32.GlobalLock(hGlobal);
        if (ptr == IntPtr.Zero)
        {
            Kernel32.GlobalFree(hGlobal);
            throw new InvalidOperationException("GlobalLock failed.");
        }

        try
        {
            Marshal.Copy(bytes, 0, ptr, bytes.Length);
        }
        finally
        {
            Kernel32.GlobalUnlock(hGlobal);
        }

        return hGlobal.DangerousGetHandle();
    }

    private static string GetClipboardFormatDisplayName(ushort cf)
    {
        string? std = GetStandardClipboardFormatName(cf);
        if (std != null)
            return std;

        var sb = new StringBuilder(256);
        int len = User32.GetClipboardFormatName(cf, sb, sb.Capacity);
        if (len > 0)
            return sb.ToString();

        return $"Format_{cf}";
    }

    private static string? GetStandardClipboardFormatName(ushort cf) => cf switch
    {
        1 => "CF_TEXT",
        2 => "CF_BITMAP",
        3 => "CF_METAFILEPICT",
        8 => "CF_DIB",
        13 => "CF_UNICODETEXT",
        14 => "CF_ENHMETAFILE",
        15 => "CF_HDROP",
        17 => "CF_DIBV5",
        _ => null
    };

    private static void ReleaseComObjectSafe(object? obj)
    {
        if (obj != null && Marshal.IsComObject(obj))
            Marshal.ReleaseComObject(obj);
    }
}