using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace WpfApp1.Views
{
    public partial class DrawBoardPage : Page
    {
        // ── Win32 剪贴板 ──
        [DllImport("user32.dll")] private static extern bool OpenClipboard(IntPtr hWndNewOwner);
        [DllImport("user32.dll")] private static extern bool CloseClipboard();
        [DllImport("user32.dll")] private static extern bool EmptyClipboard();
        [DllImport("user32.dll")] private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);
        private const uint CF_UNICODETEXT = 13;

        private const string ExcalidrawUrl = "https://excalidraw.com";
        private bool _isWebViewReady = false;

        public DrawBoardPage()
        {
            InitializeComponent();
            Loaded += DrawBoardPage_Loaded;
        }

        // ═══════════════════════════════════════
        // 初始化 WebView2
        // ═══════════════════════════════════════

        private async void DrawBoardPage_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 设置 WebView2 用户数据目录（避免权限问题）
                var userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "CCToolbox", "WebView2");

                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await ExcalidrawWebView.EnsureCoreWebView2Async(env);

                _isWebViewReady = true;

                // 配置 WebView2
                var settings = ExcalidrawWebView.CoreWebView2.Settings;
                settings.IsStatusBarEnabled = false;
                settings.AreDefaultContextMenusEnabled = true;
                settings.IsZoomControlEnabled = true;
                settings.AreDevToolsEnabled = false;

                // 监听导航完成
                ExcalidrawWebView.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
                ExcalidrawWebView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

                // 加载 Excalidraw
                ExcalidrawWebView.CoreWebView2.Navigate(ExcalidrawUrl);
                TxtStatus.Text = "  连接中...";
            }
            catch (Exception ex)
            {
                TxtLoadingMsg.Text = $"❌ WebView2 初始化失败";
                TxtStatus.Text = "  初始化失败";
                MessageBox.Show(
                    $"WebView2 初始化失败：\n\n{ex.Message}\n\n" +
                    "请确认已安装 Microsoft Edge WebView2 Runtime。\n" +
                    "下载地址：https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
                    "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CoreWebView2_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                if (e.IsSuccess)
                {
                    LoadingOverlay.Visibility = Visibility.Collapsed;
                    TxtStatus.Text = "  ✅ 就绪";

                    // 注入辅助 JS（用于导出功能）
                    InjectHelperScript();
                }
                else
                {
                    TxtLoadingMsg.Text = "❌ 页面加载失败，请检查网络连接";
                    TxtStatus.Text = $"  加载失败 ({e.WebErrorStatus})";
                }
            });
        }

        private void CoreWebView2_WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            // 接收来自 JS 的消息（导出数据等）
            try
            {
                var message = e.TryGetWebMessageAsString();
                if (message != null && message.StartsWith("EXPORT:"))
                {
                    var data = message.Substring(7);
                    HandleExportMessage(data);
                }
                else if (message != null && message.StartsWith("SCENE:"))
                {
                    var data = message.Substring(6);
                    HandleSceneData(data);
                }
                else if (message != null && message.StartsWith("LOAD:"))
                {
                    var result = message.Substring(5);
                    Dispatcher.Invoke(() =>
                    {
                        if (result == "OK")
                        {
                            TxtStatus.Text = "  ✅ 文件加载成功";
                        }
                        else if (result.StartsWith("ERROR:"))
                        {
                            var errMsg = result.Substring(6);
                            TxtStatus.Text = $"  ❌ 加载失败: {errMsg}";
                        }
                    });
                }
            }
            catch { }
        }

        // ═══════════════════════════════════════
        // 注入辅助 JS
        // ═══════════════════════════════════════

        private async void InjectHelperScript()
        {
            if (!_isWebViewReady) return;

            var js = @"
                // 辅助函数：获取 Excalidraw 场景数据
                window.__ccGetScene = async function() {
                    try {
                        // Excalidraw 将数据存储在 localStorage
                        var data = localStorage.getItem('excalidraw');
                        if (data) {
                            window.chrome.webview.postMessage('SCENE:' + data);
                        } else {
                            window.chrome.webview.postMessage('SCENE:{}');
                        }
                    } catch(e) {
                        window.chrome.webview.postMessage('SCENE:{}');
                    }
                };

                // 辅助函数：导出为 PNG（通过 canvas）
                window.__ccExportPng = async function() {
                    try {
                        var canvas = document.querySelector('canvas');
                        if (canvas) {
                            var dataUrl = canvas.toDataURL('image/png');
                            window.chrome.webview.postMessage('EXPORT:PNG:' + dataUrl);
                        } else {
                            window.chrome.webview.postMessage('EXPORT:ERROR:找不到画布元素');
                        }
                    } catch(e) {
                        window.chrome.webview.postMessage('EXPORT:ERROR:' + e.message);
                    }
                };

                // 辅助函数：导出为 SVG
                window.__ccExportSvg = async function() {
                    try {
                        // 尝试通过 Excalidraw 的内部 API 获取 SVG
                        var svgEl = document.querySelector('svg.excalidraw-svg');
                        if (svgEl) {
                            var svgData = new XMLSerializer().serializeToString(svgEl);
                            window.chrome.webview.postMessage('EXPORT:SVG:' + svgData);
                        } else {
                            // 回退：从 canvas 导出
                            window.chrome.webview.postMessage('EXPORT:ERROR:请使用 Excalidraw 自带的导出功能（左上角菜单）');
                        }
                    } catch(e) {
                        window.chrome.webview.postMessage('EXPORT:ERROR:' + e.message);
                    }
                };

                // 辅助函数：加载场景数据（通过 Excalidraw 内部 API）
                window.__ccLoadScene = function(jsonStr) {
                    try {
                        var data = JSON.parse(jsonStr);
                        
                        // 分离 elements 和 appState
                        var elements = data.elements || [];
                        var appState = data.appState || {};
                        var files = data.files || {};
                        
                        // 尝试通过 Excalidraw 内部 API 加载
                        // 方式1：查找 React Fiber 上的 updateScene
                        var excalidrawEl = document.querySelector('.excalidraw');
                        if (excalidrawEl) {
                            // 遍历 React Fiber 树查找 Excalidraw 实例
                            var fiberKey = Object.keys(excalidrawEl).find(k => k.startsWith('__reactFiber$') || k.startsWith('__reactInternalInstance$'));
                            if (fiberKey) {
                                var fiber = excalidrawEl[fiberKey];
                                var node = fiber;
                                // 向上遍历找到有 updateScene 的组件
                                for (var i = 0; i < 50 && node; i++) {
                                    if (node.memoizedProps && node.memoizedProps.excalidrawAPI) {
                                        var api = node.memoizedProps.excalidrawAPI;
                                        if (api && api.updateScene) {
                                            api.updateScene({ elements: elements });
                                            if (api.addFiles && Object.keys(files).length > 0) {
                                                api.addFiles(Object.values(files));
                                            }
                                            api.scrollToContent();
                                            window.chrome.webview.postMessage('LOAD:OK');
                                            return;
                                        }
                                    }
                                    if (node.stateNode && node.stateNode.updateScene) {
                                        node.stateNode.updateScene({ elements: elements });
                                        window.chrome.webview.postMessage('LOAD:OK');
                                        return;
                                    }
                                    node = node.return;
                                }
                            }
                        }
                        
                        // 方式2：通过 Blob URL 重新加载
                        var blob = new Blob([jsonStr], { type: 'application/json' });
                        var url = URL.createObjectURL(blob);
                        
                        // 使用 Excalidraw 的文件加载机制
                        // 构造一个 File 对象并触发 loadFromBlob
                        var file = new File([jsonStr], 'drawing.excalidraw', { type: 'application/json' });
                        
                        // 尝试调用全局的 loadFromBlob（如果存在）
                        if (window.loadFromBlob) {
                            window.loadFromBlob(blob, null, null).then(function(data) {
                                window.chrome.webview.postMessage('LOAD:OK');
                            });
                            return;
                        }
                        
                        // 方式3：回退到 localStorage + reload（兼容旧版本）
                        // 同时写入多个可能的 key
                        localStorage.setItem('excalidraw', jsonStr);
                        localStorage.setItem('excalidraw-state', jsonStr);
                        
                        // 尝试写入 IndexedDB
                        var dbReq = indexedDB.open('excalidraw', 1);
                        dbReq.onsuccess = function(event) {
                            try {
                                var db = event.target.result;
                                var storeNames = db.objectStoreNames;
                                if (storeNames.length > 0) {
                                    var tx = db.transaction(storeNames[0], 'readwrite');
                                    var store = tx.objectStore(storeNames[0]);
                                    store.put(data, 'scene');
                                }
                            } catch(e) {}
                            location.reload();
                        };
                        dbReq.onerror = function() {
                            location.reload();
                        };
                    } catch(e) {
                        alert('加载失败: ' + e.message);
                    }
                };

                console.log('[CC Toolbox] Helper scripts injected.');
            ";

            try
            {
                await ExcalidrawWebView.CoreWebView2.ExecuteScriptAsync(js);
            }
            catch { }
        }

        // ═══════════════════════════════════════
        // 打开 / 保存 .excalidraw 文件
        // ═══════════════════════════════════════

        private async void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            if (!_isWebViewReady) return;

            var dlg = new OpenFileDialog
            {
                Filter = "Excalidraw 文件|*.excalidraw;*.excalidraw.json|JSON 文件|*.json|所有文件|*.*",
                Title = "打开 Excalidraw 文件"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var json = File.ReadAllText(dlg.FileName, Encoding.UTF8);
                    var fileName = Path.GetFileName(dlg.FileName);

                    // 使用 Base64 传输避免转义问题
                    var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

                    // 通过模拟拖放 File 对象来加载（Excalidraw 原生支持）
                    var loadScript = $@"
                        (async function() {{
                            try {{
                                var base64 = '{base64}';
                                var raw = atob(base64);
                                var bytes = new Uint8Array(raw.length);
                                for (var i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
                                var blob = new Blob([bytes], {{ type: 'application/json' }});
                                var file = new File([blob], '{fileName.Replace("'", "\\'")}', {{ type: 'application/json' }});

                                // 模拟拖放事件让 Excalidraw 原生处理文件
                                var dt = new DataTransfer();
                                dt.items.add(file);
                                var dropTarget = document.querySelector('.excalidraw') || document.querySelector('canvas') || document.body;
                                
                                var dragEnter = new DragEvent('dragenter', {{ dataTransfer: dt, bubbles: true }});
                                var dragOver = new DragEvent('dragover', {{ dataTransfer: dt, bubbles: true }});
                                var drop = new DragEvent('drop', {{ dataTransfer: dt, bubbles: true }});
                                
                                dropTarget.dispatchEvent(dragEnter);
                                dropTarget.dispatchEvent(dragOver);
                                dropTarget.dispatchEvent(drop);
                                
                                window.chrome.webview.postMessage('LOAD:OK');
                            }} catch(e) {{
                                window.chrome.webview.postMessage('LOAD:ERROR:' + e.message);
                            }}
                        }})();
                    ";

                    TxtStatus.Text = $"  ⏳ 正在加载: {fileName}";
                    await ExcalidrawWebView.CoreWebView2.ExecuteScriptAsync(loadScript);
                    TxtStatus.Text = $"  ✅ 已加载: {fileName}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"文件加载失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    TxtStatus.Text = "  ❌ 加载失败";
                }
            }
        }

        private string? _pendingSavePath;

        private async void BtnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            if (!_isWebViewReady) return;

            var dlg = new SaveFileDialog
            {
                Filter = "Excalidraw 文件|*.excalidraw|JSON 文件|*.json",
                FileName = $"drawing_{DateTime.Now:yyyyMMdd_HHmmss}.excalidraw",
                Title = "保存 Excalidraw 文件"
            };

            if (dlg.ShowDialog() == true)
            {
                _pendingSavePath = dlg.FileName;
                try
                {
                    await ExcalidrawWebView.CoreWebView2.ExecuteScriptAsync("window.__ccGetScene();");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"获取场景数据失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    _pendingSavePath = null;
                }
            }
        }

        private void HandleSceneData(string data)
        {
            Dispatcher.Invoke(() =>
            {
                if (_pendingSavePath != null)
                {
                    try
                    {
                        File.WriteAllText(_pendingSavePath, data, Encoding.UTF8);
                        TxtStatus.Text = $"  ✅ 已保存: {Path.GetFileName(_pendingSavePath)}";
                        MessageBox.Show($"文件已保存到：\n{_pendingSavePath}", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"保存失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    finally
                    {
                        _pendingSavePath = null;
                    }
                }
            });
        }

        // ═══════════════════════════════════════
        // 导出 PNG / SVG
        // ═══════════════════════════════════════

        private async void BtnExportPng_Click(object sender, RoutedEventArgs e)
        {
            if (!_isWebViewReady) return;

            try
            {
                await ExcalidrawWebView.CoreWebView2.ExecuteScriptAsync("window.__ccExportPng();");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnExportSvg_Click(object sender, RoutedEventArgs e)
        {
            if (!_isWebViewReady) return;

            try
            {
                await ExcalidrawWebView.CoreWebView2.ExecuteScriptAsync("window.__ccExportSvg();");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void HandleExportMessage(string data)
        {
            Dispatcher.Invoke(() =>
            {
                if (data.StartsWith("ERROR:"))
                {
                    var errMsg = data.Substring(6);
                    MessageBox.Show(errMsg, "导出提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (data.StartsWith("PNG:"))
                {
                    var dataUrl = data.Substring(4);
                    SaveDataUrlToFile(dataUrl, "PNG 图片|*.png", "drawing.png");
                }
                else if (data.StartsWith("SVG:"))
                {
                    var svgData = data.Substring(4);
                    var dlg = new SaveFileDialog
                    {
                        Filter = "SVG 文件|*.svg",
                        FileName = $"drawing_{DateTime.Now:yyyyMMdd_HHmmss}.svg"
                    };
                    if (dlg.ShowDialog() == true)
                    {
                        File.WriteAllText(dlg.FileName, svgData, Encoding.UTF8);
                        TxtStatus.Text = $"  ✅ SVG 已导出";
                        MessageBox.Show($"SVG 已导出到：\n{dlg.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            });
        }

        private void SaveDataUrlToFile(string dataUrl, string filter, string defaultName)
        {
            try
            {
                // data:image/png;base64,xxxxx
                var base64Start = dataUrl.IndexOf(",");
                if (base64Start < 0) return;

                var base64Data = dataUrl.Substring(base64Start + 1);
                var bytes = Convert.FromBase64String(base64Data);

                var dlg = new SaveFileDialog
                {
                    Filter = filter,
                    FileName = $"drawing_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(defaultName)}"
                };

                if (dlg.ShowDialog() == true)
                {
                    File.WriteAllBytes(dlg.FileName, bytes);
                    TxtStatus.Text = $"  ✅ PNG 已导出";
                    MessageBox.Show($"图片已导出到：\n{dlg.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ═══════════════════════════════════════
        // 工具栏按钮
        // ═══════════════════════════════════════

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            if (!_isWebViewReady) return;
            LoadingOverlay.Visibility = Visibility.Visible;
            TxtLoadingMsg.Text = "正在刷新...";
            TxtStatus.Text = "  刷新中...";
            ExcalidrawWebView.CoreWebView2.Reload();
        }

        private void BtnOpenInBrowser_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = ExcalidrawUrl,
                    UseShellExecute = true
                });
            }
            catch { }
        }
    }
}
