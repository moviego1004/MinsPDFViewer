using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Media.Imaging;

namespace MinsPDFViewer
{
    /// <summary>
    /// 화면 렌더 스케줄러 (전용 스레드 하나).
    /// <para>화면이 <see cref="UpdateViewport"/>로 "지금 보이는 페이지와 보이는 부분"을 알려 주면, 보이는 페이지 → 상세 타일 →
    /// 미리 그릴 페이지 → 주석 읽기 순으로 처리한다.</para>
    /// <para><b>해상도 보장</b>: 페이지마다 지금 그려진 사양(<see cref="PdfPageViewModel.RenderedBaseKey"/>)을 기록하고,
    /// 목표와 다르면 같아질 때까지 다시 그린다. 요청을 버리는 구조가 없다.</para>
    /// <para><b>즉시 취소</b>: 점진 렌더 중 보이는 영역이 바뀌어 이 작업이 더는 필요 없으면 멈춘다.</para>
    /// <para><b>고배율</b>: 전체 그림은 <see cref="MaxBasePixels"/> 이하로 줄여 그리고, 보이는 부분은 원래 해상도 타일로 겹친다.</para>
    /// </summary>
    public sealed class PageRenderScheduler
    {
        public static PageRenderScheduler Instance { get; } = new();

        /// <summary>페이지 전체 그림의 최대 픽셀 수 (약 48MB). 넘으면 줄여 그리고 상세 타일로 보완.</summary>
        public const long MaxBasePixels = 12_000_000;
        private const int DetailMargin = 96;
        private const int DetailGrid = 256;
        private const long CacheBudgetBytes = 384L * 1024 * 1024;

        /// <summary>렌더 대상. VisibleRect = 페이지 안에서 보이는 부분(DIP). 비어 있으면 미리 그리기만.</summary>
        public sealed record PageTarget(PdfPageViewModel Page, Rect VisibleRect)
        {
            public bool IsVisible => !VisibleRect.IsEmpty && VisibleRect.Width > 0 && VisibleRect.Height > 0;
        }

        private sealed record Viewport(PdfDocumentModel Model, IReadOnlyList<PageTarget> Targets, double DpiScale, long Generation);

        private enum JobKind { Base, Detail, Annotations }

        private sealed record Job(JobKind Kind, PdfDocumentModel Model, PdfPageViewModel Page, string Key, int FullWidth, int FullHeight,
            int Rotate, PixelRegion Region, double DpiScale, long Generation);

        private volatile Viewport? _viewport;
        private long _generation;
        private readonly AutoResetEvent _signal = new(false);
        private readonly ConcurrentDictionary<PdfPageViewModel, PdfDocumentModel> _requested = new();
        private readonly ConcurrentDictionary<(PdfPageViewModel Page, string Key), byte> _failed = new();
        private volatile bool _busy;

        // 비트맵 캐시 (LRU, 바이트 한도)
        private readonly object _cacheLock = new();
        private readonly Dictionary<(PdfDocumentModel Model, PdfPageViewModel Page, string Key), LinkedListNode<CacheEntry>> _cache = new();
        private readonly LinkedList<CacheEntry> _cacheOrder = new();
        private long _cacheBytes;

        private sealed record CacheEntry(PdfDocumentModel Model, PdfPageViewModel Page, string Key, BitmapSource Bitmap, long Bytes);

        private PageRenderScheduler()
        {
            var thread = new Thread(WorkerLoop) { IsBackground = true, Name = "PageRenderScheduler", Priority = ThreadPriority.AboveNormal };
            thread.Start();
        }

        /// <summary>렌더 작업 중인지 (썸네일 등이 양보할 때 확인).</summary>
        public bool HasPendingWork => _busy;

        // ── 화면 → 스케줄러 ──────────────────────────────────────────────

        /// <summary>
        /// 보이는 영역 갱신 (UI 스레드). targets는 우선순위 순 (보이는 페이지 먼저, 미리 그릴 페이지는 VisibleRect 비움).
        /// 캐시에 맞는 그림이 있으면 바로 붙이고, 멀리 벗어난 페이지의 그림은 내려놓는다(캐시에는 남음).
        /// </summary>
        public void UpdateViewport(PdfDocumentModel? model, IReadOnlyList<PageTarget> targets, double dpiScale)
        {
            if (model == null || model.IsDisposed)
            {
                _viewport = null;
                Interlocked.Increment(ref _generation);
                return;
            }

            var previous = _viewport;
            if (previous != null && !ReferenceEquals(previous.Model, model))
                _requested.Clear();

            long generation = Interlocked.Increment(ref _generation);
            var wanted = new HashSet<PdfPageViewModel>();
            foreach (var target in targets)
            {
                var page = target.Page;
                wanted.Add(page);
                if (page.IsBlankPage)
                    continue;

                var spec = ComputeBaseSpec(page, dpiScale);
                if (page.RenderedBaseKey != spec.Key && TryGetCached(model, page, spec.Key, out var cached))
                {
                    page.ImageSource = cached;
                    page.RenderedBaseKey = spec.Key;
                }

                // 줌·회전이 바뀌면 이전 상세 타일은 위치가 맞지 않는다
                if (page.RenderedDetailKey != null && page.RenderedDetailKey != DetailKey(spec))
                    page.ClearDetail();
                if (!spec.Capped && page.RenderedDetailKey != null)
                    page.ClearDetail();
            }

            // 멀리 벗어난 페이지는 그림을 내려놓아 메모리를 일정하게 (가상화가 재활용하는 컨테이너 포함)
            if (targets.Count > 0)
            {
                int minIndex = targets.Min(t => t.Page.PageIndex);
                int maxIndex = targets.Max(t => t.Page.PageIndex);
                foreach (var page in model.Pages)
                {
                    if (wanted.Contains(page) || (page.PageIndex >= minIndex - 4 && page.PageIndex <= maxIndex + 4))
                        continue;
                    if (page.ImageSource != null || page.DetailImageSource != null)
                    {
                        page.ImageSource = null;
                        page.RenderedBaseKey = null;
                        page.ClearDetail();
                    }
                }
            }

            _viewport = new Viewport(model, targets, dpiScale, generation);
            _signal.Set();
        }

        /// <summary>이 페이지도 미리 그려 둔다 (보이는 영역 계산과 별개).</summary>
        public void RequestPage(PdfDocumentModel model, PdfPageViewModel page)
        {
            _requested[page] = model;
            _signal.Set();
        }

        /// <summary>문서를 닫을 때: 진행 중 작업 취소, 캐시 정리.</summary>
        public void Forget(PdfDocumentModel model)
        {
            var viewport = _viewport;
            if (viewport != null && ReferenceEquals(viewport.Model, model))
                _viewport = null;
            Interlocked.Increment(ref _generation);
            foreach (var key in _requested.Where(kv => ReferenceEquals(kv.Value, model)).Select(kv => kv.Key).ToList())
                _requested.TryRemove(key, out _);
            lock (_cacheLock)
            {
                foreach (var key in _cache.Keys.Where(k => ReferenceEquals(k.Model, model)).ToList())
                    RemoveCacheEntry(key);
            }
        }

        public void ClearCache()
        {
            lock (_cacheLock)
            {
                _cache.Clear();
                _cacheOrder.Clear();
                _cacheBytes = 0;
            }
            _failed.Clear();
        }

        // ── 사양 계산 ────────────────────────────────────────────────────

        private readonly record struct BaseSpec(string Key, int FullWidth, int FullHeight, int Width, int Height, int Rotate, bool Capped);

        private static BaseSpec ComputeBaseSpec(PdfPageViewModel page, double dpiScale)
        {
            int fullW = Math.Max(1, (int)Math.Round(page.Width * dpiScale));
            int fullH = Math.Max(1, (int)Math.Round(page.Height * dpiScale));
            int rotate = PdfService.ToPdfiumRotate(page.Rotation);
            long pixels = (long)fullW * fullH;
            double scale = pixels > MaxBasePixels ? Math.Sqrt(MaxBasePixels / (double)pixels) : 1.0;
            int w = Math.Max(1, (int)Math.Round(fullW * scale));
            int h = Math.Max(1, (int)Math.Round(fullH * scale));
            return new BaseSpec($"{w}x{h}r{rotate}", fullW, fullH, w, h, rotate, scale < 1.0);
        }

        private static string DetailKey(BaseSpec spec) => $"{spec.FullWidth}x{spec.FullHeight}r{spec.Rotate}";

        /// <summary>보이는 부분(DIP) → 상세 타일 영역(전체 해상도 픽셀). 여유를 두고 격자에 맞춰 스크롤마다 다시 그리지 않게.</summary>
        private static (PixelRegion Needed, PixelRegion Render) ComputeDetailRegion(Rect visible, BaseSpec spec, double dpiScale)
        {
            int x0 = (int)Math.Floor(visible.Left * dpiScale);
            int y0 = (int)Math.Floor(visible.Top * dpiScale);
            int x1 = (int)Math.Ceiling(visible.Right * dpiScale);
            int y1 = (int)Math.Ceiling(visible.Bottom * dpiScale);
            x0 = Math.Clamp(x0, 0, spec.FullWidth);
            y0 = Math.Clamp(y0, 0, spec.FullHeight);
            x1 = Math.Clamp(x1, 0, spec.FullWidth);
            y1 = Math.Clamp(y1, 0, spec.FullHeight);
            var needed = new PixelRegion(x0, y0, Math.Max(0, x1 - x0), Math.Max(0, y1 - y0));

            int rx0 = Math.Max(0, (x0 - DetailMargin) / DetailGrid * DetailGrid);
            int ry0 = Math.Max(0, (y0 - DetailMargin) / DetailGrid * DetailGrid);
            int rx1 = Math.Min(spec.FullWidth, (x1 + DetailMargin + DetailGrid - 1) / DetailGrid * DetailGrid);
            int ry1 = Math.Min(spec.FullHeight, (y1 + DetailMargin + DetailGrid - 1) / DetailGrid * DetailGrid);
            return (needed, new PixelRegion(rx0, ry0, Math.Max(0, rx1 - rx0), Math.Max(0, ry1 - ry0)));
        }

        // ── 작업 고르기 ──────────────────────────────────────────────────

        private Job? NextJob(Viewport viewport)
        {
            var model = viewport.Model;
            if (model.IsDisposed || model.Document == null)
                return null;
            double dpi = viewport.DpiScale;

            // 1) 보이는 페이지의 전체 그림
            foreach (var target in viewport.Targets.Where(t => t.IsVisible))
            {
                var job = BaseJobIfNeeded(viewport, target.Page, dpi);
                if (job != null) return job;
            }

            // 2) 보이는 페이지의 상세 타일 (고배율)
            foreach (var target in viewport.Targets.Where(t => t.IsVisible))
            {
                var page = target.Page;
                if (page.IsBlankPage) continue;
                var spec = ComputeBaseSpec(page, dpi);
                if (!spec.Capped) continue;
                var (needed, render) = ComputeDetailRegion(target.VisibleRect, spec, dpi);
                if (needed.Width <= 0 || needed.Height <= 0 || render.Width <= 0 || render.Height <= 0) continue;
                string key = DetailKey(spec);
                if (page.RenderedDetailKey == key && page.RenderedDetailRegion.Contains(needed)) continue;
                if (_failed.ContainsKey((page, key + "@" + render))) continue;
                return new Job(JobKind.Detail, model, page, key, spec.FullWidth, spec.FullHeight, spec.Rotate, render, dpi, viewport.Generation);
            }

            // 3) 미리 그릴 페이지
            foreach (var target in viewport.Targets.Where(t => !t.IsVisible))
            {
                var job = BaseJobIfNeeded(viewport, target.Page, dpi);
                if (job != null) return job;
            }
            foreach (var (page, owner) in _requested.ToArray())
            {
                if (!ReferenceEquals(owner, model) || !model.Pages.Contains(page))
                {
                    _requested.TryRemove(page, out _);
                    continue;
                }
                var job = BaseJobIfNeeded(viewport, page, dpi);
                if (job != null) return job;
                _requested.TryRemove(page, out _);
            }

            // 4) 보이는 페이지의 주석 읽기
            foreach (var target in viewport.Targets)
            {
                if (!target.Page.AnnotationsLoaded && !target.Page.IsBlankPage)
                    return new Job(JobKind.Annotations, model, target.Page, "", 0, 0, 0, default, dpi, viewport.Generation);
            }

            return null;
        }

        private Job? BaseJobIfNeeded(Viewport viewport, PdfPageViewModel page, double dpi)
        {
            if (page.IsBlankPage || page.Width <= 0 || page.Height <= 0)
                return null;
            var spec = ComputeBaseSpec(page, dpi);
            if (page.RenderedBaseKey == spec.Key || _failed.ContainsKey((page, spec.Key)))
                return null;
            return new Job(JobKind.Base, viewport.Model, page, spec.Key, spec.Width, spec.Height, spec.Rotate,
                PixelRegion.Full(spec.Width, spec.Height), dpi, viewport.Generation);
        }

        /// <summary>작업 중에 보이는 영역이 바뀌었을 때, 이 작업이 아직 필요한지.</summary>
        private bool StillWanted(Job job)
        {
            var viewport = _viewport;
            if (viewport == null || !ReferenceEquals(viewport.Model, job.Model) || job.Model.IsDisposed)
                return false;
            if (Interlocked.Read(ref _generation) == job.Generation)
                return true;

            var target = viewport.Targets.FirstOrDefault(t => ReferenceEquals(t.Page, job.Page));
            if (target == null && !(job.Kind == JobKind.Base && _requested.ContainsKey(job.Page)))
                return false;

            var spec = ComputeBaseSpec(job.Page, viewport.DpiScale);
            if (job.Kind == JobKind.Base)
                return spec.Key == job.Key;
            if (job.Kind == JobKind.Detail)
            {
                if (target == null || !target.IsVisible || DetailKey(spec) != job.Key)
                    return false;
                var (needed, _) = ComputeDetailRegion(target.VisibleRect, spec, viewport.DpiScale);
                return job.Region.Contains(needed);
            }
            return true;
        }

        // ── 실행 ─────────────────────────────────────────────────────────

        private void WorkerLoop()
        {
            while (true)
            {
                _signal.WaitOne();
                while (true)
                {
                    var viewport = _viewport;
                    if (viewport == null)
                        break;
                    Job? job;
                    try
                    {
                        job = NextJob(viewport);
                    }
                    catch (Exception ex)
                    {
                        Log($"[ERROR] NextJob: {ex.Message}");
                        break;
                    }
                    if (job == null)
                        break;

                    _busy = true;
                    try
                    {
                        Execute(job);
                    }
                    catch (Exception ex)
                    {
                        Log($"[ERROR] {job.Kind} page {job.Page.PageIndex + 1}: {ex.Message}");
                        _failed[(job.Page, job.Kind == JobKind.Detail ? job.Key + "@" + job.Region : job.Key)] = 0;
                        if (job.Kind == JobKind.Annotations)
                            job.Page.AnnotationsLoaded = true;
                    }
                    finally
                    {
                        _busy = false;
                    }
                }
            }
        }

        private void Execute(Job job)
        {
            var document = job.Model.Document;
            if (document == null || job.Model.IsDisposed)
                return;

            if (job.Kind == JobKind.Annotations)
            {
                PdfService.LoadAnnotationsLazy(job.Model, job.Page);
                return;
            }

            var watch = Stopwatch.StartNew();
            bool cancelled = false;
            Func<bool> cancel = () =>
            {
                if (StillWanted(job))
                    return false;
                cancelled = true;
                return true;
            };

            var bitmap = PdfService.RenderBitmap(document, job.Page.OriginalPageIndex, job.FullWidth, job.FullHeight, job.Rotate,
                PdfService.ScreenRenderFlags, job.Region, postProcess: true, dpiScale: job.DpiScale, cancel: cancel);

            if (bitmap == null)
            {
                if (!cancelled && !job.Model.IsDisposed)
                {
                    Log($"[WARN] {job.Kind} render failed page {job.Page.PageIndex + 1} key={job.Key}");
                    _failed[(job.Page, job.Kind == JobKind.Detail ? job.Key + "@" + job.Region : job.Key)] = 0;
                }
                return;
            }

            var page = job.Page;
            if (job.Kind == JobKind.Base)
            {
                AddToCache(job.Model, page, job.Key, bitmap);
                page.RenderedBaseKey = job.Key;
                OnUi(() =>
                {
                    if (page.RenderedBaseKey == job.Key)
                    {
                        page.ImageSource = bitmap;
                        page.IsHighResRendered = true;
                    }
                });
            }
            else
            {
                page.RenderedDetailKey = job.Key;
                page.RenderedDetailRegion = job.Region;
                double dpi = job.DpiScale;
                var rect = new Rect(job.Region.X / dpi, job.Region.Y / dpi, job.Region.Width / dpi, job.Region.Height / dpi);
                OnUi(() =>
                {
                    if (page.RenderedDetailKey == job.Key && page.RenderedDetailRegion == job.Region)
                    {
                        page.DetailRect = rect;
                        page.DetailImageSource = bitmap;
                    }
                });
            }

            if (watch.ElapsedMilliseconds > 250)
                Log($"[RENDER] {job.Kind} page {page.PageIndex + 1} {job.Region.Width}x{job.Region.Height} {watch.ElapsedMilliseconds}ms");
        }

        private static void OnUi(Action action)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
                action();
            else
                dispatcher.BeginInvoke(action);
        }

        // ── 캐시 ─────────────────────────────────────────────────────────

        private void AddToCache(PdfDocumentModel model, PdfPageViewModel page, string key, BitmapSource bitmap)
        {
            long bytes = (long)bitmap.PixelWidth * bitmap.PixelHeight * 4;
            lock (_cacheLock)
            {
                var cacheKey = (model, page, key);
                if (_cache.ContainsKey(cacheKey))
                    RemoveCacheEntry(cacheKey);
                var node = _cacheOrder.AddFirst(new CacheEntry(model, page, key, bitmap, bytes));
                _cache[cacheKey] = node;
                _cacheBytes += bytes;
                while (_cacheBytes > CacheBudgetBytes && _cacheOrder.Last != null && _cacheOrder.Last != node)
                {
                    var last = _cacheOrder.Last.Value;
                    RemoveCacheEntry((last.Model, last.Page, last.Key));
                }
            }
        }

        private bool TryGetCached(PdfDocumentModel model, PdfPageViewModel page, string key, out BitmapSource bitmap)
        {
            lock (_cacheLock)
            {
                if (_cache.TryGetValue((model, page, key), out var node))
                {
                    _cacheOrder.Remove(node);
                    _cacheOrder.AddFirst(node);
                    bitmap = node.Value.Bitmap;
                    return true;
                }
            }
            bitmap = null!;
            return false;
        }

        private void RemoveCacheEntry((PdfDocumentModel, PdfPageViewModel, string) key)
        {
            if (_cache.Remove(key, out var node))
            {
                _cacheOrder.Remove(node);
                _cacheBytes -= node.Value.Bytes;
            }
        }

        private static void Log(string message)
        {
            try
            {
                DebugLog.Append("Render", message);
            }
            catch
            {
            }
        }
    }
}
