#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <d2d1.h>
#include <dwrite.h>
#include <dwrite_3.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <string>
#include <vector>
#include <deque>
#include <unordered_map>
#include <algorithm>
#include <memory>
#include <shlwapi.h>

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")

/**
 * ============================================================
 * miniterm.cpp - ConPTY WinPE Terminal
 * ============================================================
 *
 * BUILD:
 * cl miniterm.cpp miniterm.res /EHsc /std:c++17 /MT user32.lib kernel32.lib gdi32.lib d2d1.lib dwrite.lib ole32.lib shlwapi.lib comdlg32.lib comctl32.lib /link /SUBSYSTEM:WINDOWS /ENTRY:wmainCRTStartup /MANIFEST:EMBED /MANIFESTINPUT:miniterm.exe.manifest
 *
 * NOTE: /SUBSYSTEM:WINDOWS /ENTRY:wmainCRTStartup is required. Without it
 * the linker auto-detects CONSOLE subsystem (it sees wmain, not WinMain),
 * which allocates a console for the process -- on Windows 11 with Windows
 * Terminal set as the default terminal host, that console opens inside a
 * Windows Terminal window alongside miniterm's own window. wmainCRTStartup
 * keeps the wide-char wmain(argc, argv) entry point while /SUBSYSTEM:WINDOWS
 * suppresses the console allocation.
 *
 * VERSION HISTORY:
 * v1.4   - Stable Launch Fix
 * v1.5   - Resolver + Error Visibility
 * v1.5.1 - Case-insensitive argument parsing
 * v1.6   - Keyboard input + UTF-8 decoder + Nerd Font glyph support
 * v1.7   - Render thread safety + proper VT100/ANSI processor
 * v1.7.1 - WinPE font loading + null safety
 * v1.7.2 - Nerd Font glyph fix + exit handling + Tab completion
 * v1.8   - Dynamic sizing + menus + backspace fix + exit fix
 * v1.9   - Per-cell colour + Solarized themes + status bar + rendering fix
 * v2.0   - Scrollback buffer + text selection + colour fix (arch flawed)
 * v2.1   - Scrollback architecture fix + horizontal clip fix
 * v2.2   - Block cursor + status col/row + BEL + Alt fix + font picker
 * v2.3   - Double-width glyph support + pty_resize forward decl fix
 * v2.4   - Wide glyph clip fix + history bleed-through fix
 * v2.5   - Wide glyph DrawGlyphRun fix + ESC[G] tail-erase fix
 * v2.6   - OOM fix + glyph cache + wide glyph scale + wrap fix +
 *          right-click context menu
 * v2.7   - Debug mode + code cleanup + two-space indent
 *
 * v2.8   - Glyph rendering fix + fat debug mode
 *          - Root cause of missing glyphs identified: get_cached_layout
 *            was creating IDWriteTextLayout with maxWidth=CELL_W which
 *            truncates icon glyph ink at layout creation time, before
 *            any clip rect is applied. Fixed by creating narrow layouts
 *            with maxWidth=CELL_W*3 so ink is never truncated at source.
 *          - Narrow glyphs rendered with NO clip rect so ink can bleed
 *            into adjacent cell space (Windows Terminal approach).
 *            Adjacent cell background rects cover any bleed on next draw.
 *          - Wide glyph detection threshold lowered: now triggers at
 *            advanceWidth >= ref_advance (same width as M) for PUA
 *            range U+E000-U+F8FF and Nerd Font ranges, in addition to
 *            the existing 1.5x threshold for CJK etc. This correctly
 *            catches Nerd Font icons whose advance == CELL_W but whose
 *            ink overflows.
 *          - -debug flag now logs: font file probe results, font load
 *            success/failure, IDWriteFactory3 availability, fontface
 *            extraction, ref_advance, CELL_W/H, per-glyph advance/rsb/
 *            lsb/ink metrics, is_wide result, layout maxWidth used,
 *            actual DirectWrite rendered width per glyph, and frame
 *            render stats.
 *
 * ============================================================
 */

// =====================================================
// FORWARD DECLARATIONS
// =====================================================
static void pty_resize();
static void rebuild_fmt(const std::wstring& face, float size);
static void measure_cell_dims();

// =====================================================
// CELL
// =====================================================
struct Cell {
  uint32_t cp        = ' ';
  uint8_t  fg        = 7;
  uint8_t  bg        = 0;
  bool     wide_cont = false;
};

// =====================================================
// SCREEN + HISTORY — forward declared after Cell is complete
// =====================================================
static CRITICAL_SECTION               g_cs;
static std::vector<std::vector<Cell>> g_screen;
static std::deque<std::vector<Cell>>  g_history;
static std::vector<int>               g_row_max_cx;
static int  g_view_top  = 0;
static bool g_at_bottom = true;

// Forward declared here so clipboard helpers can use it before definition
static const std::vector<Cell>& get_view_row(int view_row);

// =====================================================
// DIMENSIONS
// =====================================================
static int   W                = 120;
static int   H                = 40;
static int   SCROLLBACK_LINES = 1000;
static float CELL_W           = 9.6f;
static float CELL_H           = 20.0f;

// Fixed tab-strip height in pixels. Queried dynamically it would need at
// least one tab item inserted first, but the first pty_resize() runs
// before any session/tab exists -- a fixed constant sidesteps that
// bootstrapping order problem. Matches native tab control height closely
// enough at default system font/DPI.
static const int kTabStripH = 24;

// =====================================================
// DEBUG
// =====================================================
static bool  g_debug      = false;
static FILE* g_debug_file = nullptr;

static void dbg(const char* fmt_str, ...)
{
  if (!g_debug || !g_debug_file) return;
  va_list args;
  va_start(args, fmt_str);
  vfprintf(g_debug_file, fmt_str, args);
  va_end(args);
  fflush(g_debug_file);
}

static void debug_open()
{
  if (!g_debug) return;
  // Write to a fixed easy path to avoid GetModuleFileName issues
  _wfopen_s(&g_debug_file, L"C:\\miniterm_debug.txt", L"w");
  if (!g_debug_file) {
    // Fallback: try current directory
    fopen_s(&g_debug_file, "miniterm_debug.txt", "w");
  }
  dbg("miniterm v2.8 debug log\n");
  dbg("======================\n");
}

// =====================================================
// MENU / CONTEXT MENU / TIMER IDs
// =====================================================
#define IDM_FILE_THEME_DARK      1002
#define IDM_FILE_THEME_LIGHT     1003
#define IDM_FILE_SAVELOG         1004
#define IDM_FILE_EXIT            1005
#define IDM_FILE_FONT            1006
#define IDM_HELP_ABOUT           1007
#define IDM_FILE_PREFS           1008
#define IDM_CTX_COPY             2001
#define IDM_CTX_PASTE            2002
#define IDM_CTX_CLEAR_SCROLLBACK 2003
#define IDT_SIZE_POLL            3001
#define IDT_CURSOR_BLINK         3002
#define IDC_PREF_THEME           4001
#define IDC_PREF_FONT_NAME       4002
#define IDC_PREF_FONT_SIZE       4003
#define IDC_PREF_FONT_BROWSE     4004
#define IDC_PREF_SCROLLBACK      4005
#define IDM_SHELL_BASE           5000
#define IDM_SHELL_MAX            32
#define IDM_NEWTAB_BASE          9500
#define IDC_PREF_DEFAULT_SHELL   4006
#define IDC_PREF_OK              4007
#define IDC_PREF_CANCEL          4008
#define IDM_TAB_CLOSE            6001
#define IDM_TAB_RENAME           6002
#define WM_APP_SESSION_ENDED     (WM_APP + 1)
#define IDC_TABCTRL              7001
#define IDC_RENAME_EDIT          8001
#define IDC_RENAME_OK            8002
#define IDC_RENAME_CANCEL        8003
#define IDC_ABOUT_TEXT           8101
#define IDC_ABOUT_OK             8102

// =====================================================
// THEME + PALETTE
// =====================================================
enum class Theme { SolarizedDark, SolarizedLight };
static Theme        g_theme = Theme::SolarizedDark;
static D2D1_COLOR_F g_palette[256];

static const D2D1_COLOR_F k_sol_dark[16] = {
  { 0.027f, 0.212f, 0.259f, 1.0f },
  { 0.863f, 0.196f, 0.184f, 1.0f },
  { 0.522f, 0.600f, 0.000f, 1.0f },
  { 0.718f, 0.545f, 0.000f, 1.0f },
  { 0.149f, 0.545f, 0.824f, 1.0f },
  { 0.827f, 0.212f, 0.510f, 1.0f },
  { 0.165f, 0.631f, 0.596f, 1.0f },
  { 0.933f, 0.910f, 0.835f, 1.0f },
  { 0.345f, 0.431f, 0.459f, 1.0f },
  { 0.980f, 0.380f, 0.310f, 1.0f },
  { 0.670f, 0.750f, 0.100f, 1.0f },
  { 0.910f, 0.710f, 0.100f, 1.0f },
  { 0.424f, 0.443f, 0.769f, 1.0f },
  { 0.910f, 0.420f, 0.650f, 1.0f },
  { 0.350f, 0.780f, 0.740f, 1.0f },
  { 0.992f, 0.965f, 0.890f, 1.0f },
};

static const D2D1_COLOR_F k_sol_light[16] = {
  { 0.933f, 0.910f, 0.835f, 1.0f },
  { 0.863f, 0.196f, 0.184f, 1.0f },
  { 0.522f, 0.600f, 0.000f, 1.0f },
  { 0.600f, 0.440f, 0.000f, 1.0f },
  { 0.149f, 0.545f, 0.824f, 1.0f },
  { 0.827f, 0.212f, 0.510f, 1.0f },
  { 0.000f, 0.490f, 0.537f, 1.0f },
  { 0.027f, 0.212f, 0.259f, 1.0f },
  { 0.576f, 0.631f, 0.631f, 1.0f },
  { 0.750f, 0.120f, 0.100f, 1.0f },
  { 0.380f, 0.460f, 0.000f, 1.0f },
  { 0.500f, 0.360f, 0.000f, 1.0f },
  { 0.100f, 0.380f, 0.620f, 1.0f },
  { 0.620f, 0.130f, 0.380f, 1.0f },
  { 0.000f, 0.360f, 0.400f, 1.0f },
  { 0.000f, 0.169f, 0.212f, 1.0f },
};

static void build_palette(Theme t)
{
  const D2D1_COLOR_F* base =
    (t == Theme::SolarizedDark) ? k_sol_dark : k_sol_light;
  for (int i = 0; i < 16; i++) g_palette[i] = base[i];
  for (int i = 16; i < 232; i++) {
    int idx = i - 16;
    int b = idx % 6, g_ = (idx / 6) % 6, r = idx / 36;
    auto cv = [](int v) -> float {
      return v == 0 ? 0.0f : (55.0f + v * 40.0f) / 255.0f;
    };
    g_palette[i] = { cv(r), cv(g_), cv(b), 1.0f };
  }
  for (int i = 232; i < 256; i++) {
    float v = (8.0f + (i - 232) * 10.0f) / 255.0f;
    g_palette[i] = { v, v, v, 1.0f };
  }
}

static uint8_t theme_bg() { return 0; }
static uint8_t theme_fg() { return 7; }

// =====================================================
// CURSOR + SGR STATE
// =====================================================
static int     cx               = 0;
static int     cy               = 0;
static int     saved_cx         = 0;
static int     saved_cy         = 0;
static bool    g_cursor_visible = true;
static uint8_t g_cur_fg         = 7;
static uint8_t g_cur_bg         = 0;

// =====================================================
// SELECTION
// =====================================================
static bool g_sel_active   = false;
static bool g_sel_dragging = false;
static int  g_sel_r0 = 0, g_sel_c0 = 0;
static int  g_sel_r1 = 0, g_sel_c1 = 0;

// =====================================================
// CONPTY
// =====================================================
static HPCON   g_hpc  = nullptr;
static HANDLE  hInR   = nullptr;
static HANDLE  hInW   = nullptr;
static HANDLE  hOutR  = nullptr;
static HANDLE  hOutW  = nullptr;

static PROCESS_INFORMATION g_child_pi        = {};
static HANDLE               g_reader_thread   = nullptr;
static volatile bool         g_shell_switching = false;

// Set (from the main thread's 500ms poll timer) when the active
// session's child process is found to have exited but reader()'s
// ReadFile hadn't noticed on its own -- ConPTY doesn't reliably signal
// EOF on the output pipe just because the hosted process exited; the
// client is expected to notice via the process handle and act.
// CancelIoEx unblocks the pending ReadFile so reader() can act on it;
// this flag lets reader() tell that apart from an ordinary pause/switch
// cancellation, which also uses CancelIoEx but must NOT be reported as
// a session ending.
static volatile bool g_watcher_signals_exit = false;

// =====================================================
// SESSIONS (tabs)
//
// Only the ACTIVE session's reader thread ever runs. The bare globals
// above (g_screen, cx, cy, g_hpc, hInW, hOutR, g_exe, ...) always mean
// "whichever session is active right now" -- render(), VtProcessor,
// screen_clear_region() etc. all keep operating on them completely
// unchanged. Switching tabs pauses the outgoing session's reader via
// CancelIoEx (its handles stay open, NOT closed) and captures the live
// globals into its Session struct; switching back restores them and
// resumes reading. A backgrounded shell keeps running -- its output
// just queues in the OS pipe buffer until the tab is revisited.
// =====================================================
struct Session {
  std::wstring name, shellname, exe, args;

  std::vector<std::vector<Cell>> screen;
  std::deque<std::vector<Cell>>  history;
  std::vector<int>               row_max_cx;
  int  view_top  = 0;
  bool at_bottom = true;

  int     cx = 0, cy = 0, saved_cx = 0, saved_cy = 0;
  bool    cursor_visible = true;
  uint8_t cur_fg = 7, cur_bg = 0;

  bool sel_active = false, sel_dragging = false;
  int  sel_r0 = 0, sel_c0 = 0, sel_r1 = 0, sel_c1 = 0;

  HPCON               hpc      = nullptr;
  HANDLE               hInR = nullptr, hInW = nullptr, hOutR = nullptr, hOutW = nullptr;
  PROCESS_INFORMATION  child_pi = {};
  HANDLE               reader_thread = nullptr;
};

static std::vector<std::unique_ptr<Session>> g_sessions;
static int  g_active_session   = -1;
static int  g_next_session_num = 1;

// =====================================================
// DIRECTWRITE
// =====================================================
static IDWriteFactory*        dw          = nullptr;
static IDWriteTextFormat*     fmt         = nullptr;
static IDWriteFontFace*       g_fontface  = nullptr;
static IDWriteFontCollection* g_fontcol   = nullptr;
static ID2D1Factory*          d2d         = nullptr;
static ID2D1HwndRenderTarget* rt          = nullptr;
static std::wstring           g_font_name = L"";
static float                  g_font_size = 16.0f;
static UINT32                 g_ref_advance = 0;
static UINT32                 g_design_units_per_em = 2048;

// Cache lock covers wide_cache, wide_width_cache, layout_cache
static CRITICAL_SECTION                      g_cache_cs;
static std::unordered_map<uint32_t, bool>    g_wide_cache;
static std::unordered_map<uint32_t, float>   g_wide_width_cache;
static std::unordered_map<uint32_t,
                          IDWriteTextLayout*> g_layout_cache;
static const size_t k_layout_cache_max = 8192;

// =====================================================
// WINDOW HANDLES
// =====================================================
static HWND hwnd        = nullptr;
static HWND g_statusbar = nullptr;
static HWND g_scrollbar = nullptr;
static HWND g_tabctrl   = nullptr;

// =====================================================
// SHELL CONFIG
// =====================================================
static std::wstring g_exe       = L"pwsh.exe";
static std::wstring g_args      = L"-NoLogo -NoProfile";
static std::wstring g_shellname = L"pwsh.exe";

// =====================================================
// LAST KNOWN CLIENT SIZE
// =====================================================
static int g_last_client_w = 0;
static int g_last_client_h = 0;

// =====================================================
// FALLBACK SEARCH DIRS
// =====================================================
static const wchar_t* k_fallback_dirs[] = {
  L"X:\\Tools\\pwsh\\",
  nullptr
};

// =====================================================
// DESIGN-UNITS TO DIPS CONVERSION
// =====================================================
static float du_to_dip(UINT32 du)
{
  if (g_design_units_per_em == 0) return 0.0f;
  return (float)du / (float)g_design_units_per_em
       * g_font_size * 96.0f / 72.0f;
}

// =====================================================
// IS NERD FONT / PUA RANGE
// Returns true for codepoints in known Nerd Font icon
// ranges that should be treated as potentially wide.
// =====================================================
static bool is_nerd_font_range(uint32_t cp)
{
  // Unicode Private Use Area
  if (cp >= 0xE000  && cp <= 0xF8FF)  return true;
  // Supplementary PUA-A
  if (cp >= 0xF0000 && cp <= 0xFFFFF) return true;
  // Nerd Font specific ranges commonly used
  if (cp >= 0x1F300 && cp <= 0x1FAFF) return true; // emoji/symbols
  return false;
}

// =====================================================
// WIDE GLYPH DETECTION
// A glyph is "wide" if:
//   a) Its advance >= 1.5x the reference ('M') advance, OR
//   b) It is in a Nerd Font PUA range AND its ink extends
//      beyond CELL_W (negative rightSideBearing)
// Also stores the required render width in g_wide_width_cache.
// =====================================================
static bool is_wide_glyph(uint32_t cp)
{
  if (cp < 0x80) return false;
  if (cp == 0x20 || cp == 0xA0) return false;

  EnterCriticalSection(&g_cache_cs);
  auto it = g_wide_cache.find(cp);
  if (it != g_wide_cache.end()) {
    bool r = it->second;
    LeaveCriticalSection(&g_cache_cs);
    return r;
  }
  LeaveCriticalSection(&g_cache_cs);

  bool  wide         = false;
  float needed_w     = CELL_W * 2.0f;
  bool  in_nf_range  = is_nerd_font_range(cp);

  if (g_fontface && g_ref_advance > 0) {
    UINT16 gi  = 0;
    UINT32 cpt = cp;
    HRESULT hr = g_fontface->GetGlyphIndices(&cpt, 1, &gi);

    if (SUCCEEDED(hr) && gi != 0) {
      DWRITE_GLYPH_METRICS gm = {};
      hr = g_fontface->GetDesignGlyphMetrics(&gi, 1, &gm, FALSE);

      if (SUCCEEDED(hr)) {
        float adv_dip = du_to_dip(gm.advanceWidth);
        INT32 rsb     = gm.rightSideBearing; // signed; negative = ink overflows
        INT32 lsb     = gm.leftSideBearing;
        float rsb_dip = du_to_dip((UINT32)(rsb < 0 ? (UINT32)(-rsb) : 0));

        // Standard wide detection: advance >= 1.5x reference
        bool wide_by_advance = (gm.advanceWidth >= g_ref_advance * 3 / 2);

        // Nerd Font detection: in PUA range and has negative rsb
        // (ink bleeds past advance box)
        bool wide_by_ink = in_nf_range && (rsb < 0);

        wide = wide_by_advance || wide_by_ink;

        if (wide) {
          // Calculate total required width including ink overhang
          needed_w = adv_dip + (rsb < 0 ? rsb_dip : 0.0f);
          if (needed_w < CELL_W * 2.0f) needed_w = CELL_W * 2.0f;
        }

        dbg("GLYPH U+%04X gi=%u adv_du=%u adv_dip=%.3f "
            "rsb=%d lsb=%d rsb_dip=%.3f "
            "wide_by_adv=%d wide_by_ink=%d WIDE=%d "
            "in_nf=%d needed_w=%.3f CELL_W=%.3f\n",
            cp, gi,
            gm.advanceWidth, adv_dip,
            rsb, lsb, rsb_dip,
            (int)wide_by_advance, (int)wide_by_ink, (int)wide,
            (int)in_nf_range, needed_w, CELL_W);
      } else {
        dbg("GLYPH U+%04X gi=%u GetDesignGlyphMetrics FAILED hr=0x%08X\n",
            cp, gi, (unsigned)hr);
      }
    } else {
      dbg("GLYPH U+%04X GetGlyphIndices FAILED or gi=0 hr=0x%08X gi=%u\n",
          cp, (unsigned)hr, gi);
    }
  } else {
    dbg("GLYPH U+%04X SKIPPED: fontface=%p ref_advance=%u\n",
        cp, (void*)g_fontface, g_ref_advance);
  }

  EnterCriticalSection(&g_cache_cs);
  g_wide_cache[cp]       = wide;
  g_wide_width_cache[cp] = needed_w;
  LeaveCriticalSection(&g_cache_cs);
  return wide;
}

static float wide_glyph_w(uint32_t cp)
{
  EnterCriticalSection(&g_cache_cs);
  auto it = g_wide_width_cache.find(cp);
  float w = (it != g_wide_width_cache.end()) ? it->second : CELL_W * 2.0f;
  LeaveCriticalSection(&g_cache_cs);
  return w;
}

// =====================================================
// GLYPH LAYOUT CACHE
// CRITICAL FIX: maxWidth must be large enough to not
// truncate icon glyph ink at layout creation time.
// We use CELL_W * 3 — wider than any expected glyph —
// so DirectWrite never clips ink before we even render.
// The clip rect in render() controls actual visible area.
// =====================================================
static IDWriteTextLayout* get_cached_layout(uint32_t cp)
{
  EnterCriticalSection(&g_cache_cs);
  auto it = g_layout_cache.find(cp);
  if (it != g_layout_cache.end()) {
    IDWriteTextLayout* layout = it->second;
    LeaveCriticalSection(&g_cache_cs);
    return layout;
  }
  LeaveCriticalSection(&g_cache_cs);

  wchar_t buf[3] = {};
  int     len    = 0;
  if (cp <= 0xFFFF) {
    buf[0] = (wchar_t)cp; len = 1;
  } else {
    uint32_t sv = cp - 0x10000;
    buf[0] = (wchar_t)(0xD800 | (sv >> 10));
    buf[1] = (wchar_t)(0xDC00 | (sv & 0x3FF));
    len    = 2;
  }

  if (!dw || !fmt) return nullptr;

  // maxWidth = CELL_W * 3: generous enough to never truncate
  // any glyph ink, including wide Nerd Font icons.
  // The render loop does NOT clip narrow glyphs, so this
  // is the only constraint on ink width.
  IDWriteTextLayout* layout = nullptr;
  HRESULT hr = dw->CreateTextLayout(
    buf, (UINT32)len, fmt,
    CELL_W * 3.0f, CELL_H,
    &layout);

  if (FAILED(hr) || !layout) {
    dbg("LAYOUT_CACHE U+%04X CreateTextLayout FAILED hr=0x%08X\n",
        cp, (unsigned)hr);
    return nullptr;
  }

  layout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
  layout->SetLineSpacing(
    DWRITE_LINE_SPACING_METHOD_UNIFORM, CELL_H, CELL_H * 0.8f);

  if (g_debug) {
    DWRITE_TEXT_METRICS tm = {};
    layout->GetMetrics(&tm);
    dbg("LAYOUT_CACHE U+%04X maxW=%.2f actual_w=%.2f trailing_w=%.2f\n",
        cp, CELL_W * 3.0f, tm.width,
        tm.widthIncludingTrailingWhitespace);
  }

  EnterCriticalSection(&g_cache_cs);
  if (g_layout_cache.size() >= k_layout_cache_max) {
    dbg("LAYOUT_CACHE evicting %zu entries\n", g_layout_cache.size());
    for (auto& kv : g_layout_cache) if (kv.second) kv.second->Release();
    g_layout_cache.clear();
  }
  g_layout_cache[cp] = layout;
  LeaveCriticalSection(&g_cache_cs);
  return layout;
}

static void clear_layout_cache()
{
  EnterCriticalSection(&g_cache_cs);
  dbg("clear_layout_cache: releasing %zu layouts\n", g_layout_cache.size());
  for (auto& kv : g_layout_cache) if (kv.second) kv.second->Release();
  g_layout_cache.clear();
  g_wide_cache.clear();
  g_wide_width_cache.clear();
  LeaveCriticalSection(&g_cache_cs);
}

// =====================================================
// SCREEN HELPERS
// =====================================================
static void screen_clear_region(int x0, int y0, int x1, int y1)
{
  if (y0 < 0) y0 = 0; if (y1 >= H) y1 = H - 1;
  if (x0 < 0) x0 = 0; if (x1 >= W) x1 = W - 1;
  for (int y = y0; y <= y1; y++) {
    for (int x = x0; x <= x1; x++)
      g_screen[y][x] = Cell{ ' ', g_cur_fg, theme_bg(), false };
    g_row_max_cx[y] = (x0 == 0) ? 0 : min(g_row_max_cx[y], x0);
  }
}

static void screen_scroll_up()
{
  if ((int)g_history.size() >= SCROLLBACK_LINES)
    g_history.pop_front();
  g_history.push_back(g_screen[0]);
  for (int y = 1; y < H; y++) {
    g_screen[y-1]     = g_screen[y];
    g_row_max_cx[y-1] = g_row_max_cx[y];
  }
  g_screen[H-1].assign(W, Cell{ ' ', g_cur_fg, theme_bg(), false });
  g_row_max_cx[H-1] = 0;
  if (g_at_bottom) g_view_top = (int)g_history.size();
}

static void erase_line_tail()
{
  if (cy < H && g_row_max_cx[cy] > cx)
    for (int x = cx; x <= g_row_max_cx[cy] && x < W; x++)
      g_screen[cy][x] = Cell{ ' ', g_cur_fg, theme_bg(), false };
  // Only shrink the tracked content bound to what's actually left of cx —
  // forcing it to 0 unconditionally lost track of untouched content to the
  // left of the cursor, causing later same-row redraws to under-clear and
  // leave stale glyphs behind.
  if (cy < H) g_row_max_cx[cy] = (cx == 0) ? 0 : min(g_row_max_cx[cy], cx - 1);
}

// =====================================================
// SCROLLBAR
// =====================================================
static void update_scrollbar()
{
  if (!g_scrollbar) return;
  int hist = (int)g_history.size(), total = hist + H;
  if (total <= H) { ShowScrollBar(g_scrollbar, SB_CTL, FALSE); return; }
  ShowScrollBar(g_scrollbar, SB_CTL, TRUE);
  SCROLLINFO si = {};
  si.cbSize = sizeof(si);
  si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
  si.nMin   = 0; si.nMax = total - 1;
  si.nPage  = (UINT)H; si.nPos = g_view_top;
  SetScrollInfo(g_scrollbar, SB_CTL, &si, TRUE);
}

// =====================================================
// STATUS BAR
// =====================================================
static void update_statusbar()
{
  if (!g_statusbar) return;
  const wchar_t* tn =
    (g_theme == Theme::SolarizedDark) ? L"Solarized Dark" : L"Solarized Light";

  RECT sb = {};
  GetClientRect(g_statusbar, &sb);
  int w = sb.right - sb.left;
  if (w <= 0) w = 800;
  int parts[4] = { w * 35 / 100, w * 60 / 100, w * 82 / 100, -1 };
  SendMessageW(g_statusbar, SB_SETPARTS, 4, (LPARAM)parts);

  const wchar_t* session_name =
    (g_active_session >= 0 && g_active_session < (int)g_sessions.size())
    ? g_sessions[g_active_session]->name.c_str() : L"";

  wchar_t buf[256];
  wsprintfW(buf, L"  %s", session_name);
  SendMessageW(g_statusbar, SB_SETTEXTW, 0, (LPARAM)buf);
  wsprintfW(buf, L"Shell: %s", g_shellname.c_str());
  SendMessageW(g_statusbar, SB_SETTEXTW, 1, (LPARAM)buf);
  wsprintfW(buf, L"Theme: %s", tn);
  SendMessageW(g_statusbar, SB_SETTEXTW, 2, (LPARAM)buf);
  wsprintfW(buf, L"Ln %d  Col %d", cy + 1, cx + 1);
  SendMessageW(g_statusbar, SB_SETTEXTW, 3, (LPARAM)buf);
}

// =====================================================
// GRID RESIZE
// =====================================================
static void grid_resize(int newW, int newH)
{
  W = newW; H = newH;
  g_screen.assign(H, std::vector<Cell>(W, Cell{ ' ', theme_fg(), theme_bg(), false }));
  g_row_max_cx.assign(H, 0);
  for (auto& row : g_history)
    row.resize(W, Cell{ ' ', theme_fg(), theme_bg(), false });
  if (cx >= W) cx = W - 1;
  if (cy >= H) cy = H - 1;
  if (cx < 0)  cx = 0;
  if (cy < 0)  cy = 0;
}

// =====================================================
// UTF-8 DECODER
// =====================================================
struct Utf8Decoder {
  uint32_t codepoint = 0;
  int      remaining = 0;

  bool feed(uint8_t byte, uint32_t& cp) {
    if (remaining > 0) {
      if ((byte & 0xC0) == 0x80) {
        codepoint = (codepoint << 6) | (byte & 0x3F);
        remaining--;
        if (remaining == 0) { cp = codepoint; return true; }
        return false;
      }
      remaining = 0;
    }
    if ((byte & 0x80) == 0)    { cp = byte;               return true;  }
    if ((byte & 0xE0) == 0xC0) { codepoint = byte & 0x1F; remaining = 1; return false; }
    if ((byte & 0xF0) == 0xE0) { codepoint = byte & 0x0F; remaining = 2; return false; }
    if ((byte & 0xF8) == 0xF0) { codepoint = byte & 0x07; remaining = 3; return false; }
    return false;
  }
};

// =====================================================
// VT/ANSI PROCESSOR
// =====================================================
struct VtProcessor {
  enum class State { Normal, Esc, Csi, Osc, OscEsc };
  State state = State::Normal;

  static const int MAX_PARAMS = 32;
  int  params[MAX_PARAMS];
  int  nparams;
  bool private_mode;

  VtProcessor() { reset_params(); }

  void reset_params() {
    for (int i = 0; i < MAX_PARAMS; i++) params[i] = -1;
    nparams = 0; private_mode = false;
  }

  int p(int n, int def) const {
    return (n >= nparams || params[n] < 0) ? def : params[n];
  }

  static void clamp_cursor() {
    if (cx < 0) cx = 0; if (cx >= W) cx = W - 1;
    if (cy < 0) cy = 0; if (cy >= H) cy = H - 1;
  }

  void handle_sgr() {
    if (nparams == 0) { g_cur_fg = theme_fg(); g_cur_bg = theme_bg(); return; }
    int i = 0;
    while (i < nparams) {
      int v = p(i, 0);
      switch (v) {
        case 0:  g_cur_fg = theme_fg(); g_cur_bg = theme_bg(); break;
        case 1: case 22: break;
        case 39: g_cur_fg = theme_fg(); break;
        case 49: g_cur_bg = theme_bg(); break;
        case 30: case 31: case 32: case 33:
        case 34: case 35: case 36: case 37:
          g_cur_fg = (uint8_t)(v - 30); break;
        case 40: case 41: case 42: case 43:
        case 44: case 45: case 46: case 47:
          g_cur_bg = (uint8_t)(v - 40); break;
        case 90: case 91: case 92: case 93:
        case 94: case 95: case 96: case 97:
          g_cur_fg = (uint8_t)(v - 90 + 8); break;
        case 100: case 101: case 102: case 103:
        case 104: case 105: case 106: case 107:
          g_cur_bg = (uint8_t)(v - 100 + 8); break;
        case 38:
          if (p(i+1,-1) == 5 && i+2 < nparams)
          { g_cur_fg = (uint8_t)p(i+2, theme_fg()); i += 2; }
          break;
        case 48:
          if (p(i+1,-1) == 5 && i+2 < nparams)
          { g_cur_bg = (uint8_t)p(i+2, theme_bg()); i += 2; }
          break;
        default: break;
      }
      i++;
    }
  }

  void dispatch_csi(wchar_t cmd) {
    switch (cmd) {
      case L'A': cy -= p(0,1); break;
      case L'B': cy += p(0,1); break;
      case L'C': cx += p(0,1); break;
      case L'D': cx -= p(0,1); break;
      case L'E': cy += p(0,1); cx = 0; break;
      case L'F': cy -= p(0,1); cx = 0; break;
      case L'G': {
        int new_cx = p(0,1) - 1;
        if (new_cx < 0) new_cx = 0;
        if (new_cx == 0) erase_line_tail();
        cx = new_cx;
        break;
      }
      case L'H': case L'f':
        cy = p(0,1) - 1; cx = p(1,1) - 1; break;
      case L'J':
        switch (p(0,0)) {
          case 0:
            screen_clear_region(cx, cy, W-1, cy);
            screen_clear_region(0, cy+1, W-1, H-1); break;
          case 1:
            screen_clear_region(0, 0, W-1, cy-1);
            screen_clear_region(0, cy, cx, cy); break;
          case 2: case 3:
            screen_clear_region(0, 0, W-1, H-1); break;
        }
        break;
      case L'K':
        switch (p(0,0)) {
          case 0: screen_clear_region(cx, cy, W-1, cy); break;
          case 1: screen_clear_region(0,  cy, cx,  cy); break;
          case 2: screen_clear_region(0,  cy, W-1, cy); break;
        }
        break;
      case L's': saved_cx = cx; saved_cy = cy; break;
      case L'u': cx = saved_cx; cy = saved_cy; break;
      case L'm': handle_sgr(); break;
      case L'n':
        if (p(0,0) == 6) {
          char reply[32];
          wsprintfA(reply, "\x1b[%d;%dR", cy+1, cx+1);
          DWORD w2;
          WriteFile(hInW, reply, (DWORD)strlen(reply), &w2, NULL);
        }
        break;
      default: break;
    }
    clamp_cursor();
  }

  void feed(uint32_t cp) {
    switch (state) {
      case State::Normal:
        if (cp == 0x07) { Beep(800, 100); return; }
        if (cp == 0x1B) { state = State::Esc; return; }
        if (cp == '\r') { erase_line_tail(); cx = 0; return; }
        if (cp == '\n') {
          cy++;
          if (cy >= H) { screen_scroll_up(); cy = H-1; }
          return;
        }
        if (cp == '\b') { if (cx > 0) cx--; return; }
        if (cp == '\t') {
          int next = (cx + 8) & ~7;
          if (next >= W) next = W-1;
          while (cx < next)
            g_screen[cy][cx++] = { ' ', g_cur_fg, g_cur_bg, false };
          return;
        }
        if (cp < 0x20) return;
        {
          bool wide = is_wide_glyph(cp);
          if (wide && cx+1 < W) {
            g_screen[cy][cx]   = { cp, g_cur_fg, g_cur_bg, false };
            g_screen[cy][cx+1] = { 0,  g_cur_fg, g_cur_bg, true  };
            if (cy < H) g_row_max_cx[cy] = max(g_row_max_cx[cy], cx+1);
            cx += 2;
          } else if (wide && cx+1 >= W) {
            cy++;
            if (cy >= H) { screen_scroll_up(); cy = H-1; }
            cx = 0;
            g_screen[cy][0] = { cp, g_cur_fg, g_cur_bg, false };
            g_screen[cy][1] = { 0,  g_cur_fg, g_cur_bg, true  };
            if (cy < H) g_row_max_cx[cy] = max(g_row_max_cx[cy], 1);
            cx = 2;
          } else {
            g_screen[cy][cx] = { cp, g_cur_fg, g_cur_bg, false };
            if (cy < H) g_row_max_cx[cy] = max(g_row_max_cx[cy], cx);
            cx++;
          }
          if (cx >= W) {
            cx = 0; cy++;
            if (cy >= H) { screen_scroll_up(); cy = H-1; }
          }
        }
        return;

      case State::Esc:
        switch (cp) {
          case L'[': reset_params(); state = State::Csi; return;
          case L']': state = State::Osc; return;
          case L'M':
            if (cy > 0) cy--;
            else {
              for (int y = H-1; y > 0; y--) {
                g_screen[y]     = g_screen[y-1];
                g_row_max_cx[y] = g_row_max_cx[y-1];
              }
              g_screen[0].assign(W, Cell{ ' ', g_cur_fg, theme_bg(), false });
              g_row_max_cx[0] = 0;
            }
            state = State::Normal; return;
          case L'7': saved_cx=cx; saved_cy=cy; state=State::Normal; return;
          case L'8': cx=saved_cx; cy=saved_cy; state=State::Normal; return;
          default:   state = State::Normal; return;
        }

      case State::Csi:
        if (cp == L'?' && nparams == 0) { private_mode = true; return; }
        if (cp >= L'0' && cp <= L'9') {
          if (nparams == 0) nparams = 1;
          int& cur = params[nparams-1];
          if (cur < 0) cur = 0;
          cur = cur * 10 + (int)(cp - L'0');
          return;
        }
        if (cp == L';') { if (nparams < MAX_PARAMS) nparams++; return; }
        if (cp >= 0x40 && cp <= 0x7E) {
          if (!private_mode) dispatch_csi((wchar_t)cp);
          state = State::Normal; return;
        }
        state = State::Normal; return;

      case State::Osc:
        if (cp == 0x07) { state = State::Normal; return; }
        if (cp == 0x1B) { state = State::OscEsc; return; }
        return;

      case State::OscEsc:
        state = State::Normal; return;
    }
  }
};

// =====================================================
// EXE RESOLVER
// =====================================================
static std::wstring resolve_exe(const std::wstring& name)
{
  if (name.find(L'\\') != std::wstring::npos ||
      name.find(L'/')  != std::wstring::npos)
    if (GetFileAttributesW(name.c_str()) != INVALID_FILE_ATTRIBUTES)
      return name;

  std::wstring bare = name;
  {
    size_t sl = bare.rfind(L'\\');
    if (sl != std::wstring::npos) bare = bare.substr(sl + 1);
    size_t fs = bare.rfind(L'/');
    if (fs != std::wstring::npos) bare = bare.substr(fs + 1);
  }

  auto probe = [](const std::wstring& dir,
                  const std::wstring& file) -> std::wstring {
    std::wstring c = dir;
    if (!c.empty() && c.back() != L'\\') c += L'\\';
    c += file;
    return (GetFileAttributesW(c.c_str()) != INVALID_FILE_ATTRIBUTES) ? c : L"";
  };

  { wchar_t cwd[MAX_PATH];
    if (GetCurrentDirectoryW(MAX_PATH, cwd)) {
      auto r = probe(cwd, bare); if (!r.empty()) return r;
    } }
  { wchar_t pe[32768];
    if (GetEnvironmentVariableW(L"PATH", pe, 32768)) {
      wchar_t* ctx = nullptr, *tok = wcstok_s(pe, L";", &ctx);
      while (tok) {
        auto r = probe(tok, bare); if (!r.empty()) return r;
        tok = wcstok_s(nullptr, L";", &ctx);
      }
    } }
  for (int i = 0; k_fallback_dirs[i]; i++) {
    auto r = probe(k_fallback_dirs[i], bare); if (!r.empty()) return r;
  }
  return L"";
}

// =====================================================
// CLI PARSER
// =====================================================
static void parse_args(int argc, wchar_t** argv)
{
  for (int i = 1; i < argc; i++) {
    std::wstring a = argv[i];
    std::transform(a.begin(), a.end(), a.begin(), ::towlower);

    if (a == L"-useshell" && i+1 < argc) {
      std::wstring full = argv[i+1];
      size_t sp = full.find(L' ');
      if (sp == std::wstring::npos) { g_exe = full; g_args.clear(); }
      else { g_exe = full.substr(0,sp); g_args = full.substr(sp+1); }
      std::wstring shell = g_exe;
      size_t sl = shell.rfind(L'\\');
      if (sl != std::wstring::npos) shell = shell.substr(sl+1);
      g_shellname = shell;
      i++;
    } else if (a == L"-scrollback" && i+1 < argc) {
      int n = _wtoi(argv[i+1]);
      if (n >= 100) SCROLLBACK_LINES = n;
      i++;
    } else if (a == L"-debug") {
      g_debug = true;
    }
  }
}

// =====================================================
// MINIMAL JSON — scoped exactly to miniterm.json's needs.
// Read side is a generic value tree (robust to hand-edits and
// whitespace/ordering differences); write side is schema-specific
// (we always control the shape we emit, so no generic serializer
// is needed). No external dependency — single translation unit.
// =====================================================
struct JsonValue {
  enum class Type { Null, Bool, Number, String, Array, Object } type = Type::Null;
  bool                                             b   = false;
  double                                           num = 0.0;
  std::wstring                                     str;
  std::vector<JsonValue>                           arr;
  std::vector<std::pair<std::wstring, JsonValue>>  obj;

  const JsonValue* find(const wchar_t* key) const {
    if (type != Type::Object) return nullptr;
    for (auto& kv : obj) if (kv.first == key) return &kv.second;
    return nullptr;
  }
  std::wstring get_str(const wchar_t* key, const std::wstring& def) const {
    const JsonValue* v = find(key);
    return (v && v->type == Type::String) ? v->str : def;
  }
  double get_num(const wchar_t* key, double def) const {
    const JsonValue* v = find(key);
    return (v && v->type == Type::Number) ? v->num : def;
  }
  int get_int(const wchar_t* key, int def) const {
    return (int)get_num(key, (double)def);
  }
};

struct JsonParser {
  const char* p;
  const char* end;
  bool        ok = true;

  explicit JsonParser(const std::string& text)
    : p(text.data()), end(text.data() + text.size()) {}

  void skip_ws() { while (p < end && (unsigned char)*p <= ' ') p++; }

  static void append_cp(std::wstring& s, uint32_t cp) {
    if (cp <= 0xFFFF) { s += (wchar_t)cp; }
    else {
      uint32_t sv = cp - 0x10000;
      s += (wchar_t)(0xD800 | (sv >> 10));
      s += (wchar_t)(0xDC00 | (sv & 0x3FF));
    }
  }

  std::wstring parse_string_raw() {
    std::wstring out;
    skip_ws();
    if (p >= end || *p != '"') { ok = false; return out; }
    p++;
    Utf8Decoder dec;
    while (p < end && *p != '"') {
      if (*p == '\\') {
        p++;
        if (p >= end) { ok = false; break; }
        char e = *p++;
        if (e == '"')       out += L'"';
        else if (e == '\\') out += L'\\';
        else if (e == '/')  out += L'/';
        else if (e == 'n')  out += L'\n';
        else if (e == 't')  out += L'\t';
        else if (e == 'r')  out += L'\r';
        else if (e == 'b')  out += L'\b';
        else if (e == 'f')  out += L'\f';
        else if (e == 'u') {
          if (end - p < 4) { ok = false; break; }
          uint32_t cp = 0;
          for (int i = 0; i < 4; i++) {
            char c = p[i]; cp <<= 4;
            if      (c >= '0' && c <= '9') cp |= (uint32_t)(c - '0');
            else if (c >= 'a' && c <= 'f') cp |= (uint32_t)(c - 'a' + 10);
            else if (c >= 'A' && c <= 'F') cp |= (uint32_t)(c - 'A' + 10);
            else ok = false;
          }
          p += 4;
          append_cp(out, cp);
        } else ok = false;
      } else {
        uint32_t cp = 0;
        if (dec.feed((uint8_t)*p, cp)) append_cp(out, cp);
        p++;
      }
    }
    if (p < end && *p == '"') p++; else ok = false;
    return out;
  }

  JsonValue parse_number() {
    const char* start = p;
    if (p < end && (*p == '-' || *p == '+')) p++;
    while (p < end && ((*p >= '0' && *p <= '9') || *p=='.' || *p=='e' || *p=='E' || *p=='-' || *p=='+')) p++;
    JsonValue v; v.type = JsonValue::Type::Number;
    if (p == start) { ok = false; return v; }
    v.num = atof(std::string(start, p).c_str());
    return v;
  }

  JsonValue parse_object() {
    JsonValue v; v.type = JsonValue::Type::Object;
    p++; skip_ws();
    if (p < end && *p == '}') { p++; return v; }
    while (p < end) {
      std::wstring key = parse_string_raw();
      skip_ws();
      if (p >= end || *p != ':') { ok = false; break; }
      p++;
      v.obj.emplace_back(key, parse_value());
      skip_ws();
      if (p < end && *p == ',') { p++; continue; }
      if (p < end && *p == '}') { p++; break; }
      ok = false; break;
    }
    return v;
  }

  JsonValue parse_array() {
    JsonValue v; v.type = JsonValue::Type::Array;
    p++; skip_ws();
    if (p < end && *p == ']') { p++; return v; }
    while (p < end) {
      v.arr.push_back(parse_value());
      skip_ws();
      if (p < end && *p == ',') { p++; continue; }
      if (p < end && *p == ']') { p++; break; }
      ok = false; break;
    }
    return v;
  }

  JsonValue parse_value() {
    skip_ws();
    if (p >= end) { ok = false; return JsonValue{}; }
    if (*p == '{') return parse_object();
    if (*p == '[') return parse_array();
    if (*p == '"') { JsonValue v; v.type = JsonValue::Type::String; v.str = parse_string_raw(); return v; }
    if (*p == 't') {
      if (end - p >= 4 && strncmp(p, "true", 4) == 0) { p += 4; JsonValue v; v.type = JsonValue::Type::Bool; v.b = true; return v; }
      ok = false; return JsonValue{};
    }
    if (*p == 'f') {
      if (end - p >= 5 && strncmp(p, "false", 5) == 0) { p += 5; JsonValue v; v.type = JsonValue::Type::Bool; v.b = false; return v; }
      ok = false; return JsonValue{};
    }
    if (*p == 'n') {
      if (end - p >= 4 && strncmp(p, "null", 4) == 0) { p += 4; return JsonValue{}; }
      ok = false; return JsonValue{};
    }
    return parse_number();
  }
};

static bool json_parse_file(const std::wstring& path, JsonValue& out)
{
  FILE* f = nullptr;
  _wfopen_s(&f, path.c_str(), L"rb");
  if (!f) return false;
  fseek(f, 0, SEEK_END);
  long sz = ftell(f);
  fseek(f, 0, SEEK_SET);
  if (sz <= 0) { fclose(f); return false; }
  std::string buf((size_t)sz, '\0');
  size_t rd = fread(&buf[0], 1, (size_t)sz, f);
  fclose(f);
  buf.resize(rd);

  size_t start = (buf.size() >= 3 &&
    (uint8_t)buf[0] == 0xEF && (uint8_t)buf[1] == 0xBB && (uint8_t)buf[2] == 0xBF) ? 3 : 0;
  JsonParser jp(buf.substr(start));
  out = jp.parse_value();
  return jp.ok && out.type == JsonValue::Type::Object;
}

static std::string wstr_to_utf8(const std::wstring& s)
{
  if (s.empty()) return std::string();
  int len = WideCharToMultiByte(CP_UTF8, 0, s.c_str(), -1, nullptr, 0, nullptr, nullptr);
  if (len <= 0) return std::string();
  std::string out((size_t)len - 1, '\0');
  WideCharToMultiByte(CP_UTF8, 0, s.c_str(), -1, &out[0], len, nullptr, nullptr);
  return out;
}

static std::string json_escape(const std::wstring& s)
{
  std::string u8 = wstr_to_utf8(s), out;
  out.reserve(u8.size() + 8);
  for (char c : u8) {
    if      (c == '"')  out += "\\\"";
    else if (c == '\\') out += "\\\\";
    else if (c == '\n') out += "\\n";
    else if (c == '\r') out += "\\r";
    else if (c == '\t') out += "\\t";
    else                out += c;
  }
  return out;
}

// =====================================================
// CONFIG (miniterm.json — lives next to the binary)
// =====================================================
struct ShellPreset {
  std::wstring name, exe, args;
};

struct MiniConfig {
  std::wstring             theme            = L"dark";
  std::wstring             font_name;
  float                    font_size        = 16.0f;
  int                      scrollback_lines = 1000;
  std::wstring             default_shell    = L"pwsh";
  std::vector<ShellPreset> shells;
};

static MiniConfig g_config;

static std::wstring theme_to_str(Theme t) { return (t == Theme::SolarizedDark) ? L"dark" : L"light"; }
static Theme        str_to_theme(const std::wstring& s) { return (s == L"light") ? Theme::SolarizedLight : Theme::SolarizedDark; }

static std::wstring config_path()
{
  wchar_t buf[MAX_PATH] = {};
  GetModuleFileNameW(nullptr, buf, MAX_PATH);
  std::wstring p = buf;
  size_t sl = p.rfind(L'\\');
  if (sl != std::wstring::npos) p = p.substr(0, sl + 1);
  return p + L"miniterm.json";
}

static void config_set_defaults(MiniConfig& c)
{
  c.shells.clear();
  c.shells.push_back({ L"cmd", L"cmd.exe", L"" });
  // Best-guess WinPE defaults for the PowerShell family: ExecutionPolicy
  // Bypass so profile-based module loading (nerd font glyph modules,
  // prompt themes) isn't blocked by WinPE's default policy, and profile
  // loading left ON (no -NoProfile) since that's what pulls those modules
  // in. Unverified against a live WinPE rig — revisit once testable.
  c.shells.push_back({ L"powershell", L"powershell.exe", L"-NoLogo -ExecutionPolicy Bypass" });
  c.shells.push_back({ L"pwsh",       L"pwsh.exe",       L"-NoLogo -ExecutionPolicy Bypass" });
}

static bool config_load(MiniConfig& c)
{
  config_set_defaults(c);
  JsonValue root;
  if (!json_parse_file(config_path(), root)) return false;

  c.theme            = root.get_str(L"theme", c.theme);
  c.scrollback_lines = root.get_int(L"scrollback_lines", c.scrollback_lines);
  c.default_shell    = root.get_str(L"default_shell", c.default_shell);

  if (const JsonValue* fontv = root.find(L"font")) {
    c.font_name = fontv->get_str(L"name", c.font_name);
    c.font_size = (float)fontv->get_num(L"size", c.font_size);
  }

  if (const JsonValue* shellsv = root.find(L"shells")) {
    if (shellsv->type == JsonValue::Type::Array && !shellsv->arr.empty()) {
      std::vector<ShellPreset> loaded;
      for (const auto& sv : shellsv->arr) {
        if (sv.type != JsonValue::Type::Object) continue;
        ShellPreset sp;
        sp.name = sv.get_str(L"name", L"");
        sp.exe  = sv.get_str(L"exe",  L"");
        sp.args = sv.get_str(L"args", L"");
        if (!sp.name.empty() && !sp.exe.empty()) loaded.push_back(sp);
      }
      if (!loaded.empty()) c.shells = loaded;
    }
  }
  return true;
}

static bool config_save(const MiniConfig& c)
{
  FILE* f = nullptr;
  _wfopen_s(&f, config_path().c_str(), L"wb");
  if (!f) return false;

  std::string out;
  out += "{\n";
  out += "  \"theme\": \""            + json_escape(c.theme)         + "\",\n";
  out += "  \"font\": { \"name\": \"" + json_escape(c.font_name)     + "\", \"size\": "
       + std::to_string((int)c.font_size) + " },\n";
  out += "  \"scrollback_lines\": "    + std::to_string(c.scrollback_lines) + ",\n";
  out += "  \"default_shell\": \""     + json_escape(c.default_shell) + "\",\n";
  out += "  \"shells\": [\n";
  for (size_t i = 0; i < c.shells.size(); i++) {
    const auto& s = c.shells[i];
    out += "    { \"name\": \"" + json_escape(s.name) + "\", \"exe\": \"" + json_escape(s.exe)
         + "\", \"args\": \"" + json_escape(s.args) + "\" }";
    out += (i + 1 < c.shells.size()) ? ",\n" : "\n";
  }
  out += "  ]\n}\n";

  fwrite(out.data(), 1, out.size(), f);
  fclose(f);
  return true;
}

static void config_sync_from_globals()
{
  g_config.theme            = theme_to_str(g_theme);
  g_config.font_name        = g_font_name;
  g_config.font_size        = g_font_size;
  g_config.scrollback_lines = SCROLLBACK_LINES;
  // default_shell deliberately NOT synced from g_shellname here: it's a
  // preference about which shell *future* launches use, set explicitly
  // via the Preferences dialog. g_shellname is whatever the CURRENT
  // session happens to be running right now -- syncing from it meant
  // every config_save_now() call (theme change, font change, tab
  // switch, exit, ...) silently overwrote the user's chosen default
  // back to whatever shell was already running.
}

static void config_save_now()
{
  config_sync_from_globals();
  config_save(g_config);
}

// =====================================================
// GRID HELPERS
// =====================================================
static void grid_put_string(const wchar_t* msg)
{
  while (*msg) {
    uint32_t c = *msg++;
    if (c == '\n') {
      cy++; if (cy >= H) { screen_scroll_up(); cy = H-1; } cx = 0; continue;
    }
    if (cy < H && cx < W) g_screen[cy][cx++] = { c, g_cur_fg, g_cur_bg, false };
    if (cx >= W) { cx = 0; cy++; if (cy >= H) { screen_scroll_up(); cy = H-1; } }
  }
  cy++; if (cy >= H) { screen_scroll_up(); cy = H-1; } cx = 0;
}

// =====================================================
// CLIPBOARD HELPERS
// =====================================================
static void copy_selection_to_clipboard()
{
  if (!g_sel_active) return;
  int r0=g_sel_r0, c0=g_sel_c0, r1=g_sel_r1, c1=g_sel_c1;
  if (r0>r1 || (r0==r1 && c0>c1)) { std::swap(r0,r1); std::swap(c0,c1); }

  std::wstring text;
  EnterCriticalSection(&g_cs);
  for (int r = r0; r <= r1; r++) {
    const auto& row = get_view_row(r);
    int cs = (r==r0)?c0:0, ce = (r==r1)?c1:(int)row.size()-1;
    while (ce > cs && (ce >= (int)row.size() || row[ce].cp == ' ')) ce--;
    for (int c = cs; c <= ce && c < (int)row.size(); c++) {
      if (row[c].wide_cont) continue;
      uint32_t cp = row[c].cp; if (!cp) cp = ' ';
      if (cp <= 0xFFFF) text += (wchar_t)cp;
      else {
        uint32_t sv = cp - 0x10000;
        text += (wchar_t)(0xD800|(sv>>10));
        text += (wchar_t)(0xDC00|(sv&0x3FF));
      }
    }
    if (r < r1) text += L"\r\n";
  }
  LeaveCriticalSection(&g_cs);

  if (text.empty()) return;
  size_t  bytes = (text.size()+1)*sizeof(wchar_t);
  HGLOBAL hmem  = GlobalAlloc(GMEM_MOVEABLE, bytes);
  if (!hmem) return;
  memcpy(GlobalLock(hmem), text.c_str(), bytes);
  GlobalUnlock(hmem);
  OpenClipboard(hwnd); EmptyClipboard();
  SetClipboardData(CF_UNICODETEXT, hmem);
  CloseClipboard();
}

static void paste_from_clipboard()
{
  if (!OpenClipboard(hwnd)) return;
  HANDLE hdata = GetClipboardData(CF_UNICODETEXT);
  if (hdata) {
    const wchar_t* wstr = (const wchar_t*)GlobalLock(hdata);
    if (wstr) {
      int len = WideCharToMultiByte(CP_UTF8, 0, wstr, -1, nullptr, 0, nullptr, nullptr);
      if (len > 1) {
        std::vector<char> buf(len);
        WideCharToMultiByte(CP_UTF8, 0, wstr, -1, buf.data(), len, nullptr, nullptr);
        if (hInW) { DWORD written; WriteFile(hInW, buf.data(), (DWORD)(len-1), &written, NULL); }
      }
      GlobalUnlock(hdata);
    }
  }
  CloseClipboard();
}

static void clear_scrollback()
{
  EnterCriticalSection(&g_cs);
  g_history.clear(); g_view_top = 0; g_at_bottom = true;
  LeaveCriticalSection(&g_cs);
  update_scrollbar();
  InvalidateRect(hwnd, NULL, FALSE);
}

// =====================================================
// SELECTION HELPERS
// =====================================================
static bool cell_selected(int view_r, int c)
{
  if (!g_sel_active) return false;
  int r0=g_sel_r0, c0=g_sel_c0, r1=g_sel_r1, c1=g_sel_c1;
  if (r0>r1 || (r0==r1 && c0>c1)) { std::swap(r0,r1); std::swap(c0,c1); }
  if (view_r < r0 || view_r > r1) return false;
  if (view_r == r0 && c < c0)     return false;
  if (view_r == r1 && c > c1)     return false;
  return true;
}

static const std::vector<Cell>& get_view_row(int view_row)
{
  int abs  = g_view_top + view_row;
  int hist = (int)g_history.size();
  if (abs < hist) return g_history[abs];
  int sr = abs - hist;
  return (sr >= 0 && sr < H) ? g_screen[sr] : g_screen[H-1];
}

static void pixel_to_cell(int px, int py, int& row, int& col)
{
  col = max(0, min(W-1, (int)(px / CELL_W)));
  row = max(0, min(H-1, (int)((py - kTabStripH) / CELL_H)));
}

// =====================================================
// SAVE SCROLLBACK
// =====================================================
static void save_scrollback()
{
  wchar_t path[MAX_PATH] = L"miniterm_log.txt";
  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn); ofn.hwndOwner = hwnd;
  ofn.lpstrFilter = L"Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
  ofn.lpstrFile = path; ofn.nMaxFile = MAX_PATH;
  ofn.lpstrDefExt = L"txt"; ofn.lpstrTitle = L"Save Scrollback Log";
  ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
  if (!GetSaveFileNameW(&ofn)) return;

  FILE* f = nullptr; _wfopen_s(&f, path, L"wb"); if (!f) return;
  const uint8_t bom[] = { 0xEF, 0xBB, 0xBF }; fwrite(bom, 1, 3, f);

  auto wcp = [&](uint32_t cp) {
    if (!cp) cp = ' ';
    if (cp < 0x80) { uint8_t b=(uint8_t)cp; fwrite(&b,1,1,f); }
    else if (cp < 0x800) { uint8_t b[2]={(uint8_t)(0xC0|(cp>>6)),(uint8_t)(0x80|(cp&0x3F))}; fwrite(b,1,2,f); }
    else if (cp < 0x10000) { uint8_t b[3]={(uint8_t)(0xE0|(cp>>12)),(uint8_t)(0x80|((cp>>6)&0x3F)),(uint8_t)(0x80|(cp&0x3F))}; fwrite(b,1,3,f); }
    else { uint8_t b[4]={(uint8_t)(0xF0|(cp>>18)),(uint8_t)(0x80|((cp>>12)&0x3F)),(uint8_t)(0x80|((cp>>6)&0x3F)),(uint8_t)(0x80|(cp&0x3F))}; fwrite(b,1,4,f); }
  };

  EnterCriticalSection(&g_cs);
  for (const auto& row : g_history) { for (const auto& c : row) if (!c.wide_cont) wcp(c.cp); fwrite("\r\n",1,2,f); }
  for (int r = 0; r < H; r++) { for (int c = 0; c < W; c++) if (!g_screen[r][c].wide_cont) wcp(g_screen[r][c].cp); fwrite("\r\n",1,2,f); }
  LeaveCriticalSection(&g_cs);
  fclose(f);
}

// =====================================================
// FONT PICKER
// =====================================================
static void do_font_picker()
{
  LOGFONTW lf = {};
  lf.lfCharSet = DEFAULT_CHARSET;
  lf.lfPitchAndFamily = FIXED_PITCH | FF_MODERN;
  lf.lfHeight = -(int)(g_font_size * GetDeviceCaps(GetDC(hwnd), LOGPIXELSY) / 72.0f);
  if (!g_font_name.empty()) wcsncpy_s(lf.lfFaceName, g_font_name.c_str(), LF_FACESIZE-1);
  CHOOSEFONTW cf = {}; cf.lStructSize = sizeof(cf); cf.hwndOwner = hwnd; cf.lpLogFont = &lf;
  cf.Flags = CF_INITTOLOGFONTSTRUCT | CF_FIXEDPITCHONLY | CF_FORCEFONTEXIST | CF_SCREENFONTS;
  if (!ChooseFontW(&cf)) return;
  std::wstring nf = lf.lfFaceName; float ns = (float)(cf.iPointSize / 10);
  if (nf.empty() || ns < 6.0f) return;
  std::wstring of = g_font_name; float os = g_font_size;
  rebuild_fmt(nf, ns); if (!fmt) rebuild_fmt(of, os);
  g_last_client_w = 0; g_last_client_h = 0;
  pty_resize(); InvalidateRect(hwnd, NULL, FALSE);
  config_save_now();
}

// =====================================================
// PREFERENCES DIALOG
// Built as an in-memory DLGTEMPLATE so no .rc dialog resource is
// required — keeps this a single self-contained translation unit.
// =====================================================
static void dlg_align(std::vector<BYTE>& b) { while (b.size() % 4) b.push_back(0); }
static void dlg_word(std::vector<BYTE>& b, WORD w) { b.push_back((BYTE)(w & 0xFF)); b.push_back((BYTE)(w >> 8)); }
static void dlg_dword(std::vector<BYTE>& b, DWORD d) {
  for (int i = 0; i < 4; i++) b.push_back((BYTE)((d >> (i * 8)) & 0xFF));
}
static void dlg_wstr(std::vector<BYTE>& b, const wchar_t* s) {
  for (const wchar_t* p = s; ; p++) { dlg_word(b, (WORD)*p); if (!*p) break; }
}

static void dlg_item(std::vector<BYTE>& b, DWORD style, short x, short y, short cx, short cy,
                      WORD id, WORD cls_atom, const wchar_t* text)
{
  dlg_align(b);
  dlg_dword(b, style | WS_CHILD | WS_VISIBLE);
  dlg_dword(b, 0);                     // dwExtendedStyle
  dlg_word(b, x); dlg_word(b, y); dlg_word(b, cx); dlg_word(b, cy);
  dlg_word(b, id);
  dlg_word(b, 0xFFFF); dlg_word(b, cls_atom);
  dlg_wstr(b, text);
  dlg_word(b, 0);                      // no creation data
}

static const WORD kClsButton = 0x0080, kClsEdit = 0x0081,
                  kClsStatic = 0x0082, kClsCombo = 0x0085;

static std::vector<BYTE> build_prefs_template()
{
  std::vector<BYTE> b;
  const WORD kItemCount = 13;

  dlg_dword(b, DS_SETFONT | WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME);
  dlg_dword(b, 0);
  dlg_word(b, kItemCount);
  dlg_word(b, 0); dlg_word(b, 0);
  dlg_word(b, 240); dlg_word(b, 150);
  dlg_word(b, 0);                      // no menu
  dlg_word(b, 0);                      // default dialog class
  dlg_wstr(b, L"miniterm Preferences");
  dlg_word(b, 9);                      // point size
  dlg_wstr(b, L"Segoe UI");

  dlg_item(b, SS_LEFT,                       6,   8, 60, 10, (WORD)-1,              kClsStatic, L"Theme:");
  dlg_item(b, CBS_DROPDOWNLIST|WS_TABSTOP,  70,   6,100, 60, IDC_PREF_THEME,        kClsCombo,  L"");

  dlg_item(b, SS_LEFT,                       6,  24, 60, 10, (WORD)-1,              kClsStatic, L"Font name:");
  dlg_item(b, ES_AUTOHSCROLL|WS_TABSTOP|
              WS_BORDER,                    70,  22,100, 12, IDC_PREF_FONT_NAME,    kClsEdit,   L"");
  dlg_item(b, WS_TABSTOP,                  174,  22, 60, 13, IDC_PREF_FONT_BROWSE,  kClsButton, L"Browse...");

  dlg_item(b, SS_LEFT,                       6,  40, 60, 10, (WORD)-1,              kClsStatic, L"Font size:");
  dlg_item(b, ES_AUTOHSCROLL|ES_NUMBER|
              WS_TABSTOP|WS_BORDER,         70,  38, 40, 12, IDC_PREF_FONT_SIZE,    kClsEdit,   L"");

  dlg_item(b, SS_LEFT,                       6,  56, 60, 10, (WORD)-1,              kClsStatic, L"Scrollback lines:");
  dlg_item(b, ES_AUTOHSCROLL|ES_NUMBER|
              WS_TABSTOP|WS_BORDER,         70,  54, 60, 12, IDC_PREF_SCROLLBACK,   kClsEdit,   L"");

  dlg_item(b, SS_LEFT,                       6,  72, 60, 10, (WORD)-1,              kClsStatic, L"Default shell:");
  dlg_item(b, CBS_DROPDOWNLIST|WS_TABSTOP,  70,  70,100, 60, IDC_PREF_DEFAULT_SHELL,kClsCombo,  L"");

  dlg_item(b, BS_DEFPUSHBUTTON|WS_TABSTOP,  96, 128, 60, 14, IDC_PREF_OK,           kClsButton, L"OK");
  dlg_item(b, WS_TABSTOP,                  162, 128, 60, 14, IDC_PREF_CANCEL,       kClsButton, L"Cancel");

  return b;
}

struct PrefsResult {
  Theme        theme         = Theme::SolarizedDark;
  std::wstring font_name;
  float        font_size     = 16.0f;
  int          scrollback    = 1000;
  std::wstring default_shell;
} g_prefs_result;

static INT_PTR CALLBACK prefs_dlg_proc(HWND hDlg, UINT msg, WPARAM wp, LPARAM)
{
  switch (msg) {
    case WM_INITDIALOG: {
      HWND cbTheme = GetDlgItem(hDlg, IDC_PREF_THEME);
      SendMessageW(cbTheme, CB_ADDSTRING, 0, (LPARAM)L"Solarized Dark");
      SendMessageW(cbTheme, CB_ADDSTRING, 0, (LPARAM)L"Solarized Light");
      SendMessageW(cbTheme, CB_SETCURSEL, (g_theme == Theme::SolarizedDark) ? 0 : 1, 0);

      SetDlgItemTextW(hDlg, IDC_PREF_FONT_NAME, g_font_name.c_str());
      wchar_t buf[32];
      wsprintfW(buf, L"%d", (int)g_font_size);
      SetDlgItemTextW(hDlg, IDC_PREF_FONT_SIZE, buf);
      wsprintfW(buf, L"%d", SCROLLBACK_LINES);
      SetDlgItemTextW(hDlg, IDC_PREF_SCROLLBACK, buf);

      HWND cbShell = GetDlgItem(hDlg, IDC_PREF_DEFAULT_SHELL);
      int sel = 0;
      for (size_t i = 0; i < g_config.shells.size(); i++) {
        SendMessageW(cbShell, CB_ADDSTRING, 0, (LPARAM)g_config.shells[i].name.c_str());
        if (g_config.shells[i].name == g_shellname) sel = (int)i;
      }
      SendMessageW(cbShell, CB_SETCURSEL, sel, 0);
      return TRUE;
    }

    case WM_COMMAND:
      switch (LOWORD(wp)) {
        case IDC_PREF_FONT_BROWSE: {
          LOGFONTW lf = {};
          lf.lfCharSet = DEFAULT_CHARSET;
          lf.lfPitchAndFamily = FIXED_PITCH | FF_MODERN;
          wchar_t cur[256] = {};
          GetDlgItemTextW(hDlg, IDC_PREF_FONT_NAME, cur, 256);
          if (cur[0]) wcsncpy_s(lf.lfFaceName, cur, LF_FACESIZE - 1);
          wchar_t szbuf[32] = {};
          GetDlgItemTextW(hDlg, IDC_PREF_FONT_SIZE, szbuf, 32);
          int pts = szbuf[0] ? _wtoi(szbuf) : 16;
          lf.lfHeight = -(int)(pts * GetDeviceCaps(GetDC(hDlg), LOGPIXELSY) / 72.0f);
          CHOOSEFONTW cf = {}; cf.lStructSize = sizeof(cf); cf.hwndOwner = hDlg; cf.lpLogFont = &lf;
          cf.Flags = CF_INITTOLOGFONTSTRUCT | CF_FIXEDPITCHONLY | CF_FORCEFONTEXIST | CF_SCREENFONTS;
          if (ChooseFontW(&cf)) {
            SetDlgItemTextW(hDlg, IDC_PREF_FONT_NAME, lf.lfFaceName);
            wchar_t nb[32]; wsprintfW(nb, L"%d", cf.iPointSize / 10);
            SetDlgItemTextW(hDlg, IDC_PREF_FONT_SIZE, nb);
          }
          return TRUE;
        }
        case IDC_PREF_OK: {
          HWND cbTheme = GetDlgItem(hDlg, IDC_PREF_THEME);
          g_prefs_result.theme = (SendMessageW(cbTheme, CB_GETCURSEL, 0, 0) == 0)
            ? Theme::SolarizedDark : Theme::SolarizedLight;

          wchar_t buf[256];
          GetDlgItemTextW(hDlg, IDC_PREF_FONT_NAME, buf, 256);
          g_prefs_result.font_name = buf;

          wchar_t szbuf[32];
          GetDlgItemTextW(hDlg, IDC_PREF_FONT_SIZE, szbuf, 32);
          g_prefs_result.font_size = szbuf[0] ? (float)_wtoi(szbuf) : g_font_size;

          wchar_t sbbuf[32];
          GetDlgItemTextW(hDlg, IDC_PREF_SCROLLBACK, sbbuf, 32);
          g_prefs_result.scrollback = sbbuf[0] ? _wtoi(sbbuf) : SCROLLBACK_LINES;

          HWND cbShell = GetDlgItem(hDlg, IDC_PREF_DEFAULT_SHELL);
          int sel = (int)SendMessageW(cbShell, CB_GETCURSEL, 0, 0);
          wchar_t shbuf[64] = {};
          if (sel >= 0) SendMessageW(cbShell, CB_GETLBTEXT, sel, (LPARAM)shbuf);
          g_prefs_result.default_shell = shbuf;

          EndDialog(hDlg, IDC_PREF_OK);
          return TRUE;
        }
        case IDC_PREF_CANCEL: EndDialog(hDlg, IDC_PREF_CANCEL); return TRUE;
      }
      return FALSE;

    case WM_CLOSE:
      EndDialog(hDlg, IDC_PREF_CANCEL);
      return TRUE;
  }
  return FALSE;
}

static void switch_shell(const ShellPreset& sp);
static void new_session(const ShellPreset& sp);
static void close_session(int idx, bool already_dead);
static void switch_to_session(int idx);
static void rebuild_tab_ui();

static void do_preferences()
{
  std::vector<BYTE> tmpl = build_prefs_template();
  INT_PTR r = DialogBoxIndirectParamW(GetModuleHandleW(NULL),
    (LPCDLGTEMPLATEW)tmpl.data(), hwnd, prefs_dlg_proc, 0);
  if (r != IDC_PREF_OK) return;

  bool theme_changed = (g_prefs_result.theme != g_theme);
  bool font_changed  = (g_prefs_result.font_name != g_font_name) ||
                        (fabsf(g_prefs_result.font_size - g_font_size) > 0.01f);

  if (theme_changed) {
    g_theme = g_prefs_result.theme;
    g_cur_fg = theme_fg(); g_cur_bg = theme_bg();
    build_palette(g_theme);
  }

  if (font_changed && !g_prefs_result.font_name.empty() && g_prefs_result.font_size >= 6.0f) {
    std::wstring of = g_font_name; float os = g_font_size;
    rebuild_fmt(g_prefs_result.font_name, g_prefs_result.font_size);
    if (!fmt) rebuild_fmt(of, os);
    g_last_client_w = 0; g_last_client_h = 0;
    pty_resize();
  }

  if (g_prefs_result.scrollback >= 100) SCROLLBACK_LINES = g_prefs_result.scrollback;

  // Default shell only takes effect for the next launched session — it
  // does not relaunch the shell already running in this window.
  if (!g_prefs_result.default_shell.empty())
    g_config.default_shell = g_prefs_result.default_shell;

  update_statusbar();
  config_save_now();
  InvalidateRect(hwnd, NULL, FALSE);
}

// =====================================================
// RENAME TAB DIALOG — same in-memory DLGTEMPLATE technique as
// Preferences, reusing its dlg_* helpers.
// =====================================================
static std::vector<BYTE> build_rename_template()
{
  std::vector<BYTE> b;
  const WORD kItemCount = 3;

  dlg_dword(b, DS_SETFONT | WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME);
  dlg_dword(b, 0);
  dlg_word(b, kItemCount);
  dlg_word(b, 0); dlg_word(b, 0);
  dlg_word(b, 180); dlg_word(b, 58);
  dlg_word(b, 0);
  dlg_word(b, 0);
  dlg_wstr(b, L"Rename Tab");
  dlg_word(b, 9);
  dlg_wstr(b, L"Segoe UI");

  dlg_item(b, ES_AUTOHSCROLL|WS_TABSTOP|WS_BORDER, 6,  8, 168, 12, IDC_RENAME_EDIT,   kClsEdit,   L"");
  dlg_item(b, BS_DEFPUSHBUTTON|WS_TABSTOP,        36, 32,  50, 14, IDC_RENAME_OK,     kClsButton, L"OK");
  dlg_item(b, WS_TABSTOP,                         94, 32,  50, 14, IDC_RENAME_CANCEL, kClsButton, L"Cancel");

  return b;
}

struct RenameResult { std::wstring name; bool ok = false; } g_rename_result;

static INT_PTR CALLBACK rename_dlg_proc(HWND hDlg, UINT msg, WPARAM wp, LPARAM)
{
  switch (msg) {
    case WM_INITDIALOG:
      SetDlgItemTextW(hDlg, IDC_RENAME_EDIT, g_rename_result.name.c_str());
      SendMessageW(GetDlgItem(hDlg, IDC_RENAME_EDIT), EM_SETSEL, 0, -1);
      SetFocus(GetDlgItem(hDlg, IDC_RENAME_EDIT));
      return FALSE;  // we set focus ourselves
    case WM_COMMAND:
      switch (LOWORD(wp)) {
        case IDC_RENAME_OK: {
          wchar_t buf[64] = {};
          GetDlgItemTextW(hDlg, IDC_RENAME_EDIT, buf, 64);
          g_rename_result.name = buf;
          g_rename_result.ok = true;
          EndDialog(hDlg, IDC_RENAME_OK);
          return TRUE;
        }
        case IDC_RENAME_CANCEL: EndDialog(hDlg, IDC_RENAME_CANCEL); return TRUE;
      }
      return FALSE;
    case WM_CLOSE:
      EndDialog(hDlg, IDC_RENAME_CANCEL);
      return TRUE;
  }
  return FALSE;
}

static void do_rename_tab(int idx)
{
  if (idx < 0 || idx >= (int)g_sessions.size()) return;
  g_rename_result.name = g_sessions[idx]->name;
  g_rename_result.ok   = false;

  std::vector<BYTE> tmpl = build_rename_template();
  DialogBoxIndirectParamW(GetModuleHandleW(NULL),
    (LPCDLGTEMPLATEW)tmpl.data(), hwnd, rename_dlg_proc, 0);

  if (!g_rename_result.ok || g_rename_result.name.empty()) return;
  g_sessions[idx]->name = g_rename_result.name;
  rebuild_tab_ui();
}

// =====================================================
// ABOUT DIALOG
// The about text lives in a read-only multiline edit (rather than a
// static) purely so it can receive WM_CHAR: typing 'ø' or 'Ø' closes
// the box and shows a small Danish aside. Uses \u escapes for the
// Danish characters rather than literal bytes, since the build never
// pins a source charset (no /utf-8 or /source-charset flag) -- literal
// non-ASCII bytes here would be at the mercy of whatever codepage MSVC
// assumes for this file.
// =====================================================
static const wchar_t* k_about_text =
  L"miniterm v2.8\r\nConPTY WinPE Terminal\r\n\r\n"
  L"Glyph rendering fix: layout maxWidth=CELL_W*3,\r\n"
  L"no clip on narrow glyphs, fat -debug mode.";

static std::vector<BYTE> build_about_template()
{
  std::vector<BYTE> b;
  const WORD kItemCount = 2;

  dlg_dword(b, DS_SETFONT | WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME);
  dlg_dword(b, 0);
  dlg_word(b, kItemCount);
  dlg_word(b, 0); dlg_word(b, 0);
  dlg_word(b, 220); dlg_word(b, 92);
  dlg_word(b, 0);
  dlg_word(b, 0);
  dlg_wstr(b, L"About miniterm");
  dlg_word(b, 9);
  dlg_wstr(b, L"Segoe UI");

  dlg_item(b, ES_MULTILINE|ES_READONLY|WS_TABSTOP,
           6,  8, 208, 56, IDC_ABOUT_TEXT, kClsEdit,   L"");
  dlg_item(b, BS_DEFPUSHBUTTON|WS_TABSTOP,
           82, 70,  56, 14, IDC_ABOUT_OK,  kClsButton, L"OK");

  return b;
}

static LRESULT CALLBACK about_edit_subclass(HWND hEdit, UINT msg, WPARAM wp, LPARAM lp,
                                             UINT_PTR, DWORD_PTR)
{
  if (msg == WM_CHAR && (wp == 0x00F8 || wp == 0x00D8)) {  // ø = o slash, Ø = O slash
    SYSTEMTIME st = {};
    GetLocalTime(&st);
    bool christmas_eve = (st.wMonth == 12 && st.wDay == 24);
    const wchar_t* egg = christmas_eve
      ? L"R\u00F8dgr\u00F8d med fl\u00F8de"
      : L"Der er ingen ko p\u00E5 isen!";
    EndDialog(GetParent(hEdit), IDOK);
    MessageBoxW(hwnd, egg, L"miniterm", MB_OK | MB_ICONINFORMATION);
    return 0;
  }
  return DefSubclassProc(hEdit, msg, wp, lp);
}

static INT_PTR CALLBACK about_dlg_proc(HWND hDlg, UINT msg, WPARAM wp, LPARAM)
{
  switch (msg) {
    case WM_INITDIALOG: {
      HWND edit = GetDlgItem(hDlg, IDC_ABOUT_TEXT);
      SetWindowTextW(edit, k_about_text);
      SetWindowSubclass(edit, about_edit_subclass, 1, 0);
      SetFocus(edit);
      return FALSE;  // we set focus ourselves
    }
    case WM_COMMAND:
      if (LOWORD(wp) == IDC_ABOUT_OK) { EndDialog(hDlg, IDOK); return TRUE; }
      return FALSE;
    case WM_CLOSE:
      EndDialog(hDlg, IDCANCEL);
      return TRUE;
  }
  return FALSE;
}

static void do_about()
{
  std::vector<BYTE> tmpl = build_about_template();
  DialogBoxIndirectParamW(GetModuleHandleW(NULL),
    (LPCDLGTEMPLATEW)tmpl.data(), hwnd, about_dlg_proc, 0);
}

// =====================================================
// RENDER — message thread only
//
// KEY DESIGN DECISIONS:
//
// 1. Narrow glyphs (including Nerd Font icons):
//    - Background rect filled at CELL_W width
//    - Glyph drawn with NO clip rect
//    - Layout created with maxWidth=CELL_W*3 so ink is never
//      truncated at layout creation time
//    - Ink allowed to bleed right; next cell's bg rect covers it
//    This is how Windows Terminal renders — never clip glyphs.
//
// 2. Wide glyphs (CJK, advance >= 1.5x M):
//    - Background rect filled at wide_glyph_w() width
//    - Clipped to that rect (ink contained within 2-cell area)
//    - Layout with NO_WRAP and wide_glyph_w() max width
//
// 3. Two reusable brushes per frame — no per-cell COM alloc.
// =====================================================
static void render()
{
  if (!rt || !fmt) return;

  int snapW = W, snapH = H;
  std::vector<std::vector<Cell>> snap(snapH, std::vector<Cell>(snapW));
  int snap_cx, snap_cy; bool at_bottom;

  EnterCriticalSection(&g_cs);
  for (int r = 0; r < snapH; r++) {
    const auto& src = get_view_row(r);
    for (int c = 0; c < snapW; c++)
      snap[r][c] = (c < (int)src.size()) ? src[c] : Cell{};
  }
  snap_cx = cx; snap_cy = cy; at_bottom = g_at_bottom;
  LeaveCriticalSection(&g_cs);

  rt->BeginDraw();
  rt->Clear(g_palette[theme_bg()]);
  // Shift all terminal content down below the tab strip -- that top
  // strip is painted over by the tab control child window regardless,
  // this just keeps the grid's row 0 visually aligned under it.
  rt->SetTransform(D2D1::Matrix3x2F::Translation(0.0f, (float)kTabStripH));

  ID2D1SolidColorBrush* fg_brush = nullptr;
  ID2D1SolidColorBrush* bg_brush = nullptr;
  rt->CreateSolidColorBrush(g_palette[theme_fg()], &fg_brush);
  rt->CreateSolidColorBrush(g_palette[theme_bg()], &bg_brush);

  if (!fg_brush || !bg_brush) {
    if (fg_brush) fg_brush->Release();
    if (bg_brush) bg_brush->Release();
    rt->EndDraw(); return;
  }

  // Debug: count glyphs rendered this frame
  static int s_frame = 0;
  int dbg_wide = 0, dbg_narrow = 0, dbg_skip = 0;
  s_frame++;

  for (int r = 0; r < snapH; r++) {
    float y = r * CELL_H;
    for (int c = 0; c < snapW; c++) {
      const Cell& cell = snap[r][c];
      if (cell.wide_cont) { dbg_skip++; continue; }

      float    x    = c * CELL_W;
      bool     wide = (c+1 < snapW && snap[r][c+1].wide_cont);
      uint32_t cp   = cell.cp;
      if (cp < 0x20) cp = ' ';

      bool is_cursor = at_bottom && g_cursor_visible
                    && r == snap_cy && c == snap_cx;
      bool selected  = cell_selected(r, c);

      D2D1_COLOR_F bg_col, fg_col;
      if (is_cursor) {
        bg_col = g_palette[cell.fg == cell.bg ? theme_fg() : cell.fg];
        fg_col = g_palette[cell.bg];
      } else if (selected) {
        bg_col = g_palette[cell.fg]; fg_col = g_palette[cell.bg];
      } else {
        bg_col = g_palette[cell.bg]; fg_col = g_palette[cell.fg];
      }

      // Background rect
      float bg_w = wide ? wide_glyph_w(cp) : CELL_W;
      D2D1_RECT_F bg_rect = D2D1::RectF(x, y, x + bg_w, y + CELL_H);
      bg_brush->SetColor(bg_col);
      rt->FillRectangle(bg_rect, bg_brush);

      fg_brush->SetColor(fg_col);

      if (wide) {
        // Wide glyph: clip to measured width, NO_WRAP layout
        dbg_wide++;
        float glyph_w = wide_glyph_w(cp);
        D2D1_RECT_F clip = D2D1::RectF(x, y, x + glyph_w, y + CELL_H);
        rt->PushAxisAlignedClip(clip, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        wchar_t gbuf[3] = {}; int glen = 0;
        if (cp <= 0xFFFF) { gbuf[0] = (wchar_t)cp; glen = 1; }
        else {
          uint32_t sv = cp - 0x10000;
          gbuf[0] = (wchar_t)(0xD800|(sv>>10));
          gbuf[1] = (wchar_t)(0xDC00|(sv&0x3FF));
          glen = 2;
        }

        IDWriteTextLayout* wl = nullptr;
        if (SUCCEEDED(dw->CreateTextLayout(
            gbuf, (UINT32)glen, fmt, glyph_w, CELL_H, &wl)) && wl) {
          wl->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
          wl->SetLineSpacing(DWRITE_LINE_SPACING_METHOD_UNIFORM, CELL_H, CELL_H * 0.8f);
          if (g_debug) {
            DWRITE_TEXT_METRICS tm = {};
            wl->GetMetrics(&tm);
            dbg("RENDER_WIDE U+%04X x=%.1f glyph_w=%.2f "
                "layout_w=%.2f trailing=%.2f\n",
                cp, x, glyph_w, tm.width,
                tm.widthIncludingTrailingWhitespace);
          }
          rt->DrawTextLayout(D2D1::Point2F(x, y), wl, fg_brush);
          wl->Release();
        }
        rt->PopAxisAlignedClip();

      } else {
        // Narrow glyph: NO clip — ink may bleed rightward.
        // Layout maxWidth=CELL_W*3 ensures ink not truncated at source.
        // Next cell's background rect covers any bleed.
        dbg_narrow++;
        IDWriteTextLayout* layout = get_cached_layout(cp);
        if (layout)
          rt->DrawTextLayout(D2D1::Point2F(x, y), layout, fg_brush);
      }
    }
  }

  fg_brush->Release();
  bg_brush->Release();
  rt->EndDraw();

  if (g_debug && (s_frame % 60 == 1)) {
    dbg("FRAME %d: wide=%d narrow=%d skip(wide_cont)=%d\n",
        s_frame, dbg_wide, dbg_narrow, dbg_skip);
  }
}

// =====================================================
// PTY + RENDER TARGET RESIZE
// =====================================================
static void pty_resize()
{
  if (!hwnd) return;
  RECT rc; GetClientRect(hwnd, &rc);
  int sb_h = 0, scr_w = GetSystemMetrics(SM_CXVSCROLL);
  if (g_statusbar) { RECT s; GetWindowRect(g_statusbar, &s); sb_h = s.bottom - s.top; }

  if (g_tabctrl) SetWindowPos(g_tabctrl, NULL, 0, 0, rc.right - rc.left, kTabStripH, SWP_NOZORDER);

  int client_h = (rc.bottom - rc.top) - sb_h - kTabStripH;
  int client_w = (rc.right  - rc.left) - scr_w;

  if (client_w == g_last_client_w && client_h == g_last_client_h) return;
  g_last_client_w = client_w; g_last_client_h = client_h;

  int newW = max(10, (int)floorf((float)client_w / CELL_W));
  int newH = max(4,  (int)floorf((float)client_h / CELL_H));

  EnterCriticalSection(&g_cs); grid_resize(newW, newH); LeaveCriticalSection(&g_cs);

  if (g_hpc) { COORD sz = {(SHORT)newW,(SHORT)newH}; ResizePseudoConsole(g_hpc, sz); }
  if (g_scrollbar) SetWindowPos(g_scrollbar, NULL, rc.right-scr_w, kTabStripH, scr_w, client_h, SWP_NOZORDER);
  if (rt) rt->Resize(D2D1::SizeU(rc.right-rc.left, rc.bottom-rc.top));
  update_scrollbar();

  dbg("pty_resize: client=%dx%d CELL=%.3fx%.3f W=%d H=%d\n",
      client_w, client_h, CELL_W, CELL_H, newW, newH);
}

// =====================================================
// KEYBOARD INPUT
// =====================================================
static void write_pty(const char* buf, DWORD len)
{ if (!hInW) return; DWORD w; WriteFile(hInW, buf, len, &w, NULL); }

static void write_pty_str(const char* s) { write_pty(s, (DWORD)strlen(s)); }

static void scroll_view(int delta)
{
  EnterCriticalSection(&g_cs);
  int hist = (int)g_history.size(), total = hist+H, max_top = max(0,total-H);
  g_view_top = max(0, min(max_top, g_view_top+delta));
  g_at_bottom = (g_view_top >= hist);
  LeaveCriticalSection(&g_cs);
  update_scrollbar(); InvalidateRect(hwnd, NULL, FALSE);
}

static void snap_to_bottom()
{
  EnterCriticalSection(&g_cs);
  g_view_top = (int)g_history.size(); g_at_bottom = true;
  LeaveCriticalSection(&g_cs);
  update_scrollbar();
}

static void on_char(wchar_t wc)
{
  if (wc == '\t' || wc == '\x08') return;
  if (wc == '\x03' && g_sel_active) {
    copy_selection_to_clipboard(); g_sel_active = false;
    InvalidateRect(hwnd, NULL, FALSE); return;
  }
  snap_to_bottom();
  char buf[5] = {}; uint32_t cp = wc;
  if (cp < 0x80) { buf[0]=(char)cp; write_pty(buf,1); }
  else if (cp < 0x800) { buf[0]=(char)(0xC0|(cp>>6)); buf[1]=(char)(0x80|(cp&0x3F)); write_pty(buf,2); }
  else { buf[0]=(char)(0xE0|(cp>>12)); buf[1]=(char)(0x80|((cp>>6)&0x3F)); buf[2]=(char)(0x80|(cp&0x3F)); write_pty(buf,3); }
}

static void on_keydown(WPARAM vk)
{
  switch (vk) {
    case VK_TAB:    write_pty_str("\t");       break;
    case VK_BACK:   write_pty_str("\x7f");     break;
    case VK_PRIOR:  scroll_view(-H);           break;
    case VK_NEXT:   scroll_view(+H);           break;
    case VK_UP:     write_pty_str("\x1b[A");   break;
    case VK_DOWN:   write_pty_str("\x1b[B");   break;
    case VK_RIGHT:  write_pty_str("\x1b[C");   break;
    case VK_LEFT:   write_pty_str("\x1b[D");   break;
    case VK_HOME:   write_pty_str("\x1b[H");   break;
    case VK_END:    write_pty_str("\x1b[F");   break;
    case VK_DELETE: write_pty_str("\x1b[3~");  break;
    case VK_F1:     write_pty_str("\x1bOP");   break;
    case VK_F2:     write_pty_str("\x1bOQ");   break;
    case VK_F3:     write_pty_str("\x1bOR");   break;
    case VK_F4:     write_pty_str("\x1bOS");   break;
    case VK_F5:     write_pty_str("\x1b[15~"); break;
    case VK_F6:     write_pty_str("\x1b[17~"); break;
    case VK_F7:     write_pty_str("\x1b[18~"); break;
    case VK_F8:     write_pty_str("\x1b[19~"); break;
    case VK_F9:     write_pty_str("\x1b[20~"); break;
    case VK_F10:    write_pty_str("\x1b[21~"); break;
    case VK_F11:    write_pty_str("\x1b[23~"); break;
    case VK_F12:    write_pty_str("\x1b[24~"); break;
    default: break;
  }
}

// =====================================================
// CONTEXT MENU
// =====================================================
static void show_context_menu(int sx, int sy)
{
  HMENU menu = CreatePopupMenu();
  UINT copy_flags = MF_STRING | (g_sel_active ? 0 : MF_GRAYED);
  AppendMenuW(menu, copy_flags, IDM_CTX_COPY,             L"Copy");
  AppendMenuW(menu, MF_STRING,  IDM_CTX_PASTE,            L"Paste");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING,  IDM_CTX_CLEAR_SCROLLBACK, L"Clear Scrollback");
  TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, sx, sy, 0, hwnd, nullptr);
  DestroyMenu(menu);
}

// =====================================================
// MENU BAR
// =====================================================
static HMENU g_shell_menu  = nullptr;
static HMENU g_newtab_menu = nullptr;
static int   g_ctx_tab_idx = -1;

static HMENU create_menu()
{
  HMENU bar = CreateMenu(), file = CreatePopupMenu(), term = CreatePopupMenu(),
        theme = CreatePopupMenu(), help = CreatePopupMenu();
  g_shell_menu = CreatePopupMenu();
  for (size_t i = 0; i < g_config.shells.size() && i < IDM_SHELL_MAX; i++) {
    UINT flags = MF_STRING | ((g_config.shells[i].name == g_shellname) ? MF_CHECKED : 0);
    AppendMenuW(g_shell_menu, flags, IDM_SHELL_BASE + (UINT)i, g_config.shells[i].name.c_str());
  }
  g_newtab_menu = CreatePopupMenu();
  for (size_t i = 0; i < g_config.shells.size() && i < IDM_SHELL_MAX; i++)
    AppendMenuW(g_newtab_menu, MF_STRING, IDM_NEWTAB_BASE + (UINT)i, g_config.shells[i].name.c_str());
  AppendMenuW(theme, MF_STRING, IDM_FILE_THEME_DARK,  L"Solarized Dark");
  AppendMenuW(theme, MF_STRING, IDM_FILE_THEME_LIGHT, L"Solarized Light");

  AppendMenuW(file, MF_POPUP, (UINT_PTR)g_newtab_menu, L"New");
  AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(file, MF_STRING, IDM_FILE_SAVELOG, L"Save Scrollback to Log...");
  AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(file, MF_STRING, IDM_FILE_PREFS,   L"Preferences...");
  AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(file, MF_STRING, IDM_FILE_EXIT,    L"Exit");

  AppendMenuW(term, MF_POPUP, (UINT_PTR)g_shell_menu, L"Shell");
  AppendMenuW(term, MF_POPUP, (UINT_PTR)theme, L"Theme");
  AppendMenuW(term, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(term, MF_STRING, IDM_FILE_FONT, L"Font...");

  AppendMenuW(help, MF_STRING, IDM_HELP_ABOUT, L"About");

  AppendMenuW(bar, MF_POPUP, (UINT_PTR)file, L"File");
  AppendMenuW(bar, MF_POPUP, (UINT_PTR)term, L"Terminal");
  AppendMenuW(bar, MF_POPUP, (UINT_PTR)help, L"Help");
  return bar;
}

// =====================================================
// WINDOW PROC
// =====================================================
LRESULT CALLBACK WndProc(HWND h, UINT m, WPARAM w, LPARAM l)
{
  switch (m) {
    case WM_CHAR:    on_char((wchar_t)w); return 0;
    case WM_KEYDOWN: on_keydown(w);       return 0;

    case WM_SYSKEYDOWN:
      if (w == VK_F4 || w == VK_MENU) break;
      { char s[3]={'\x1b',0,0};
        UINT ch = MapVirtualKeyW((UINT)w, MAPVK_VK_TO_CHAR);
        if (ch >= 32 && ch < 127) { s[1]=(char)(ch&0xFF); write_pty(s,2); } }
      return 0;

    case WM_SYSCHAR:
      if (w == VK_F4) break; return 0;

    case WM_LBUTTONDOWN: {
      int row, col; pixel_to_cell(LOWORD(l), HIWORD(l), row, col);
      g_sel_r0=g_sel_r1=row; g_sel_c0=g_sel_c1=col;
      g_sel_active=false; g_sel_dragging=true; SetCapture(h); return 0; }

    case WM_MOUSEMOVE:
      if (g_sel_dragging) {
        int row, col; pixel_to_cell(LOWORD(l), HIWORD(l), row, col);
        g_sel_r1=row; g_sel_c1=col;
        g_sel_active=(g_sel_r1!=g_sel_r0||g_sel_c1!=g_sel_c0);
        InvalidateRect(h, NULL, FALSE); }
      return 0;

    case WM_LBUTTONUP:
      if (g_sel_dragging) {
        int row, col; pixel_to_cell(LOWORD(l), HIWORD(l), row, col);
        g_sel_r1=row; g_sel_c1=col;
        g_sel_active=(g_sel_r1!=g_sel_r0||g_sel_c1!=g_sel_c0);
        g_sel_dragging=false; ReleaseCapture(); InvalidateRect(h, NULL, FALSE); }
      return 0;

    case WM_RBUTTONDOWN:
      if (g_sel_active) {
        copy_selection_to_clipboard(); g_sel_active=false; InvalidateRect(h,NULL,FALSE);
      } else {
        POINT pt={LOWORD(l),HIWORD(l)}; ClientToScreen(h,&pt);
        show_context_menu(pt.x, pt.y); }
      return 0;

    case WM_MOUSEWHEEL:
      scroll_view(GET_WHEEL_DELTA_WPARAM(w) > 0 ? -3 : 3); return 0;

    case WM_VSCROLL:
      if ((HWND)l == g_scrollbar) {
        EnterCriticalSection(&g_cs);
        int hist=(int)g_history.size(), total=hist+H, max_top=max(0,total-H);
        LeaveCriticalSection(&g_cs);
        SCROLLINFO si={}; si.cbSize=sizeof(si); si.fMask=SIF_ALL;
        GetScrollInfo(g_scrollbar, SB_CTL, &si);
        int nt=g_view_top;
        switch (LOWORD(w)) {
          case SB_LINEUP:        nt--;            break;
          case SB_LINEDOWN:      nt++;            break;
          case SB_PAGEUP:        nt-=H;           break;
          case SB_PAGEDOWN:      nt+=H;           break;
          case SB_THUMBTRACK:    nt=si.nTrackPos; break;
          case SB_THUMBPOSITION: nt=si.nPos;      break;
          case SB_TOP:           nt=0;            break;
          case SB_BOTTOM:        nt=max_top;      break;
        }
        nt=max(0,min(max_top,nt));
        g_view_top=nt; g_at_bottom=(nt>=(int)g_history.size());
        update_scrollbar(); InvalidateRect(h,NULL,FALSE); }
      return 0;

    case WM_SIZE:
      if (g_statusbar) SendMessageW(g_statusbar, WM_SIZE, 0, 0);
      pty_resize(); InvalidateRect(h, NULL, FALSE); return 0;

    case WM_TIMER:
      if (w == IDT_SIZE_POLL) {
        pty_resize();
        // ConPTY doesn't reliably close the output pipe just because the
        // hosted shell exited (e.g. typing "exit") -- reader() can stay
        // blocked in ReadFile indefinitely otherwise. Piggyback on this
        // existing poll to notice and unblock it. Non-blocking check
        // (timeout 0), main thread only, so no new synchronisation needed.
        if (g_reader_thread && g_child_pi.hProcess &&
            WaitForSingleObject(g_child_pi.hProcess, 0) == WAIT_OBJECT_0) {
          g_watcher_signals_exit = true;
          if (hOutR) CancelIoEx(hOutR, NULL);
        }
      } else if (w == IDT_CURSOR_BLINK) {
        g_cursor_visible = !g_cursor_visible;
        if (g_at_bottom) InvalidateRect(h, NULL, FALSE); }
      return 0;

    case WM_PAINT: {
      PAINTSTRUCT ps; BeginPaint(h, &ps); render(); EndPaint(h, &ps); return 0; }

    case WM_COMMAND:
      if (LOWORD(w) >= IDM_SHELL_BASE && LOWORD(w) < IDM_SHELL_BASE + IDM_SHELL_MAX) {
        size_t idx = LOWORD(w) - IDM_SHELL_BASE;
        if (idx < g_config.shells.size()) switch_shell(g_config.shells[idx]);
        return 0;
      }
      if (LOWORD(w) >= IDM_NEWTAB_BASE && LOWORD(w) < IDM_NEWTAB_BASE + IDM_SHELL_MAX) {
        size_t idx = LOWORD(w) - IDM_NEWTAB_BASE;
        if (idx < g_config.shells.size()) new_session(g_config.shells[idx]);
        return 0;
      }
      switch (LOWORD(w)) {
        case IDM_CTX_COPY:
          if (g_sel_active) { copy_selection_to_clipboard(); g_sel_active=false; InvalidateRect(h,NULL,FALSE); } return 0;
        case IDM_CTX_PASTE:        paste_from_clipboard(); snap_to_bottom(); return 0;
        case IDM_CTX_CLEAR_SCROLLBACK: clear_scrollback(); return 0;
        case IDM_FILE_THEME_DARK:
          g_theme=Theme::SolarizedDark; g_cur_fg=theme_fg(); g_cur_bg=theme_bg();
          build_palette(g_theme); update_statusbar(); InvalidateRect(h,NULL,FALSE);
          config_save_now(); return 0;
        case IDM_FILE_THEME_LIGHT:
          g_theme=Theme::SolarizedLight; g_cur_fg=theme_fg(); g_cur_bg=theme_bg();
          build_palette(g_theme); update_statusbar(); InvalidateRect(h,NULL,FALSE);
          config_save_now(); return 0;
        case IDM_FILE_FONT:    do_font_picker();  return 0;
        case IDM_FILE_SAVELOG: save_scrollback(); return 0;
        case IDM_FILE_PREFS:   do_preferences();  return 0;
        case IDM_FILE_EXIT:    DestroyWindow(h);  return 0;
        case IDM_TAB_CLOSE:
          if (g_ctx_tab_idx >= 0) close_session(g_ctx_tab_idx, false);
          g_ctx_tab_idx = -1;
          return 0;
        case IDM_TAB_RENAME:
          if (g_ctx_tab_idx >= 0) do_rename_tab(g_ctx_tab_idx);
          g_ctx_tab_idx = -1;
          return 0;
        case IDM_HELP_ABOUT: do_about(); return 0;
      }
      return 0;

    case WM_NOTIFY: {
      NMHDR* nm = (NMHDR*)l;
      if (nm->hwndFrom == g_tabctrl) {
        if (nm->code == TCN_SELCHANGE) {
          int idx = (int)SendMessageW(g_tabctrl, TCM_GETCURSEL, 0, 0);
          switch_to_session(idx);
          return 0;
        }
        if (nm->code == NM_RCLICK) {
          POINT pt; GetCursorPos(&pt);
          POINT client = pt; ScreenToClient(g_tabctrl, &client);
          TCHITTESTINFO hit = {}; hit.pt = client;
          int idx = (int)SendMessageW(g_tabctrl, TCM_HITTEST, 0, (LPARAM)&hit);
          if (idx >= 0) {
            g_ctx_tab_idx = idx;
            HMENU menu = CreatePopupMenu();
            AppendMenuW(menu, MF_STRING, IDM_TAB_RENAME, L"Rename...");
            AppendMenuW(menu, MF_STRING, IDM_TAB_CLOSE,  L"Close Tab");
            TrackPopupMenu(menu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, pt.x, pt.y, 0, h, nullptr);
            DestroyMenu(menu);
          }
          return 0;
        }
      }
      return 0;
    }

    case WM_APP_SESSION_ENDED:
      close_session((int)w, true);
      return 0;

    case WM_DESTROY:
      KillTimer(h, IDT_SIZE_POLL); KillTimer(h, IDT_CURSOR_BLINK);
      PostQuitMessage(0); return 0;
  }
  return DefWindowProcW(h, m, w, l);
}

// =====================================================
// WINDOW INIT
// =====================================================
static void init_window()
{
  INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_BAR_CLASSES | ICC_TAB_CLASSES };
  InitCommonControlsEx(&icc);

  WNDCLASSW wc = {};
  wc.lpfnWndProc   = WndProc;
  wc.hInstance     = GetModuleHandleW(NULL);
  wc.lpszClassName = L"MINITERM";
  wc.hbrBackground = (HBRUSH)GetStockObject(BLACK_BRUSH);
  RegisterClassW(&wc);

  // WS_CLIPCHILDREN is required here: without it, Direct2D's full-client
  // Clear()/Present() each frame paints straight over the tab control
  // (and statusbar/scrollbar), which then has to repaint itself on top --
  // a visible flicker cycle, worst at the tab strip since that's right
  // where the terminal's actively-redrawn content now renders.
  hwnd = CreateWindowW(L"MINITERM", L"miniterm",
    WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, 100, 100, 1000, 700,
    NULL, create_menu(), wc.hInstance, NULL);

  g_tabctrl = CreateWindowW(WC_TABCONTROLW, nullptr,
    WS_CHILD|WS_VISIBLE|TCS_FOCUSNEVER,
    0, 0, 0, 0, hwnd, (HMENU)(UINT_PTR)IDC_TABCTRL, wc.hInstance, nullptr);

  g_statusbar = CreateWindowW(L"msctls_statusbar32", nullptr,
    WS_CHILD|WS_VISIBLE|SBARS_SIZEGRIP,
    0, 0, 0, 0, hwnd, nullptr, wc.hInstance, nullptr);

  g_scrollbar = CreateWindowW(L"SCROLLBAR", nullptr,
    WS_CHILD|WS_VISIBLE|SBS_VERT,
    0, 0, GetSystemMetrics(SM_CXVSCROLL), 100,
    hwnd, nullptr, wc.hInstance, nullptr);

  ShowWindow(hwnd, SW_SHOW);
  SetTimer(hwnd, IDT_SIZE_POLL,    500, nullptr);
  SetTimer(hwnd, IDT_CURSOR_BLINK, 500, nullptr);
  update_statusbar();
}

// =====================================================
// DIRECTWRITE HELPERS
// =====================================================
static void measure_cell_dims()
{
  if (!fmt || !dw) return;
  IDWriteTextLayout* probe = nullptr;
  if (SUCCEEDED(dw->CreateTextLayout(
      L"MMMMMMMMMM", 10, fmt, 2000.0f, 100.0f, &probe)) && probe) {
    DWRITE_TEXT_METRICS m = {};
    if (SUCCEEDED(probe->GetMetrics(&m)) && m.width > 0.0f)
      CELL_W = m.width / 10.0f;
    probe->Release();
  }
  CELL_H = ceilf(g_font_size * 96.0f / 72.0f) + 2.0f;
  dbg("measure_cell_dims: CELL_W=%.4f CELL_H=%.4f font_size=%.1f\n",
      CELL_W, CELL_H, g_font_size);
}

static void update_fontface()
{
  if (g_fontface) { g_fontface->Release(); g_fontface = nullptr; }
  g_ref_advance = 0; g_design_units_per_em = 2048;
  if (!fmt || !dw) return;

  dbg("update_fontface: looking up font family...\n");

  IDWriteFontCollection* col = nullptr;
  fmt->GetFontCollection(&col);
  if (!col) {
    dbg("  fmt->GetFontCollection returned null, trying system collection\n");
    dw->GetSystemFontCollection(&col);
  }
  if (!col) { dbg("  ERROR: no font collection available\n"); return; }

  wchar_t fn[256] = {}; fmt->GetFontFamilyName(fn, 256);
  dbg("  family name from fmt: '%S'\n", fn);

  UINT32 idx = 0; BOOL ex = FALSE;
  col->FindFamilyName(fn, &idx, &ex);
  dbg("  FindFamilyName: found=%d idx=%u\n", (int)ex, idx);

  if (!ex) { col->Release(); return; }

  IDWriteFontFamily* family = nullptr;
  col->GetFontFamily(idx, &family);
  col->Release();
  if (!family) { dbg("  ERROR: GetFontFamily failed\n"); return; }

  IDWriteFont* font = nullptr;
  HRESULT hr = family->GetFirstMatchingFont(
    fmt->GetFontWeight(), fmt->GetFontStretch(), fmt->GetFontStyle(), &font);
  family->Release();
  dbg("  GetFirstMatchingFont hr=0x%08X font=%p\n", (unsigned)hr, (void*)font);
  if (!font) return;

  hr = font->CreateFontFace(&g_fontface);
  font->Release();
  dbg("  CreateFontFace hr=0x%08X fontface=%p\n", (unsigned)hr, (void*)g_fontface);
  if (!g_fontface) return;

  // Measure reference advance (glyph 'M')
  UINT32 cp = L'M'; UINT16 gi = 0;
  g_fontface->GetGlyphIndices(&cp, 1, &gi);
  DWRITE_GLYPH_METRICS gm = {};
  g_fontface->GetDesignGlyphMetrics(&gi, 1, &gm, FALSE);
  g_ref_advance = gm.advanceWidth;

  DWRITE_FONT_METRICS fm = {};
  g_fontface->GetMetrics(&fm);
  g_design_units_per_em = fm.designUnitsPerEm;

  float m_dip = du_to_dip(g_ref_advance);
  dbg("  ref_advance=%u designUnitsPerEm=%u M_dip=%.4f CELL_W=%.4f match=%s\n",
      g_ref_advance, g_design_units_per_em, m_dip, CELL_W,
      (fabsf(m_dip - CELL_W) < 0.5f) ? "YES" : "NO (unexpected)");

  clear_layout_cache();
}

static void rebuild_fmt(const std::wstring& face, float size)
{
  if (fmt) { fmt->Release(); fmt = nullptr; }
  g_font_name = face; g_font_size = size;

  dbg("rebuild_fmt: face='%S' size=%.1f fontcol=%p\n",
      face.c_str(), size, (void*)g_fontcol);

  HRESULT hr = dw->CreateTextFormat(face.c_str(), g_fontcol,
    DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL,
    DWRITE_FONT_STRETCH_NORMAL, size, L"en-us", &fmt);
  dbg("  CreateTextFormat hr=0x%08X fmt=%p\n", (unsigned)hr, (void*)fmt);

  if (fmt) { measure_cell_dims(); update_fontface(); }
}

// =====================================================
// DIRECTWRITE INIT — with full debug logging
// =====================================================
static void init_dw()
{
  dbg("init_dw: starting\n");

  HRESULT hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, &d2d);
  dbg("  D2D1CreateFactory hr=0x%08X d2d=%p\n", (unsigned)hr, (void*)d2d);

  hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED,
    __uuidof(IDWriteFactory), (IUnknown**)&dw);
  dbg("  DWriteCreateFactory hr=0x%08X dw=%p\n", (unsigned)hr, (void*)dw);

  // Check IDWriteFactory3 availability
  IDWriteFactory3* dw3_test = nullptr;
  hr = dw->QueryInterface(__uuidof(IDWriteFactory3), (void**)&dw3_test);
  dbg("  IDWriteFactory3 available: %s (hr=0x%08X)\n",
      SUCCEEDED(hr) ? "YES" : "NO", (unsigned)hr);
  if (dw3_test) { dw3_test->Release(); dw3_test = nullptr; }

  RECT rc; GetClientRect(hwnd, &rc);
  hr = d2d->CreateHwndRenderTarget(
    D2D1::RenderTargetProperties(),
    D2D1::HwndRenderTargetProperties(hwnd,
      D2D1::SizeU(rc.right-rc.left, rc.bottom-rc.top)),
    &rt);
  dbg("  CreateHwndRenderTarget hr=0x%08X rt=%p\n", (unsigned)hr, (void*)rt);

  // Probe font file paths
  const wchar_t* fps[] = {
    L"X:\\Windows\\Fonts\\JetBrainsMonoNerdFont-Regular.ttf",
    L"X:\\Windows\\Fonts\\JetBrainsMono-Regular.ttf",
    L"C:\\Windows\\Fonts\\JetBrainsMonoNerdFont-Regular.ttf",
    L"C:\\Windows\\Fonts\\JetBrainsMono-Regular.ttf",
    nullptr
  };
  const wchar_t* ff = nullptr;
  dbg("  Probing font file paths:\n");
  for (int i = 0; fps[i]; i++) {
    DWORD attr = GetFileAttributesW(fps[i]);
    bool exists = (attr != INVALID_FILE_ATTRIBUTES);
    dbg("    [%s] %S\n", exists ? "FOUND" : "missing", fps[i]);
    if (exists && !ff) ff = fps[i];
  }

  std::wstring rn = L"Courier New";

  if (ff) {
    dbg("  Using font file: %S\n", ff);
    IDWriteFactory3* dw3 = nullptr;
    hr = dw->QueryInterface(__uuidof(IDWriteFactory3), (void**)&dw3);
    if (SUCCEEDED(hr) && dw3) {
      dbg("  IDWriteFactory3 acquired\n");
      IDWriteFontSetBuilder* b = nullptr;
      hr = dw3->CreateFontSetBuilder(&b);
      dbg("  CreateFontSetBuilder hr=0x%08X b=%p\n", (unsigned)hr, (void*)b);
      if (SUCCEEDED(hr) && b) {
        IDWriteFontFile* fo = nullptr;
        hr = dw3->CreateFontFileReference(ff, nullptr, &fo);
        dbg("  CreateFontFileReference hr=0x%08X fo=%p\n", (unsigned)hr, (void*)fo);
        if (SUCCEEDED(hr) && fo) {
          BOOL sup = FALSE; DWRITE_FONT_FILE_TYPE ft; UINT32 fc = 0;
          hr = fo->Analyze(&sup, &ft, nullptr, &fc);
          dbg("  Analyze hr=0x%08X supported=%d fileType=%d faceCount=%u\n",
              (unsigned)hr, (int)sup, (int)ft, fc);
          if (sup) {
            for (UINT32 i = 0; i < fc; i++) {
              IDWriteFontFaceReference* ref = nullptr;
              hr = dw3->CreateFontFaceReference(ff, nullptr, i,
                DWRITE_FONT_SIMULATIONS_NONE, &ref);
              dbg("  CreateFontFaceReference[%u] hr=0x%08X\n", i, (unsigned)hr);
              if (ref) { b->AddFontFaceReference(ref); ref->Release(); }
            }
            IDWriteFontSet* fs = nullptr;
            hr = b->CreateFontSet(&fs);
            dbg("  CreateFontSet hr=0x%08X fs=%p\n", (unsigned)hr, (void*)fs);
            if (SUCCEEDED(hr) && fs) {
              IDWriteFontCollection1* col1 = nullptr;
              hr = dw3->CreateFontCollectionFromFontSet(fs, &col1);
              dbg("  CreateFontCollectionFromFontSet hr=0x%08X col1=%p\n", (unsigned)hr, (void*)col1);
              if (SUCCEEDED(hr) && col1) {
                g_fontcol = col1;
                UINT32 fcount = col1->GetFontFamilyCount();
                dbg("  Font collection has %u families\n", fcount);
                IDWriteFontFamily* fam = nullptr;
                if (SUCCEEDED(col1->GetFontFamily(0, &fam)) && fam) {
                  IDWriteLocalizedStrings* ns = nullptr;
                  if (SUCCEEDED(fam->GetFamilyNames(&ns)) && ns) {
                    wchar_t nb[256] = {}; ns->GetString(0, nb, 256);
                    dbg("  Family[0] name: '%S'\n", nb);
                    rn = nb; ns->Release();
                  }
                  fam->Release();
                }
              }
              fs->Release();
            }
          }
          fo->Release();
        }
        b->Release();
      }
      dw3->Release();
    } else {
      dbg("  WARNING: IDWriteFactory3 not available, falling back to system font\n");
    }
  } else {
    dbg("  WARNING: No Nerd Font file found, using Courier New\n");
  }

  // A persisted font choice (miniterm.json, or the in-session Font picker)
  // takes precedence over the auto-detected Nerd Font. The auto-detect
  // path above is untouched — this only branches for an explicit override.
  if (!g_config.font_name.empty()) {
    dbg("  Config font override: '%S' size=%.1f\n",
        g_config.font_name.c_str(), g_config.font_size);
    rebuild_fmt(g_config.font_name, g_config.font_size);
  }
  if (!fmt) {
    dbg("  Calling rebuild_fmt with '%S'\n", rn.c_str());
    rebuild_fmt(rn, 16.0f);
  }
  if (!fmt) {
    dbg("  rebuild_fmt failed, falling back to Courier New\n");
    rebuild_fmt(L"Courier New", 16.0f);
  }

  dbg("init_dw complete: CELL_W=%.4f CELL_H=%.4f fontface=%p ref_advance=%u\n",
      CELL_W, CELL_H, (void*)g_fontface, g_ref_advance);

  pty_resize();
}

// =====================================================
// SHELL LAUNCH
// =====================================================
static bool launch_shell()
{
  std::wstring res = resolve_exe(g_exe);
  if (res.empty()) {
    EnterCriticalSection(&g_cs);
    grid_put_string(L"[miniterm] ERROR: cannot resolve executable:");
    grid_put_string(g_exe.c_str());
    LeaveCriticalSection(&g_cs);
    InvalidateRect(hwnd, NULL, FALSE); return false;
  }

  CreatePipe(&hInR, &hInW, NULL, 0);
  CreatePipe(&hOutR, &hOutW, NULL, 0);

  COORD sz = {(SHORT)W,(SHORT)H};
  HRESULT hr = CreatePseudoConsole(sz, hInR, hOutW, 0, &g_hpc);
  if (FAILED(hr)) {
    wchar_t buf[64]; wsprintfW(buf, L"[miniterm] CreatePseudoConsole: 0x%08X", (unsigned)hr);
    EnterCriticalSection(&g_cs); grid_put_string(buf); LeaveCriticalSection(&g_cs);
    InvalidateRect(hwnd, NULL, FALSE); return false;
  }

  SIZE_T s = 0; InitializeProcThreadAttributeList(NULL, 1, 0, &s);
  STARTUPINFOEXW si = {}; si.StartupInfo.cb = sizeof(si);
  si.lpAttributeList = (PPROC_THREAD_ATTRIBUTE_LIST)HeapAlloc(GetProcessHeap(), 0, s);
  InitializeProcThreadAttributeList(si.lpAttributeList, 1, 0, &s);
  UpdateProcThreadAttribute(si.lpAttributeList, 0,
    PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE, g_hpc, sizeof(g_hpc), NULL, NULL);

  // res must be quoted: an unquoted path containing spaces (e.g. under
  // "C:\Program Files\...") gets split by the child's own argv parser,
  // which then misreads the tail of its own exe path as a stray argument.
  std::wstring cmd = L"\"" + res + L"\" " + g_args;
  std::vector<wchar_t> mc(cmd.begin(), cmd.end()); mc.push_back(0);
  PROCESS_INFORMATION pi = {};
  BOOL ok = CreateProcessW(res.c_str(), mc.data(), NULL, NULL, FALSE,
    EXTENDED_STARTUPINFO_PRESENT, NULL, NULL, &si.StartupInfo, &pi);

  DeleteProcThreadAttributeList(si.lpAttributeList);
  HeapFree(GetProcessHeap(), 0, si.lpAttributeList);

  if (!ok) {
    DWORD err = GetLastError(); wchar_t buf[128];
    wsprintfW(buf, L"[miniterm] CreateProcessW: GLE=0x%08X", err);
    EnterCriticalSection(&g_cs); grid_put_string(buf); LeaveCriticalSection(&g_cs);
    InvalidateRect(hwnd, NULL, FALSE); return false;
  }
  CloseHandle(hInR); CloseHandle(hOutW);
  // Keep the process handles (g_child_pi) so switch_shell() can terminate
  // this child cleanly when the user picks a different shell.
  g_child_pi = pi;
  return true;
}

// =====================================================
// READER THREAD
// =====================================================
DWORD WINAPI reader(LPVOID)
{
  char buf[8192]; DWORD r;
  Utf8Decoder utf8; VtProcessor vt;
  while (ReadFile(hOutR, buf, sizeof(buf), &r, NULL) && r) {
    EnterCriticalSection(&g_cs);
    for (DWORD i = 0; i < r; i++) {
      uint32_t cp = 0; if (utf8.feed((uint8_t)buf[i], cp)) vt.feed(cp);
    }
    if (g_at_bottom) g_view_top = (int)g_history.size();
    LeaveCriticalSection(&g_cs);
    update_scrollbar(); update_statusbar(); InvalidateRect(hwnd, NULL, FALSE);
  }
  // Four ways this loop can end:
  //  1. Deliberate shell switch (switch_shell) closed the pipe on purpose.
  //  2. Deliberate tab backgrounding (session_pause_reader) cancelled the
  //     pending ReadFile via CancelIoEx without closing anything -- the
  //     handle stays valid so a later reader thread can resume from it.
  //  3. The child process exited and ConPTY promptly signalled EOF on
  //     its own (ReadFile fails with something other than the operator-
  //     cancelled error below).
  //  4. The child process exited but ConPTY did NOT close the pipe on
  //     its own -- the WM_TIMER poll noticed via the process handle and
  //     used the same CancelIoEx mechanism as case 2 to unblock us, but
  //     set g_watcher_signals_exit first so we can tell it apart from a
  //     genuine pause/switch.
  // Cases 3 and 4 should tell the UI thread a session died; 1 and 2 must not.
  bool watcher_exit = g_watcher_signals_exit;
  g_watcher_signals_exit = false;
  if (g_shell_switching) return 0;
  if (GetLastError() == ERROR_OPERATION_ABORTED && !watcher_exit) return 0;
  PostMessageW(hwnd, WM_APP_SESSION_ENDED, (WPARAM)g_active_session, 0);
  return 0;
}

// =====================================================
// SHELL SWITCH — tears down the running child and PTY, then
// launches a different preset in the same window. Full multi-shell
// (several shells open at once) arrives with per-tab sessions; this
// replaces the single current session's shell.
// =====================================================
static void switch_shell(const ShellPreset& sp)
{
  if (sp.name == g_shellname) return;

  g_shell_switching = true;

  if (g_hpc) { ClosePseudoConsole(g_hpc); g_hpc = nullptr; }
  if (g_child_pi.hProcess) {
    TerminateProcess(g_child_pi.hProcess, 0);
    WaitForSingleObject(g_child_pi.hProcess, 2000);
    CloseHandle(g_child_pi.hProcess);
    CloseHandle(g_child_pi.hThread);
    g_child_pi = PROCESS_INFORMATION{};
  }
  if (hInW)  { CloseHandle(hInW);  hInW  = nullptr; }
  if (hOutR) { CloseHandle(hOutR); hOutR = nullptr; }  // unblocks reader()'s ReadFile

  if (g_reader_thread) {
    WaitForSingleObject(g_reader_thread, 2000);
    CloseHandle(g_reader_thread);
    g_reader_thread = nullptr;
  }

  g_exe = sp.exe; g_args = sp.args; g_shellname = sp.name;

  EnterCriticalSection(&g_cs);
  screen_clear_region(0, 0, W - 1, H - 1);
  cx = 0; cy = 0;
  g_history.clear(); g_view_top = 0; g_at_bottom = true;
  LeaveCriticalSection(&g_cs);

  g_shell_switching = false;

  if (launch_shell())
    g_reader_thread = CreateThread(NULL, 0, reader, NULL, 0, NULL);

  if (g_shell_menu) {
    for (size_t i = 0; i < g_config.shells.size() && i < IDM_SHELL_MAX; i++)
      CheckMenuItem(g_shell_menu, IDM_SHELL_BASE + (UINT)i, MF_BYCOMMAND |
        ((g_config.shells[i].name == g_shellname) ? MF_CHECKED : MF_UNCHECKED));
  }

  update_statusbar();
  InvalidateRect(hwnd, NULL, FALSE);
  config_save_now();
}

// =====================================================
// SESSION CAPTURE / RESTORE
// Move the live globals into/out of a Session struct. See the
// "SESSIONS (tabs)" comment near the Session struct definition for
// why this swap approach was chosen over threading session state
// through every function.
// =====================================================
static void session_capture(Session& s)
{
  s.screen = g_screen; s.history = g_history; s.row_max_cx = g_row_max_cx;
  s.view_top = g_view_top; s.at_bottom = g_at_bottom;
  s.cx = cx; s.cy = cy; s.saved_cx = saved_cx; s.saved_cy = saved_cy;
  s.cursor_visible = g_cursor_visible; s.cur_fg = g_cur_fg; s.cur_bg = g_cur_bg;
  s.sel_active = g_sel_active; s.sel_dragging = g_sel_dragging;
  s.sel_r0 = g_sel_r0; s.sel_c0 = g_sel_c0; s.sel_r1 = g_sel_r1; s.sel_c1 = g_sel_c1;
  s.hpc = g_hpc; s.hInR = hInR; s.hInW = hInW; s.hOutR = hOutR; s.hOutW = hOutW;
  s.child_pi = g_child_pi; s.reader_thread = g_reader_thread;
  s.exe = g_exe; s.args = g_args; s.shellname = g_shellname;
}

static void session_restore(const Session& s)
{
  g_screen = s.screen; g_history = s.history; g_row_max_cx = s.row_max_cx;
  g_view_top = s.view_top; g_at_bottom = s.at_bottom;
  cx = s.cx; cy = s.cy; saved_cx = s.saved_cx; saved_cy = s.saved_cy;
  g_cursor_visible = s.cursor_visible; g_cur_fg = s.cur_fg; g_cur_bg = s.cur_bg;
  g_sel_active = s.sel_active; g_sel_dragging = s.sel_dragging;
  g_sel_r0 = s.sel_r0; g_sel_c0 = s.sel_c0; g_sel_r1 = s.sel_r1; g_sel_c1 = s.sel_c1;
  g_hpc = s.hpc; hInR = s.hInR; hInW = s.hInW; hOutR = s.hOutR; hOutW = s.hOutW;
  g_child_pi = s.child_pi; g_reader_thread = s.reader_thread;
  g_exe = s.exe; g_args = s.args; g_shellname = s.shellname;
}

// Cancels the active session's pending ReadFile without closing hOutR,
// so its backlog keeps accumulating in the OS pipe and a later reader
// thread can resume from the same handle.
static void session_pause_reader()
{
  if (hOutR) CancelIoEx(hOutR, NULL);
  if (g_reader_thread) {
    WaitForSingleObject(g_reader_thread, 2000);
    CloseHandle(g_reader_thread);
    g_reader_thread = nullptr;
  }
}

static void session_resume_reader()
{
  if (hOutR && !g_reader_thread)
    g_reader_thread = CreateThread(NULL, 0, reader, NULL, 0, NULL);
}

// =====================================================
// TAB STRIP UI
// =====================================================
static void rebuild_tab_ui()
{
  if (!g_tabctrl) return;
  SendMessageW(g_tabctrl, TCM_DELETEALLITEMS, 0, 0);
  for (size_t i = 0; i < g_sessions.size(); i++) {
    TCITEMW ti = {};
    ti.mask = TCIF_TEXT;
    ti.pszText = (LPWSTR)g_sessions[i]->name.c_str();
    SendMessageW(g_tabctrl, TCM_INSERTITEMW, i, (LPARAM)&ti);
  }
  if (g_active_session >= 0)
    SendMessageW(g_tabctrl, TCM_SETCURSEL, (WPARAM)g_active_session, 0);
}

// =====================================================
// SWITCH / CREATE / CLOSE SESSIONS (tabs)
// =====================================================
static void switch_to_session(int idx)
{
  if (idx < 0 || idx >= (int)g_sessions.size() || idx == g_active_session) return;

  if (g_active_session >= 0) {
    session_pause_reader();
    session_capture(*g_sessions[g_active_session]);
  }

  g_active_session = idx;
  session_restore(*g_sessions[idx]);

  // The window may have been resized while this session was backgrounded
  // (only the then-active session's grid/pty got resized at the time) --
  // reconcile this session's screen buffer and PTY to the current W x H.
  EnterCriticalSection(&g_cs);
  bool stale = g_screen.empty() || (int)g_screen.size() != H ||
               (int)g_screen[0].size() != W;
  if (stale) grid_resize(W, H);
  LeaveCriticalSection(&g_cs);
  if (g_hpc) { COORD sz = {(SHORT)W, (SHORT)H}; ResizePseudoConsole(g_hpc, sz); }

  session_resume_reader();

  if (g_tabctrl) SendMessageW(g_tabctrl, TCM_SETCURSEL, (WPARAM)idx, 0);
  update_statusbar();
  update_scrollbar();
  InvalidateRect(hwnd, NULL, FALSE);
}

static void new_session(const ShellPreset& sp)
{
  auto s = std::make_unique<Session>();
  wchar_t nb[64];
  wsprintfW(nb, L"%s #%d", sp.name.c_str(), g_next_session_num++);
  s->name      = nb;
  s->exe       = sp.exe;
  s->args      = sp.args;
  s->shellname = sp.name;
  s->screen.assign(H, std::vector<Cell>(W, Cell{ ' ', theme_fg(), theme_bg(), false }));
  s->row_max_cx.assign(H, 0);

  g_sessions.push_back(std::move(s));
  int new_idx = (int)g_sessions.size() - 1;

  if (g_active_session >= 0) {
    session_pause_reader();
    session_capture(*g_sessions[g_active_session]);
  }

  g_active_session = new_idx;
  session_restore(*g_sessions[new_idx]);

  if (launch_shell())
    g_reader_thread = CreateThread(NULL, 0, reader, NULL, 0, NULL);

  rebuild_tab_ui();
  update_statusbar();
  InvalidateRect(hwnd, NULL, FALSE);
}

// already_dead: true when called because the child process exited on its
// own (reader thread already returned) -- false when the user explicitly
// asked to close a still-running tab, which must tear the shell down.
static void close_session(int idx, bool already_dead)
{
  if (idx < 0 || idx >= (int)g_sessions.size()) return;

  if (idx == g_active_session) {
    if (!already_dead) {
      if (g_reader_thread) {
        if (hOutR) CancelIoEx(hOutR, NULL);
        WaitForSingleObject(g_reader_thread, 2000);
        CloseHandle(g_reader_thread);
        g_reader_thread = nullptr;
      }
      if (g_child_pi.hProcess) {
        TerminateProcess(g_child_pi.hProcess, 0);
        WaitForSingleObject(g_child_pi.hProcess, 2000);
      }
    }
    // Regardless of already_dead, the pseudoconsole handle itself is done
    // and must be released (ClosePseudoConsole is safe to call whether or
    // not the hosted process has already exited).
    if (g_hpc) { ClosePseudoConsole(g_hpc); g_hpc = nullptr; }
    if (g_child_pi.hProcess) { CloseHandle(g_child_pi.hProcess); CloseHandle(g_child_pi.hThread); }
    g_child_pi = PROCESS_INFORMATION{};
    if (hInW)  { CloseHandle(hInW);  hInW  = nullptr; }
    if (hOutR) { CloseHandle(hOutR); hOutR = nullptr; }
    if (g_reader_thread) { CloseHandle(g_reader_thread); g_reader_thread = nullptr; }
  } else {
    // Background tab: no live reader thread to stop, just release its
    // handles and (if still running) its child process.
    Session& s = *g_sessions[idx];
    if (s.child_pi.hProcess) {
      TerminateProcess(s.child_pi.hProcess, 0);
      WaitForSingleObject(s.child_pi.hProcess, 2000);
      CloseHandle(s.child_pi.hProcess);
      CloseHandle(s.child_pi.hThread);
    }
    if (s.hInW)  CloseHandle(s.hInW);
    if (s.hOutR) CloseHandle(s.hOutR);
    if (s.hpc)   ClosePseudoConsole(s.hpc);
  }

  bool was_active = (idx == g_active_session);
  g_sessions.erase(g_sessions.begin() + idx);

  if (g_sessions.empty()) { DestroyWindow(hwnd); return; }

  if (was_active) {
    g_active_session = -1;
    switch_to_session(min(idx, (int)g_sessions.size() - 1));
  } else if (idx < g_active_session) {
    g_active_session--;
  }

  rebuild_tab_ui();
}

// =====================================================
// MAIN
// =====================================================
int wmain(int argc, wchar_t** argv)
{
  // File-based defaults load first; CLI flags (parse_args, below) override
  // them, matching normal precedence of persisted config < explicit flags.
  config_load(g_config);
  g_theme          = str_to_theme(g_config.theme);
  SCROLLBACK_LINES = g_config.scrollback_lines;
  for (const auto& sp : g_config.shells) {
    if (sp.name == g_config.default_shell) {
      g_exe = sp.exe; g_args = sp.args; g_shellname = sp.name;
      break;
    }
  }

  parse_args(argc, argv);
  debug_open();

  InitializeCriticalSection(&g_cs);
  InitializeCriticalSection(&g_cache_cs);
  build_palette(g_theme);

  CoInitialize(NULL);
  init_window();
  init_dw();

  // The first tab, built from whatever shell config/CLI resolved above --
  // new_session() establishes the screen/history/cursor state itself, so
  // no separate manual g_screen init is needed here.
  new_session(ShellPreset{ g_shellname, g_exe, g_args });

  MSG msg;
  while (GetMessageW(&msg, NULL, 0, 0)) {
    TranslateMessage(&msg); DispatchMessageW(&msg);
  }

  config_save_now();

  // The active session's real handles live in the bare globals below, not
  // in its Session struct (that copy is only refreshed by session_capture,
  // which runs on backgrounding -- for whichever tab is active right now
  // its struct fields are a stale echo of a past state, not a live alias,
  // EXCEPT the values happen to still match if it was ever backgrounded
  // and reactivated. Skip it here to avoid closing the same handle twice.
  for (int i = 0; i < (int)g_sessions.size(); i++) {
    if (i == g_active_session) continue;
    Session& s = *g_sessions[i];
    if (s.child_pi.hProcess) { TerminateProcess(s.child_pi.hProcess, 0); CloseHandle(s.child_pi.hProcess); CloseHandle(s.child_pi.hThread); }
    if (s.hInW)  CloseHandle(s.hInW);
    if (s.hOutR) CloseHandle(s.hOutR);
    if (s.hpc)   ClosePseudoConsole(s.hpc);
  }
  if (g_child_pi.hProcess) { TerminateProcess(g_child_pi.hProcess, 0); CloseHandle(g_child_pi.hProcess); CloseHandle(g_child_pi.hThread); }
  if (hInW)  CloseHandle(hInW);
  if (hOutR) CloseHandle(hOutR);
  if (g_hpc) ClosePseudoConsole(g_hpc);
  if (g_reader_thread) { CloseHandle(g_reader_thread); g_reader_thread = nullptr; }

  clear_layout_cache();
  if (g_fontface) { g_fontface->Release(); g_fontface = nullptr; }
  if (g_fontcol)  { g_fontcol->Release();  g_fontcol  = nullptr; }
  if (fmt)        { fmt->Release();         fmt        = nullptr; }
  if (dw)         { dw->Release();          dw         = nullptr; }
  if (d2d)        { d2d->Release();         d2d        = nullptr; }
  if (g_debug_file) { fclose(g_debug_file); g_debug_file = nullptr; }

  DeleteCriticalSection(&g_cache_cs);
  DeleteCriticalSection(&g_cs);
  return 0;
}

