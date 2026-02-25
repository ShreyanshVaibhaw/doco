use std::time::{Duration, Instant};

use crate::render::image_cache::ImageCacheStats;

#[cfg(all(feature = "profiling", target_os = "windows"))]
use std::sync::OnceLock;

#[derive(Debug, Clone, Copy, Default)]
pub struct PerformanceSnapshot {
    pub fps: f32,
    pub frame_time_ms: f32,
    pub process_memory_mb: f32,
    pub image_cache_hit_rate: f32,
    pub image_cache_mb: f32,
}

#[derive(Debug, Clone)]
pub struct DebugPerformancePanel {
    pub visible: bool,
    pub snapshot: PerformanceSnapshot,
    frame_window_start: Instant,
    frame_count: u32,
}

impl Default for DebugPerformancePanel {
    fn default() -> Self {
        Self {
            visible: false,
            snapshot: PerformanceSnapshot::default(),
            frame_window_start: Instant::now(),
            frame_count: 0,
        }
    }
}

impl DebugPerformancePanel {
    pub fn toggle(&mut self) {
        self.visible = !self.visible;
    }

    pub fn set_visible(&mut self, visible: bool) {
        self.visible = visible;
    }

    pub fn update_frame_time(&mut self, frame_time_ms: f32) {
        self.snapshot.frame_time_ms = frame_time_ms.max(0.0);
        self.frame_count += 1;

        let elapsed = self.frame_window_start.elapsed();
        if elapsed >= Duration::from_secs(1) {
            let secs = elapsed.as_secs_f32().max(0.001);
            self.snapshot.fps = self.frame_count as f32 / secs;
            self.frame_count = 0;
            self.frame_window_start = Instant::now();
        }
    }

    pub fn update_memory_bytes(&mut self, bytes: u64) {
        self.snapshot.process_memory_mb = bytes as f32 / (1024.0 * 1024.0);
    }

    pub fn update_image_cache_stats(&mut self, stats: ImageCacheStats) {
        self.snapshot.image_cache_hit_rate = stats.hit_rate();
        self.snapshot.image_cache_mb =
            (stats.full_res_bytes + stats.thumbnail_bytes) as f32 / (1024.0 * 1024.0);
    }
}

pub fn emit_startup_marker(stage: &str, elapsed_ms: f64) {
    #[cfg(feature = "profiling")]
    {
        let message = format!("startup.{stage} {:.3} ms", elapsed_ms);
        eprintln!("[profiling] {message}");
        emit_etw_marker(message.as_str());
    }

    #[cfg(not(feature = "profiling"))]
    {
        let _ = (stage, elapsed_ms);
    }
}

#[cfg(feature = "profiling")]
pub struct ProfileGuard {
    name: &'static str,
    start: Instant,
}

#[cfg(feature = "profiling")]
impl ProfileGuard {
    pub fn new(name: &'static str) -> Self {
        emit_etw_marker(&format!("begin:{name}"));
        Self {
            name,
            start: Instant::now(),
        }
    }
}

#[cfg(feature = "profiling")]
impl Drop for ProfileGuard {
    fn drop(&mut self) {
        let elapsed_ms = self.start.elapsed().as_secs_f64() * 1000.0;
        eprintln!("[profiling] {} took {:.3} ms", self.name, elapsed_ms);
        emit_etw_marker(&format!("end:{} {:.3} ms", self.name, elapsed_ms));
    }
}

#[cfg(not(feature = "profiling"))]
pub struct ProfileGuard;

#[cfg(not(feature = "profiling"))]
impl ProfileGuard {
    pub fn new(_name: &'static str) -> Self {
        Self
    }
}

#[cfg(all(feature = "profiling", target_os = "windows"))]
fn etw_provider_handle() -> Option<windows::Win32::System::Diagnostics::Etw::REGHANDLE> {
    use windows::Win32::System::Diagnostics::Etw::{EventRegister, REGHANDLE};
    use windows::core::GUID;

    static HANDLE: OnceLock<Option<REGHANDLE>> = OnceLock::new();
    *HANDLE.get_or_init(|| {
        // Stable provider GUID dedicated to Doco profiling markers.
        let provider = GUID::from_u128(0x6a77f9a8_4b1f_4f63_8c9a_3d3b4df20501);
        let mut handle = REGHANDLE::default();
        let status = unsafe { EventRegister(&provider, None, None, &mut handle) };
        if status == 0 { Some(handle) } else { None }
    })
}

#[cfg(all(feature = "profiling", target_os = "windows"))]
pub fn emit_etw_marker(message: &str) {
    use windows::Win32::System::Diagnostics::Etw::EventWriteString;
    use windows::core::PCWSTR;

    let Some(handle) = etw_provider_handle() else {
        return;
    };
    let wide = message
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect::<Vec<u16>>();
    unsafe {
        let _ = EventWriteString(handle, 0, 0, PCWSTR(wide.as_ptr()));
    }
}

#[cfg(not(all(feature = "profiling", target_os = "windows")))]
pub fn emit_etw_marker(_message: &str) {}

#[cfg(target_os = "windows")]
pub fn query_process_working_set_bytes() -> Option<u64> {
    use windows::Win32::System::{
        ProcessStatus::{K32GetProcessMemoryInfo, PROCESS_MEMORY_COUNTERS},
        Threading::GetCurrentProcess,
    };

    let process = unsafe { GetCurrentProcess() };
    let mut counters = PROCESS_MEMORY_COUNTERS::default();
    let ok = unsafe {
        K32GetProcessMemoryInfo(
            process,
            &mut counters,
            std::mem::size_of::<PROCESS_MEMORY_COUNTERS>() as u32,
        )
        .as_bool()
    };

    if ok {
        Some(counters.WorkingSetSize as u64)
    } else {
        None
    }
}

#[cfg(not(target_os = "windows"))]
pub fn query_process_working_set_bytes() -> Option<u64> {
    None
}

#[macro_export]
macro_rules! profile_scope {
    ($name:expr) => {
        let _profile_guard = $crate::render::perf::ProfileGuard::new($name);
    };
}
