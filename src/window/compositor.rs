use std::time::{Duration, Instant};

#[derive(Debug, Clone)]
pub struct Compositor {
    pub enabled: bool,
    pub target_frame_time: Duration,
    last_present: Instant,
}

impl Default for Compositor {
    fn default() -> Self {
        Self {
            enabled: true,
            target_frame_time: Duration::from_millis(16),
            last_present: Instant::now(),
        }
    }
}

impl Compositor {
    pub fn begin_frame(&mut self) -> Instant {
        Instant::now()
    }

    pub fn end_frame(&mut self, frame_start: Instant) -> Duration {
        let frame_duration = frame_start.elapsed();
        self.last_present = Instant::now();
        frame_duration
    }

    pub fn should_throttle(&self) -> bool {
        self.enabled && self.last_present.elapsed() < self.target_frame_time
    }
}
