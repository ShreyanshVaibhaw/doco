#[derive(Debug, Clone, Copy)]
pub enum Easing {
    Linear,
    EaseOutCubic,
    Spring,
}

#[derive(Debug, Clone, Copy)]
pub struct Animation {
    pub start_value: f32,
    pub end_value: f32,
    pub current_value: f32,
    pub duration: f32,
    pub elapsed: f32,
    pub easing: Easing,
    pub velocity: f32,
}

impl Animation {
    pub fn new(start_value: f32, end_value: f32, duration: f32, easing: Easing) -> Self {
        Self {
            start_value,
            end_value,
            current_value: start_value,
            duration: duration.max(0.0001),
            elapsed: 0.0,
            easing,
            velocity: 0.0,
        }
    }

    pub fn reset(&mut self, start: f32, end: f32, duration: f32, easing: Easing) {
        self.start_value = start;
        self.end_value = end;
        self.current_value = start;
        self.duration = duration.max(0.0001);
        self.elapsed = 0.0;
        self.easing = easing;
        self.velocity = 0.0;
    }

    pub fn update(&mut self, dt: f32) -> bool {
        self.elapsed += dt.max(0.0);

        match self.easing {
            Easing::Linear => {
                let t = (self.elapsed / self.duration).clamp(0.0, 1.0);
                self.current_value = lerp(self.start_value, self.end_value, t);
            }
            Easing::EaseOutCubic => {
                let t = (self.elapsed / self.duration).clamp(0.0, 1.0);
                let eased = 1.0 - (1.0 - t).powi(3);
                self.current_value = lerp(self.start_value, self.end_value, eased);
            }
            Easing::Spring => {
                let stiffness = 240.0;
                let damping = 28.0;
                let x = self.current_value - self.end_value;
                let accel = -stiffness * x - damping * self.velocity;
                self.velocity += accel * dt;
                self.current_value += self.velocity * dt;
            }
        }

        if self.easing == Easing::Spring {
            let done = (self.current_value - self.end_value).abs() < 0.001 && self.velocity.abs() < 0.001;
            if done {
                self.current_value = self.end_value;
            }
            !done
        } else {
            self.elapsed < self.duration
        }
    }

    pub fn update_respecting_motion_pref(&mut self, dt: f32, reduce_motion: bool) -> bool {
        if reduce_motion {
            self.current_value = self.end_value;
            self.elapsed = self.duration;
            self.velocity = 0.0;
            return false;
        }

        self.update(dt)
    }
}

impl PartialEq for Easing {
    fn eq(&self, other: &Self) -> bool {
        matches!((self, other),
            (Easing::Linear, Easing::Linear)
            | (Easing::EaseOutCubic, Easing::EaseOutCubic)
            | (Easing::Spring, Easing::Spring)
        )
    }
}

fn lerp(a: f32, b: f32, t: f32) -> f32 {
    a + (b - a) * t
}
