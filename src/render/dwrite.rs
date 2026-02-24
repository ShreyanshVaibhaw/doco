use windows::{
    Win32::Graphics::DirectWrite::{DWRITE_FACTORY_TYPE_SHARED, DWriteCreateFactory, IDWriteFactory},
    core::Result,
};

pub fn create_factory() -> Result<IDWriteFactory> {
    unsafe { DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED) }
}
