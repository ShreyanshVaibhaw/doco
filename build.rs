use std::path::Path;

fn main() {
    println!("cargo:rerun-if-changed=resources/app.manifest");
    println!("cargo:rerun-if-changed=resources/app.rc");
    println!("cargo:rerun-if-changed=assets/icons/app.ico");

    if cfg!(target_os = "windows") {
        let mut res = winres::WindowsResource::new();
        res.set_manifest_file("resources/app.manifest");

        if Path::new("assets/icons/app.ico").exists() {
            res.set_icon("assets/icons/app.ico");
        }

        if let Err(error) = res.compile() {
            panic!("failed to compile Windows resources: {error}");
        }
    }
}
