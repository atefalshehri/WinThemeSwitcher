//! Build script: gives the exe an application manifest (`app.manifest`) and
//! a version resource, both carrying the version from `Cargo.toml`.
//!
//! It writes a compiled resource file (`.res`) into `OUT_DIR` by hand and
//! hands it to the linker, which accepts `.res` inputs directly — no crate,
//! no `rc.exe`, no `mt.exe`. The `.res` goes into every build of the bin
//! target, the unit-test executable included, so each test run also shows
//! that an exe with this manifest and the `.cargo\config.toml` linker flags
//! loads and runs (`mod tests` in `src/main.rs` checks what Windows made of
//! both resources).
//!
//! Formats: "Resource File Formats" (RESOURCEHEADER) and "VS_VERSIONINFO"
//! in the Win32 documentation. Every header and every block starts on a
//! 32-bit boundary.

use std::env;
use std::fs;
use std::path::Path;

/// Resource types (`RT_VERSION`, `RT_MANIFEST`), and the id Windows looks
/// under for each: `VS_VERSION_INFO` and `CREATEPROCESS_MANIFEST_RESOURCE_ID`
/// are both 1.
const RT_VERSION: u16 = 16;
const RT_MANIFEST: u16 = 24;
const RESOURCE_ID: u16 = 1;
/// U.S. English, with the Unicode code page (1200) for the strings.
const LANG_EN_US: u16 = 0x0409;
const CODEPAGE_UNICODE: u16 = 0x04B0;

const PRODUCT: &str = "WinThemeSwitcher";
const EXE_NAME: &str = "win-theme-switcher.exe";
/// What `app.manifest` has where the four-part version goes.
const VERSION_PLACEHOLDER: &str = "{VERSION}";
/// Settings the manifest must never gain: they change how Windows APIs
/// behave (see the comment in `app.manifest`).
const FORBIDDEN_IN_MANIFEST: [&str; 2] = ["longpathaware", "activecodepage"];

fn main() {
    // Any edit to Cargo.toml (a version bump above all) re-runs this script,
    // so the embedded version can't go stale; scripts\check-hardening.ps1
    // compares the built exe's version with Cargo.toml as the backstop.
    for file in ["build.rs", "app.manifest", "LICENSE", "Cargo.toml"] {
        println!("cargo:rerun-if-changed={file}");
    }

    // link.exe (and lld-link) take a .res as an input; GNU ld does not.
    if var("CARGO_CFG_TARGET_OS") != "windows" || var("CARGO_CFG_TARGET_ENV") != "msvc" {
        println!(
            "cargo:warning=not an MSVC Windows target: the exe gets no manifest and no version resource"
        );
        return;
    }

    let root = var("CARGO_MANIFEST_DIR");
    let root = Path::new(&root);
    let version = var("CARGO_PKG_VERSION");
    let numbers = [
        version_part("CARGO_PKG_VERSION_MAJOR"),
        version_part("CARGO_PKG_VERSION_MINOR"),
        version_part("CARGO_PKG_VERSION_PATCH"),
        0,
    ];

    let manifest = manifest_xml(&read(&root.join("app.manifest")), numbers);
    let (copyright, holder) = copyright_line(&read(&root.join("LICENSE")));
    let version_info = version_info(numbers, &version, &copyright, &holder);

    // A .res starts with an empty entry that marks it as the 32-bit format.
    let mut res = Vec::new();
    res_entry(&mut res, 0, 0, 0, 0, &[]);
    res_entry(
        &mut res,
        RT_VERSION,
        RESOURCE_ID,
        0x0030,
        LANG_EN_US,
        &version_info,
    );
    res_entry(
        &mut res,
        RT_MANIFEST,
        RESOURCE_ID,
        0x0030,
        LANG_EN_US,
        manifest.as_bytes(),
    );

    let out = Path::new(&var("OUT_DIR")).join("resources.res");
    if let Err(e) = fs::write(&out, res) {
        fail(&format!("cannot write {}: {e}", out.display()));
    }
    // "-bins": this package's bin target, in every profile (cargo applies it
    // to the bin's unit-test executable too).
    println!("cargo:rustc-link-arg-bins={}", out.display());
}

fn fail(msg: &str) -> ! {
    eprintln!("error: build.rs: {msg}");
    std::process::exit(1);
}

fn var(name: &str) -> String {
    env::var(name).unwrap_or_else(|_| fail(&format!("{name} is not set (run through cargo)")))
}

fn read(path: &Path) -> String {
    fs::read_to_string(path)
        .unwrap_or_else(|e| fail(&format!("cannot read {}: {e}", path.display())))
}

/// One part of the package version; a version resource holds 16 bits each.
fn version_part(name: &str) -> u16 {
    let text = var(name);
    text.parse().unwrap_or_else(|_| {
        fail(&format!(
            "{name}={text}: a version resource takes numbers up to 65535"
        ))
    })
}

/// `app.manifest` as it is embedded: comments and blank lines removed, LF
/// line ends (so a CRLF checkout embeds the same bytes), and the placeholder
/// replaced with the four-part version.
fn manifest_xml(source: &str, numbers: [u16; 4]) -> String {
    let mut text = String::new();
    let mut rest = source;
    while let Some(start) = rest.find("<!--") {
        text.push_str(&rest[..start]);
        match rest[start..].find("-->") {
            Some(end) => rest = &rest[start + end + 3..],
            None => fail("app.manifest: a comment is never closed"),
        }
    }
    text.push_str(rest);

    let lower = text.to_ascii_lowercase();
    for word in FORBIDDEN_IN_MANIFEST {
        if lower.contains(word) {
            fail(&format!(
                "app.manifest must not set {word}: it changes how Windows APIs behave for the app"
            ));
        }
    }
    if text.matches(VERSION_PLACEHOLDER).count() != 1 {
        fail(&format!(
            "app.manifest must contain {VERSION_PLACEHOLDER} exactly once (the assemblyIdentity version)"
        ));
    }
    let [major, minor, patch, build] = numbers;
    let text = text.replace(
        VERSION_PLACEHOLDER,
        &format!("{major}.{minor}.{patch}.{build}"),
    );

    let mut xml = String::new();
    for line in text.lines().map(str::trim_end) {
        if !line.trim_start().is_empty() {
            xml.push_str(line);
            xml.push('\n');
        }
    }
    xml
}

/// LICENSE's copyright line (`Copyright (c) <year> <holder>`) and the holder
/// named in it.
fn copyright_line(license: &str) -> (String, String) {
    const PREFIX: &str = "Copyright (c) ";
    let line = license
        .lines()
        .map(str::trim)
        .find(|l| l.starts_with(PREFIX))
        .unwrap_or_else(|| fail("LICENSE has no line starting with \"Copyright (c) \""));
    // After the prefix: the year (or years, "2026-2027"), then the holder.
    let holder = line[PREFIX.len()..]
        .trim_start_matches(|c: char| c.is_ascii_digit() || c == '-' || c == ',')
        .trim();
    if holder.is_empty() || holder.len() == line.len() - PREFIX.len() {
        fail("LICENSE: expected \"Copyright (c) <year> <holder>\"");
    }
    (line.to_string(), holder.to_string())
}

fn pad_to_32_bits(bytes: &mut Vec<u8>) {
    bytes.resize(bytes.len().next_multiple_of(4), 0);
}

fn utf16z(text: &str) -> Vec<u8> {
    text.encode_utf16()
        .chain(std::iter::once(0))
        .flat_map(u16::to_le_bytes)
        .collect()
}

/// One entry of a .res file: RESOURCEHEADER with numeric type and name (each
/// 0xFFFF followed by the number), then the data, padded to 32 bits.
fn res_entry(out: &mut Vec<u8>, kind: u16, id: u16, memory_flags: u16, language: u16, data: &[u8]) {
    let data_size = u32::try_from(data.len()).unwrap_or_else(|_| fail("resource too large"));
    out.extend(data_size.to_le_bytes()); // DataSize
    out.extend(32u32.to_le_bytes()); // HeaderSize
    for word in [0xFFFF, kind, 0xFFFF, id] {
        out.extend(word.to_le_bytes()); // TYPE, NAME
    }
    out.extend(0u32.to_le_bytes()); // DataVersion
    out.extend(memory_flags.to_le_bytes()); // MemoryFlags
    out.extend(language.to_le_bytes()); // LanguageId
    out.extend(0u32.to_le_bytes()); // Version
    out.extend(0u32.to_le_bytes()); // Characteristics
    out.extend_from_slice(data);
    pad_to_32_bits(out);
}

/// One block of a version resource — they all share a shape: wLength (the
/// whole block, children included), wValueLength, wType (1 = text), the key
/// as a NUL-terminated UTF-16 string, padding, the value, then each child
/// after padding. The block starts on a 32-bit boundary of the resource, so
/// padding relative to the block is padding relative to the resource.
fn block(
    key: &str,
    text: bool,
    value_length: usize,
    value: &[u8],
    children: &[Vec<u8>],
) -> Vec<u8> {
    let word = |n: usize| {
        u16::try_from(n)
            .unwrap_or_else(|_| fail("version resource too large"))
            .to_le_bytes()
    };
    let mut b = vec![0, 0];
    b.extend(word(value_length));
    b.extend(word(usize::from(text)));
    b.extend(utf16z(key));
    pad_to_32_bits(&mut b);
    b.extend_from_slice(value);
    for child in children {
        pad_to_32_bits(&mut b);
        b.extend_from_slice(child);
    }
    let length = word(b.len());
    b[..2].copy_from_slice(&length);
    b
}

/// A `String` block: the value's length counts UTF-16 units, NUL included.
fn string(key: &str, value: &str) -> Vec<u8> {
    let value = utf16z(value);
    block(key, true, value.len() / 2, &value, &[])
}

/// The RT_VERSION data: VS_VERSIONINFO with its VS_FIXEDFILEINFO, one
/// StringFileInfo table and the VarFileInfo translation that names it.
fn version_info(numbers: [u16; 4], version: &str, copyright: &str, holder: &str) -> Vec<u8> {
    let [major, minor, patch, build] = numbers.map(u32::from);
    let (ms, ls) = ((major << 16) | minor, (patch << 16) | build);
    let fixed: Vec<u8> = [
        0xFEEF_04BD, // dwSignature
        0x0001_0000, // dwStrucVersion
        ms,          // dwFileVersionMS
        ls,          // dwFileVersionLS
        ms,          // dwProductVersionMS
        ls,          // dwProductVersionLS
        0x3F,        // dwFileFlagsMask
        0,           // dwFileFlags
        0x0004_0004, // dwFileOS: VOS_NT_WINDOWS32
        1,           // dwFileType: VFT_APP
        0,           // dwFileSubtype
        0,           // dwFileDateMS
        0,           // dwFileDateLS
    ]
    .into_iter()
    .flat_map(u32::to_le_bytes)
    .collect();

    // FileDescription is the name Task Manager and Settings > Apps > Startup
    // show for the app.
    let strings = [
        string("CompanyName", holder),
        string("FileDescription", PRODUCT),
        string("FileVersion", version),
        string("InternalName", EXE_NAME),
        string("LegalCopyright", copyright),
        string("OriginalFilename", EXE_NAME),
        string("ProductName", PRODUCT),
        string("ProductVersion", version),
    ];
    let table_key = format!("{LANG_EN_US:04X}{CODEPAGE_UNICODE:04X}");
    let string_file_info = block(
        "StringFileInfo",
        true,
        0,
        &[],
        &[block(&table_key, true, 0, &[], &strings)],
    );
    let translation: Vec<u8> = [LANG_EN_US, CODEPAGE_UNICODE]
        .into_iter()
        .flat_map(u16::to_le_bytes)
        .collect();
    let var_file_info = block(
        "VarFileInfo",
        true,
        0,
        &[],
        &[block(
            "Translation",
            false,
            translation.len(),
            &translation,
            &[],
        )],
    );
    block(
        "VS_VERSION_INFO",
        false,
        fixed.len(),
        &fixed,
        &[string_file_info, var_file_info],
    )
}
