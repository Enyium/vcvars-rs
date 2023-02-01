#![cfg(target_os = "windows")]
#![warn(clippy::pedantic)]

use std::{borrow::Cow, collections::HashMap, env, fs, io, path::PathBuf, process::Command};

use filenamify::filenamify;
use itertools::Itertools;
use thiserror::Error;

type EnvMap = HashMap<String, String>;

pub struct Vcvars {
    env_map: Option<EnvMap>,
}

impl Vcvars {
    //! Runs vcvars in a `cmd.exe` child process (at most once) and makes available the set of environment variables the child process inherited, mutated by vcvars. The `cmd.exe` stdout output is converted with [`std::string::String::from_utf8_lossy()`].
    //!
    //! Use [`std::env::split_paths()`] to split a variable like `INCLUDE`, which could then, e.g., be passed to [`cc::Build::includes()`].
    //!
    //! # Example
    //!
    //! ```
    //! let mut vcvars = Vcvars::new();
    //! let vcvars_include = vcvars.get_cached("INCLUDE").unwrap();
    //!
    //! cxx_build::bridge("src/demo.rs")
    //!     .file("src/demo.cc")
    //!     .includes(env::split_paths(&*vcvars_include))
    //!     .compile("demo");
    //! ```

    pub fn new() -> Self {
        #![must_use]
        #![allow(clippy::new_without_default)]

        Self { env_map: None }
    }

    pub fn get_cached(&mut self, var_name: &str) -> Result<Cow<str>, VcvarsError> {
        #![allow(clippy::missing_errors_doc)]
        //! Reads the `OUT_DIR` environment variable that Cargo sets and obtains `var_name`'s value from a cache file. If the file isn't present, runs vcvars and creates a memory cache of its variables, if not done previously, to source the value from and creates the cache file. Then returns the value.
        //!
        //! The cache files are named after the variables. The filenames are sanitized to be legal on all platforms. Should this result in two variables getting the same filename, there will be incorrect behavior. (See <https://github.com/chawyehsu/filenamify-rs/blob/main/src/lib.rs>.)
        //!
        //! # Panics
        //!
        //! Panics if the `OUT_DIR` environment variable isn't set or doesn't represent an existing directory.

        // Find Cargo output directory.
        let cargo_out_dir = PathBuf::from(
            &env::var("OUT_DIR").expect("env var `OUT_DIR` should've been set by Cargo"),
        );
        assert!(
            cargo_out_dir.is_dir(),
            "env var `OUT_DIR` should be a valid directory path"
        );

        // Create cache directory.
        let mut cache_dir = cargo_out_dir;
        cache_dir.push("vcvars-cache");
        if let Err(err) = fs::create_dir_all(&cache_dir) {
            return Err(VcvarsError::CacheFailed(
                cache_dir.to_string_lossy().into_owned(),
                err,
            ));
        }

        // Read, or prepare and write cache file.
        let mut cache_file = cache_dir;
        cache_file.push(filenamify(format!("{var_name}.txt")));

        if cache_file.exists() {
            match fs::read_to_string(&cache_file) {
                Ok(value) => Ok(Cow::Owned(value)),
                Err(err) => Err(VcvarsError::CacheFailed(
                    cache_file.to_string_lossy().into_owned(),
                    err,
                )),
            }
        } else {
            match self.ensure_env_map()?.get(&var_name.to_uppercase()) {
                Some(value) => match fs::write(&cache_file, value) {
                    Ok(()) => Ok(Cow::Borrowed(value)),
                    Err(err) => Err(VcvarsError::CacheFailed(
                        cache_file.to_string_lossy().into_owned(),
                        err,
                    )),
                },
                None => Err(VcvarsError::VarNotFound(var_name.to_owned())),
            }
        }
    }

    pub fn get(&mut self, var_name: &str) -> Result<&str, VcvarsError> {
        #![allow(clippy::missing_errors_doc)]
        //! Runs vcvars and creates a memory cache of its variables, if not done previously, and returns `var_name`'s value.

        match self.ensure_env_map()?.get(&var_name.to_uppercase()) {
            Some(value) => Ok(value),
            None => Err(VcvarsError::VarNotFound(var_name.to_owned())),
        }
    }

    fn ensure_env_map(&mut self) -> Result<&EnvMap, VcvarsError> {
        if self.env_map.is_none() {
            self.env_map = Some(Self::make_env_map()?);
        };

        Ok(self.env_map.as_ref().unwrap())
    }

    fn make_env_map() -> Result<EnvMap, VcvarsError> {
        // Read env var dependencies.
        let Ok(program_files_x86_dir) = env::var("PROGRAMFILES(X86)") else {
            return Err(VcvarsError::MissingEnvVarDependency(
                "PROGRAMFILES(X86)".to_owned(),
            ));
        };

        let Ok(win_dir) = env::var("WINDIR") else {
            return Err(VcvarsError::MissingEnvVarDependency("WINDIR".to_owned()));
        };

        // Find `vswhere`.
        let mut vswhere_path = PathBuf::from(program_files_x86_dir);
        vswhere_path.push("Microsoft Visual Studio");
        vswhere_path.push("Installer");
        vswhere_path.push("vswhere.exe");

        // Note: Microsoft says about the `vswhere` path: "This is a fixed location that will be maintained." (https://github.com/Microsoft/vswhere/wiki/Installing)

        if !vswhere_path.is_file() {
            return Err(VcvarsError::FileNotFound(
                vswhere_path.to_string_lossy().into_owned(),
            ));
        }

        // Find Visual Studio.
        let visual_studio_dir = match Command::new(&vswhere_path)
            .args(["-latest", "-property", "installationPath", "-utf8"])
            .output()
        {
            Ok(output) => {
                let dir = String::from_utf8(output.stdout)
                    .expect("`vswhere.exe` with `-utf8` switch should've returned valid UTF-8");

                dir.trim().to_owned()
            }
            Err(err) => {
                return Err(VcvarsError::CouldntRun(
                    vswhere_path.to_string_lossy().into_owned(),
                    err,
                ));
            }
        };

        // Find vcvars and determine its args.
        let mut vcvars_path = PathBuf::from(visual_studio_dir);
        vcvars_path.push("VC");
        vcvars_path.push("Auxiliary");
        vcvars_path.push("Build");
        vcvars_path.push("vcvarsall.bat");

        if !vcvars_path.is_file() {
            return Err(VcvarsError::FileNotFound(
                vcvars_path.to_string_lossy().into_owned(),
            ));
        }

        let vcvars_path = vcvars_path.to_str().unwrap(); // Built from valid UTF-8.

        let architecture = if cfg!(target_pointer_width = "64") {
            "x64"
        } else {
            "x86"
        };

        // Find `cmd.exe`.
        let mut cmd_exe_path = PathBuf::from(win_dir);
        cmd_exe_path.push("System32");
        cmd_exe_path.push("cmd.exe");

        // Run `cmd.exe` with vcvars.
        let vcvars_path = vcvars_path.replace('^', "^^").replace('&', "^&"); // Try to follow `cmd.exe`'s erratic escaping rules (tested).

        // Note: Escaping `%` by writing `%%` doesn't work, and a path containing two `%`s and the name of an existing env var in between breaks the command.

        let separator_line =
            "=".repeat(20) + "_unique_separator_by_rust_crate_that_utilizes_vcvars";

        let output = Command::new(&cmd_exe_path)
            .arg("/C")
            // Note: On the regular, interactive command line, `chcp 65001` to change the active code page to UTF-8 doesn't seem to make a difference regarding the content.
            .args([&vcvars_path, architecture, "&&"])
            .args([&format!("echo.{separator_line}"), "&&"])
            .arg("set") // Lists env vars.
            .output();

        // Note: vcvars always returns exit code 0, even if it failed (as of Dec. 2022).

        let stdout = match output {
            Ok(ref output) => String::from_utf8_lossy(&output.stdout),
            Err(err) => {
                return Err(VcvarsError::CouldntRun(
                    cmd_exe_path.to_string_lossy().into_owned(),
                    err,
                ));
            }
        };

        if stdout.starts_with("[ERROR:") {
            return Err(VcvarsError::VcvarsFailed(
                Itertools::intersperse(stdout.lines(), r"\n").collect(),
            ));
        }

        // Transform output lines to key-value pairs.
        let mut env = HashMap::new();
        let mut may_collect = false;

        // Note: The format in stdout that we get is basically identical to that of the Windows API function `GetEnvironmentStrings()`, which is only for the current process.

        for line in stdout.lines() {
            if may_collect {
                if let Some((key, value)) = line.split_once('=') {
                    env.insert(key.to_uppercase(), value.to_owned());
                }
            } else if line.starts_with(&separator_line) {
                // Note: The notoriously erratic `cmd.exe` adds a space. Hence not `==`.

                may_collect = true;
            }
        }

        Ok(env)
    }
}

#[derive(Error, Debug)]
pub enum VcvarsError {
    #[error("env var `{0}` isn't set, which is a dependency to run vcvars")]
    MissingEnvVarDependency(String),
    #[error("couldn't find file `{0}`")]
    FileNotFound(String),
    #[error("couldn't run `{0}`: {1}")]
    CouldntRun(String, io::Error),
    #[error("`vcvarsall.bat` failed: {0}")]
    VcvarsFailed(String),
    #[error("I/O operation regarding cache path `{0}` failed: {1}")]
    CacheFailed(String, io::Error),
    #[error("variable `{0}` not found in vcvars environment")]
    VarNotFound(String),
}

#[cfg(test)]
mod tests {
    use crate::Vcvars;
    use regex::Regex;
    use serial_test::serial;
    use std::{env, fs, io, path::PathBuf, time::Instant};

    fn version_number_regex() -> Regex {
        Regex::new(r"^(\d+\.)+\d+$").unwrap()
    }

    #[test]
    #[serial]
    fn get() {
        let mut vcvars = Vcvars::new();

        let start = Instant::now();
        let value = vcvars.get("VisualStudioVersion").unwrap();
        assert!(version_number_regex().is_match(value), "{value}");
        let initial_get_duration = start.elapsed();

        let start = Instant::now();
        let value = vcvars.get("INCLUDE").unwrap();
        assert!(
            Regex::new(r"(?i)^[A-Z]:\\").unwrap().is_match(value)
                && value.contains("Visual Studio")
                && value.matches(';').count() >= 4,
            "{value}"
        );
        let followup_get_duration = start.elapsed();

        assert!(
            followup_get_duration < initial_get_duration / 1000,
            "getting 2nd env var should've been much faster than getting 1st"
        );
    }

    #[test]
    #[serial]
    fn get_cached() {
        let mut cache_dir =
            PathBuf::from(env::var("OUT_DIR").expect("env var `OUT_DIR` should be set"));
        cache_dir.push("vcvars-cache");
        if let Err(err) = fs::remove_dir_all(cache_dir) {
            assert!(
                matches!(err.kind(), io::ErrorKind::NotFound),
                "should've been able to remove cache dir: {err}"
            );
        }

        let start = Instant::now();
        let mut vcvars = Vcvars::new();
        let value = vcvars.get_cached("VisualStudioVersion").unwrap();
        assert!(version_number_regex().is_match(value.as_ref()), "{value}");
        let vcvars_call_get_duration = start.elapsed();

        let start = Instant::now();
        let mut vcvars = Vcvars::new();
        let value = vcvars.get_cached("VisualStudioVersion").unwrap();
        assert!(version_number_regex().is_match(value.as_ref()), "{value}");
        let cache_get_duration = start.elapsed();

        assert!(
            cache_get_duration < vcvars_call_get_duration / 100,
            "getting env var from cache should've been much faster than getting it from vcvars call"
        );

        // Note: When writing the test, HDD vs. SSD didn't make a difference in terms of by what factor the two durations differed.
    }
}
