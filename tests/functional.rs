use std::path::{Path, PathBuf};

use rexcell::{common, excell, range_ops};

#[test]
fn harness_self_check() {
    let p = reference_file("Test_Twins.xlsx");
    compare_workbooks(&p, &p).expect("self-compare must be equal");
}
