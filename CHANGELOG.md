# Changelog

All notable changes to **shop_doc_advanced** are documented here. Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and [Semantic Versioning](https://semver.org/).

## [1.1.9] - 2026-07-03

### Added

- **`pb__ap_wall_finish_barrel_swarf`** — Final Ap for `WALL_FINISH-BARREL_SWARF` (`ap_mill_multiaxis`): MM / % diameter / % flute from `mom_maximal_stepover_distance` keyed by `mom_cut_level_distance_source`.

## [1.1.8] - 2026-07-03

### Added

- **CSV / XLSX column** `mom_maximal_stepover_distance_source` (after `mom_maximal_stepover_distance`), with per-path unset in `pb__shop_reset_path_mom_vars`.

## [1.1.7] - 2026-07-03

### Changed

- **`WALL_FINISH-BARREL_SWARF`** in `ae_mill_multiaxis` → Final Ae = `"NO DATA"` (was `mom_maximal_stepover_distance`).

## [1.1.6] - 2026-07-03

### Fixed

- **`ZLEVEL_5AXIS` Ap (SCALLOP case):** `"NO DATA"` when `mom_common_depth_per_cut_type = 1` (was incorrectly keyed off non-zero `mom_scallop_common_depth_per_cut`).

## [1.1.5] - 2026-07-03

### Added

- **`pb__ap_zlevel_5axis`** — Final Ap for `ZLEVEL_5AXIS` (`ap_mill_multiaxis`): constant depth (MM / % diameter / % flute) when `mom_scallop_common_depth_per_cut = 0`; `"NO DATA"` when scallop is non-zero.

## [1.1.4] - 2026-07-03

### Changed

- **`CONTOUR_PROFILE`** in `ap_mill_multiaxis` uses `solid_profile_3d` → `pb__ap_solid_profile_3d` (increment vs passes from `mom_multi_depth_cut_type`).
- **PASSES** case in `pb__ap_solid_profile_3d` now requires `mom_stock_part_offset` non-zero (shared with `SOLID_PROFILE_3D` / `PROFILE_3D`).

## [1.1.3] - 2026-07-03

### Changed

- **`VARIABLE_AXIS_GUIDING_CURVES`** in `ap_mill_multiaxis` now uses token `var_axis_gc` (same as `ae_mill_multiaxis`) and `pb__ae_variable_axis_guiding_curves` for Final Ap.

## [1.1.2] - 2026-07-03

### Added

- **`pb__ap_multi_axis_roughing`** — Final Ap for `MULTI_AXIS_ROUGHING` (`ap_mill_multiaxis`): AUTO MM when `mom_cut_level_distance_source` is unset; % diameter (source 4) or % flute (source 7) otherwise.

## [1.1.1] - 2026-07-03

### Added

- **`pb__ap_3d_adaptive_roughing`** — Final Ap for `3D_ADAPTIVE_ROUGHING` (`ap_mill_contour`): constant MM when `mom_cut_level_distance_source` is unset; % diameter (source 4) or % flute (source 7) otherwise.

## [1.1.0] - 2026-07-03

### Added

- **CSV / XLSX columns** after existing cut-level fields:
  - `mom_common_depth_per_cut_type`
  - `mom_scallop_common_depth_per_cut`
  - `mom_multi_depth_cut_type`
  - `mom_multi_depth_cut_passes_number`
  - `mom_stock_part_offset`
- **Final Ap procs** (table-driven via `ap_mill_planar` / `ap_mill_contour`):
  - `pb__ap_face_mill_midpass` — `FACE_MILL_MIDPASS`, `FACE_MILL_SPIRAL`, `FACE_MILL_ZIGZAG`, `2D_WALL_MILL` (constant depth MM / % diameter / % flute; levels → *N Passes*)
  - `pb__ap_common_depth_per_cut` — `CAVITY_MILL`, `ADAPTIVE_MILLING`, `REST_MILLING`, `ZLEVEL_PROFILE_STEEP`, `ZLEVEL_UNDERCUT`
  - `pb__ap_fixed_axis_guiding_curves` — `FIXED_AXIS_GUIDING_CURVES`
  - `pb__ap_stepover_flow_area` — `AREA_MILL`, `CURVE_DRIVE` (stepover type 1–4)
  - `pb__ap_flow_mill_multiple` — `FLOW_MILL_MULTIPLE`, `FLOWCUT_MULTIPLE`, `FLOW_MILL_REF_TOOL`, `FLOWCUT_REF_TOOL` (stepover distance when ≤ flute length, else `NO DATA`)
  - `pb__ap_solid_profile_3d` — `SOLID_PROFILE_3D`, `PROFILE_3D` (increment vs passes)
- **Final Ae procs** (table-driven via `ae_mill_contour`):
  - `pb__ae_stepover_flow_area` — `AREA_MILL`, `CURVE_DRIVE` (and shared stepover-type logic)
  - `pb__ae_flow_mill_multiple` — `FLOW_MILL_MULTIPLE`, `FLOWCUT_MULTIPLE`, `FLOW_MILL_REF_TOOL`, `FLOWCUT_REF_TOOL`
  - `pb__ae_fixed_axis_guiding_curves` — `FIXED_AXIS_GUIDING_CURVES`

### Changed

- **Final Ap (Axial DOC)** and **Final Ae (Radial DOC)** for many `mill_planar` and `mill_contour` subtypes now use dedicated dispatch tables (`ap_*` / `ae_*` arrays) instead of a single `mom_depth_per_cut` default.
- **`mom_global_cut_depth_source`** added to per-path reset list (`pb__shop_reset_path_mom_vars`) so common-depth Ap logic does not leak between operations.
- All new `mom_*` CSV columns are unset at end of each path per shop-doc reset policy.

### Fixed

- Stale `mom_*` values no longer carry into the next CSV row for newly added columns and depth-per-cut parameters.

---

## [1.0.0] - 2026-03-01

### Added

- Initial shop-doc post: CSV export, ClosedXML XLSX conversion, table styling, validation lists, and base `mom_*` parameter columns.

[1.1.9]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.8...v1.1.9
[1.1.8]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.7...v1.1.8
[1.1.7]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.6...v1.1.7
[1.1.6]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.5...v1.1.6
[1.1.5]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.4...v1.1.5
[1.1.4]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.3...v1.1.4
[1.1.3]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.2...v1.1.3
[1.1.2]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.1...v1.1.2
[1.1.1]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/HakimHisham1991/shop_doc_advanced/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/HakimHisham1991/shop_doc_advanced/releases/tag/v1.0.0
