# `PB_CMD_shop_end_path` — Final Ae & Ap Logic by Template Subtype

> Source of truth: `shop_doc_advanced.tcl` — dispatch arrays `ae_mill_*` / `ap_mill_*` and helper procs `pb__ae_*` / `pb__ap_*`.
>
> If a subtype is not listed below, or dispatch does not match, **Final Ae** and **Final Ap** remain **`N/A`** (initial default).

---

## Reference — shared variables

| Variable | Source | Meaning |
|---|---|---|
| `sot` | `mom_stepover_type` | Current stepover type at end of path |
| `sot_early` | 1st entry in `path_stepover_type_list` | First traced `mom_stepover_type` write |
| `sds` | `mom_stepover_distance_source` | Stepover distance unit/source code |
| `sds_defined` | `[info exists mom_stepover_distance_source]` | Whether source variable exists |
| `stepover_var_1` | `mom_stepover_variable_max_min(1)` | Variable stepover max (or `N/A`) |
| `step_points_2` | `mom_step_points(2)` | Number of passes (or `N/A`) |
| `path_stepover_2` | 2nd entry in `path_stepover_distance_list` | Traced stepover distance (or `N/A`) |
| `stepover_percent_2` | 2nd entry in `path_stepover_percent_list` | Traced stepover percent (or `N/A`) |
| `max_stepover_var_tool_dep` | max of `mom_stepover_variable_tool_dependent_values(0..99)` | Used by THREAD_MILLING sot=3 |

### `mom_stepover_type` codes

| Code | Meaning |
|---|---|
| 1 | Constant |
| 2 | Scallop |
| 3 | Variable Average *(region 1/2/3)* or Multiple *(region 4/5/7)* — see handler |
| 4 | Percent of Tool Flat |
| 5 | Passes |
| 8 | Variable *(THREAD_MILLING only)* |
| 9 | Exact |

### `mom_region_cut_method` codes

| Code | Meaning |
|---|---|
| 1 | Zig-Zag |
| 2 | Zig |
| 3 | Zig-Zag with Contour |
| 4 | Follow Periphery |
| 5 | Profile |
| 7 | Follow Part |

### Common formulas

| Pattern | Formula |
|---|---|
| Constant MM | `mom_stepover_distance` (rounded) |
| Constant % Ø | `mom_stepover_distance / 100 × mom_tool_diameter` |
| % tool flat | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| Tool-dep raw (src=0) | value as-is |
| Tool-dep % Ø (src=4) | `value / 100 × mom_tool_diameter` |
| Passes label | `{n} Passes` from `step_points_2` (integer-rounded when numeric) |

### Subtype inventory (52 total)

| `mom_template_type` | Count | Subtypes |
|---|---|---|
| `mill_planar` | 15 | FACE_MILL_MIDPASS, FACE_MILL_SPIRAL, FACE_MILL_ZIGZAG, 2D_WALL_MILL, FLOOR_WALL, POCKETING, WALL_PROFILING, WALL_FLOOR_PROFILING, PLANAR_PROFILING, PLANAR_MILL, GROOVE_MILLING, PLANAR_DEBURRING, MILL_CONTROL, FLOOR_FACING, FACE_MILLING_MANUAL |
| `mill_contour` | 22 | CAVITY_MILL, ADAPTIVE_MILLING, 3D_ADAPTIVE_ROUGHING, PLUNGE_MILLING, QUICK_ROUGHING, REST_MILLING, ZLEVEL_PROFILE_STEEP, ZLEVEL_UNDERCUT, FIXED_AXIS_GUIDING_CURVES, AREA_MILL, FLOW_MILL_SINGLE, FLOWCUT_SINGLE, FLOW_MILL_MULTIPLE, FLOWCUT_MULTIPLE, FLOW_MILL_REF_TOOL, FLOWCUT_REF_TOOL, CURVE_DRIVE, SOLID_PROFILE_3D, PROFILE_3D, STREAMLINE, CONTOUR_SURFACE_AREA, 3_AXIS_DEBURRING |
| `mill_multi-axis` | 8 | MULTI_AXIS_ROUGHING, VARIABLE_AXIS_GUIDING_CURVES, CONTOUR_PROFILE, VARIABLE_STREAMLINE, VARIABLE_CONTOUR, WALL_FINISH-BARREL_SWARF, ZLEVEL_5AXIS, 5_AXIS_DEBURRING |
| `hole_making` | 7 | SPOT_DRILLING, DRILLING, BORING_REAMING, TAPPING, DEEP_HOLE_DRILLING, HOLE_MILLING, THREAD_MILLING |

> **Case sensitivity:** `mill_planar` Ae/Ap keys are matched **uppercase**. `mill_contour`, `mill_multi-axis`, and `hole_making` keys are matched **as-is** (typically uppercase in NX).

---

## Reusable Ae handlers (`mill_planar`)

### `pb__ae_std_sot_123459` — FLOOR_FACING

| sot | Condition | Final Ae |
|---|---|---|
| 1 | `sds_defined` && `sds == 4` && Ø ≠ 0 | `mom_stepover_distance / 100 × Ø` |
| 1 | `!sds_defined` | `mom_stepover_distance` |
| 2 | — | `NO DATA` |
| 3 | region ∈ {4, 5, 7} | max of `tool_dep_values(0..5)` (src 0 = raw; src 4 = `value/100 × Ø`) |
| 3 | region ∈ {1, 2, 3} | `stepover_var_1` |
| 4 | Ø and corner radius exist, flat ≠ 0 | `% tool flat` formula |
| 5 | — | `{step_points_2} Passes` |
| 9 | same as sot 1 | Constant / % Ø per `sds` |

### `pb__ae_std_sot_14` — FACE_MILL_SPIRAL

| sot | Condition | Final Ae |
|---|---|---|
| 1 | `sds == 4` && Ø ≠ 0 | `mom_stepover_distance / 100 × Ø` |
| 1 | `!sds_defined` | `mom_stepover_distance` |
| 4 | flat ≠ 0 | `% tool flat` formula |

### `pb__ae_std_sot_145` — FACE_MILL_ZIGZAG

| sot | Condition | Final Ae |
|---|---|---|
| 1 | `sds == 4` && Ø ≠ 0 | `mom_stepover_distance / 100 × Ø` |
| 1 | `!sds_defined` | `mom_stepover_distance` |
| 4 | flat ≠ 0 | `% tool flat` formula |
| 5 | — | `NO DATA` |

### `pb__ae_std_sot_1234` — FLOOR_WALL, POCKETING, WALL_PROFILING, WALL_FLOOR_PROFILING, PLANAR_MILL

| sot | Condition | Final Ae |
|---|---|---|
| 1 | `sds == 4` && Ø ≠ 0 | `mom_stepover_distance / 100 × Ø` |
| 1 | `!sds_defined` | `mom_stepover_distance` |
| 2 | — | `NO DATA` |
| 3 | — | max of `tool_dep_values(0..5)` **(no region split)** |
| 4 | flat ≠ 0 | `% tool flat` formula |

### `pb__ae_std_sot_1234_m` — FACE_MILLING_MANUAL

| sot | Condition | Final Ae |
|---|---|---|
| 1 | `sds == 4` && Ø ≠ 0 | `mom_stepover_distance / 100 × Ø` |
| 1 | `!sds_defined` | `mom_stepover_distance` |
| 2 | — | `NO DATA` |
| 3 | — | `stepover_var_1` |
| 4 | flat ≠ 0 | `% tool flat` formula |

### `pb__ae_std_sot_15` — GROOVE_MILLING

| sot | Condition | Final Ae |
|---|---|---|
| 1 | `sds == 4` && Ø ≠ 0 | `mom_stepover_distance / 100 × Ø` |
| 1 | `!sds_defined` | `mom_stepover_distance` |
| 5 | — | `{step_points_2} Passes` |

---

## Reusable Ap handlers

### `pb__ap_face_mill_midpass` — FACE_MILL_MIDPASS, FACE_MILL_SPIRAL, FACE_MILL_ZIGZAG, 2D_WALL_MILL

| `mom_cut_levels_mode` | Condition | Final Ap |
|---|---|---|
| 1 | `mom_number_of_cut_levels` exists | `{n} Passes` |
| 0 | `mom_cut_level_distance` exists, no source | `mom_cut_level_distance` |
| 0 | source = 4, Ø ≠ 0 | `mom_cut_level_distance / 100 × Ø` |
| 0 | source = 7, flute length ≠ 0 | `mom_cut_level_distance / 100 × mom_tool_flute_length` |
| other | — | `N/A` |

### `pb__ap_planar_mill` — PLANAR_PROFILING, PLANAR_MILL

| `mom_depth_of_cut_type` | Final Ap |
|---|---|
| 0, 4 | `mom_cut_level_max_depth` |
| 1, 2, 3 | `NO DATA` |
| missing | `N/A` |

### `pb__ap_groove_milling` — GROOVE_MILLING

| `mom_axial_stepover_type` | Condition | Final Ap |
|---|---|---|
| 2 | `mom_axial_stepover_passes` exists | `{n} Passes` |
| 3 | source = 4 | `mom_axial_stepover_percent / 100 × mom_tool_flute_length` |
| 6 | no source | `mom_axial_stepover_distance` |
| 6 | source defined, Ø ≠ 0 | `mom_axial_stepover_distance / 100 × Ø` |
| other | — | `N/A` |

### `pb__ap_common_depth_per_cut` — CAVITY_MILL, ADAPTIVE_MILLING, REST_MILLING, ZLEVEL_PROFILE_STEEP, ZLEVEL_UNDERCUT

| `mom_common_depth_per_cut_type` | Condition | Final Ap |
|---|---|---|
| 0 | no `mom_global_cut_depth_source` | `mom_global_cut_depth` |
| 0 | source = 4, Ø ≠ 0 | `mom_global_cut_depth / 100 × Ø` |
| 0 | source = 7, flute ≠ 0 | `mom_global_cut_depth / 100 × mom_tool_flute_length` |
| 1 | — | `NO DATA` |
| missing | — | `N/A` |

### `pb__ap_3d_adaptive_roughing` — 3D_ADAPTIVE_ROUGHING, QUICK_ROUGHING

| Condition | Final Ap |
|---|---|
| no `mom_cut_level_distance_source` | `mom_cut_level_distance` |
| source = 4, Ø ≠ 0 | `mom_cut_level_distance / 100 × Ø` |
| source = 7, flute ≠ 0 | `mom_cut_level_distance / 100 × mom_tool_flute_length` |
| other | `N/A` |

### `pb__ap_fixed_axis_guiding_curves` / `pb__ae_fixed_axis_guiding_curves` — FIXED_AXIS_GUIDING_CURVES (Ae & Ap)

| sot | Condition | Result |
|---|---|---|
| 2, 5 | — | `NO DATA` |
| 1 | `mom_stepover_distance ≤ mom_tool_flute_length` | `mom_stepover_distance` |
| 1 | `mom_stepover_distance > mom_tool_flute_length` | `NO DATA` |
| other | — | `N/A` |

### `pb__ap_stepover_flow_area` / `pb__ae_stepover_flow_area` — AREA_MILL, CURVE_DRIVE (Ae & Ap)

| sot | Condition | Result |
|---|---|---|
| 2, 3 | — | `NO DATA` |
| 1 | `mom_stepover_distance ≤ mom_tool_flute_length` | `mom_stepover_distance` |
| 1 | `mom_stepover_distance > mom_tool_flute_length` | `NO DATA` |
| 4 | flat defined | `% tool flat` formula |
| other | — | `N/A` |

### `pb__ap_flow_mill_multiple` / `pb__ae_flow_mill_multiple` — FLOW_MILL_MULTIPLE, FLOWCUT_MULTIPLE, FLOW_MILL_REF_TOOL, FLOWCUT_REF_TOOL

| Condition | Final Ae / Ap |
|---|---|
| always | `mom_stepover_distance` |

### `pb__ap_solid_profile_3d` — SOLID_PROFILE_3D, PROFILE_3D, CONTOUR_PROFILE (Ap)

| `mom_multi_depth_cut_type` | Condition | Final Ap |
|---|---|---|
| 0 | — | `mom_multi_depth_cut_increment` |
| 1 | passes ≠ 0 and `mom_stock_part_offset` ≠ 0 | `mom_stock_part_offset / mom_multi_depth_cut_passes_number` |
| other | — | `N/A` |

### `pb__ae_ap_max_scallop_limits {strict}`

Returns max of `mom_stepover_scallop`, `mom_horizonal_limit`, `mom_vertical_limit` among defined values.
- `strict = 0` — non-strict numeric check (CONTOUR_SURFACE_AREA Ap, etc.)
- `strict = 1` — strict double check (VARIABLE_CONTOUR Ap)

---

## 1. `mill_planar`

### 1a. Final Ae — dispatch summary

| Subtype | Handler | Notes |
|---|---|---|
| FACE_MILL_MIDPASS | `tool_dia` | `mom_tool_diameter` |
| 2D_WALL_MILL | `tool_dia` | `mom_tool_diameter` |
| PLANAR_PROFILING | `planar_profiling` | `NO DATA` if tool type = "User Defined Mill Tool"; else `tool_dia` |
| PLANAR_DEBURRING | `no_data` | `NO DATA` |
| MILL_CONTROL | `no_data` | `NO DATA` |
| FACE_MILL_SPIRAL | `std_sot_14` | see handler table |
| FACE_MILL_ZIGZAG | `std_sot_145` | see handler table |
| FLOOR_WALL | `std_sot_1234` | see handler table |
| POCKETING | `std_sot_1234` | same as FLOOR_WALL |
| WALL_PROFILING | `std_sot_1234` | same as FLOOR_WALL |
| WALL_FLOOR_PROFILING | `std_sot_1234` | same as FLOOR_WALL |
| PLANAR_MILL | `std_sot_1234` | same as FLOOR_WALL |
| GROOVE_MILLING | `std_sot_15` | see handler table |
| FLOOR_FACING | `std_sot_123459` | see handler table |
| FACE_MILLING_MANUAL | `std_sot_1234_m` | see handler table |

### 1b. Final Ap — dispatch summary

| Subtype | Handler | Notes |
|---|---|---|
| FACE_MILL_MIDPASS | `face_mill_midpass` | see handler table |
| FACE_MILL_SPIRAL | `face_mill_midpass` | same |
| FACE_MILL_ZIGZAG | `face_mill_midpass` | same |
| 2D_WALL_MILL | `face_mill_midpass` | same |
| FLOOR_WALL | `depth_per_cut` | `mom_depth_per_cut` |
| POCKETING | `depth_per_cut` | `mom_depth_per_cut` |
| WALL_PROFILING | `depth_per_cut` | `mom_depth_per_cut` |
| WALL_FLOOR_PROFILING | `depth_per_cut` | `mom_depth_per_cut` |
| PLANAR_PROFILING | `planar_mill` | see handler table |
| PLANAR_MILL | `planar_mill` | see handler table |
| GROOVE_MILLING | `groove_milling` | see handler table |
| PLANAR_DEBURRING | `no_data` | `NO DATA` |
| MILL_CONTROL | `no_data` | `NO DATA` |
| FLOOR_FACING | `depth_per_cut` | `mom_depth_per_cut` |
| FACE_MILLING_MANUAL | `depth_per_cut` | `mom_depth_per_cut` |

---

## 2. `mill_contour`

### 2a. Final Ae — per-subtype logic

| Subtype | Logic |
|---|---|
| **CAVITY_MILL** | `pb__ae_cavity_mill`: sot 1 — MM if no source; if source defined && Ø ≠ 0 → `/100 × Ø`. sot 2 → `NO DATA`. sot 3 — region {4,5,7} → max tool-dep; region {1,2,3} → `stepover_var_1`. sot 4 → `% tool flat`. |
| **REST_MILLING** | same as CAVITY_MILL |
| **ADAPTIVE_MILLING** | `pb__ae_adaptive_milling`: sot 1 only — MM if no source; if `sds == 4` && Ø ≠ 0 → `/100 × Ø`. |
| **PLUNGE_MILLING** | `pb__ae_plunge_mill`: sot 1 — same as CAVITY_MILL sot 1; sot 4 → `% tool flat`. |
| **3D_ADAPTIVE_ROUGHING** | `pb__ae_3d_adaptive_roughing`: sot 1 — `path_stepover_2` (MM or `/100 × Ø` if `sds == 4`). **Also** if `sot_early == 4` && `sds == 4` → `path_stepover_2 / 100 × flat`. |
| **QUICK_ROUGHING** | `pb__ae_quick_roughing`: if `sot_early == 4` → `stepover_percent_2 / 100 × flat`. Else if `sot_early == 1` && sot == 1 → `path_stepover_2` logic (same as 3D adaptive sot 1). |
| **ZLEVEL_PROFILE_STEEP** | `mom_tool_diameter` |
| **ZLEVEL_UNDERCUT** | `mom_tool_diameter` |
| **FLOW_MILL_SINGLE** | `mom_tool_diameter` |
| **FLOWCUT_SINGLE** | `mom_tool_diameter` |
| **FIXED_AXIS_GUIDING_CURVES** | `pb__ae_fixed_axis_guiding_curves` — see shared handler |
| **AREA_MILL** | `pb__ae_stepover_flow_area` — see shared handler |
| **CURVE_DRIVE** | `pb__ae_stepover_flow_area` — see shared handler |
| **FLOW_MILL_MULTIPLE** | `mom_stepover_distance` |
| **FLOWCUT_MULTIPLE** | `mom_stepover_distance` |
| **FLOW_MILL_REF_TOOL** | `mom_stepover_distance` |
| **FLOWCUT_REF_TOOL** | `mom_stepover_distance` |
| **SOLID_PROFILE_3D** | `mom_wall_increment` |
| **PROFILE_3D** | `mom_wall_increment` |
| **STREAMLINE** | `NO DATA` |
| **CONTOUR_SURFACE_AREA** | `pb__ae_contour_surface_area`: sot 2 → max(`mom_vertical_limit`, `mom_horizonal_limit`); sot 5 → `{step_points_2} Passes`; else `N/A`. |
| **3_AXIS_DEBURRING** | `mom_deburring_edge_depth` |

### 2b. Final Ap — dispatch summary

| Subtype | Handler | Notes |
|---|---|---|
| CAVITY_MILL | `common_depth_per_cut` | see handler table |
| ADAPTIVE_MILLING | `common_depth_per_cut` | same |
| REST_MILLING | `common_depth_per_cut` | same |
| ZLEVEL_PROFILE_STEEP | `common_depth_per_cut` | same |
| ZLEVEL_UNDERCUT | `common_depth_per_cut` | same |
| 3D_ADAPTIVE_ROUGHING | `adaptive_3d_ap` | see handler table |
| QUICK_ROUGHING | `adaptive_3d_ap` | same |
| PLUNGE_MILLING | `flute_length` | `mom_tool_flute_length` |
| FIXED_AXIS_GUIDING_CURVES | `fixed_axis_guiding` | see shared handler |
| AREA_MILL | `stepover_flow_area` | see shared handler |
| CURVE_DRIVE | `stepover_flow_area` | same |
| FLOW_MILL_MULTIPLE | `flow_mill_multiple` | `mom_stepover_distance` |
| FLOWCUT_MULTIPLE | `flow_mill_multiple` | same |
| FLOW_MILL_REF_TOOL | `flow_mill_multiple` | same |
| FLOWCUT_REF_TOOL | `flow_mill_multiple` | same |
| SOLID_PROFILE_3D | `solid_profile_3d` | see handler table |
| PROFILE_3D | `solid_profile_3d` | same |
| FLOW_MILL_SINGLE | `no_data` | `NO DATA` |
| FLOWCUT_SINGLE | `no_data` | `NO DATA` |
| STREAMLINE | `no_data` | `NO DATA` |
| CONTOUR_SURFACE_AREA | `max_scallop` | max(scallop, horiz, vert) — non-strict |
| 3_AXIS_DEBURRING | `deburring` | `mom_deburring_edge_depth` |

---

## 3. `mill_multi-axis`

### 3a. Final Ae — per-subtype logic

| Subtype | Logic |
|---|---|
| **MULTI_AXIS_ROUGHING** | `pb__ae_multi_axis_roughing`: sot 1 — MM if no source; if source defined && Ø ≠ 0 → `/100 × Ø`. sot 2 → `NO DATA`. sot 4 → `% tool flat`. |
| **VARIABLE_AXIS_GUIDING_CURVES** | `pb__ae_variable_axis_guiding_curves`: sot 1 — if distance numeric: min(distance, Ø) when distance > Ø else distance. sot 2, 5 → `NO DATA`. |
| **CONTOUR_PROFILE** | `pb__ae_contour_profile`: `mom_wall_step_method == 0` → `mom_wall_increment`. method == 1 && passes ≠ 0 && stock offset set → `mom_wall_stock_offset / mom_wall_number_passes`. |
| **VARIABLE_STREAMLINE** | `NO DATA` |
| **VARIABLE_CONTOUR** | `pb__ae_variable_contour`: sot 2 → max(vert, horiz limits); sot 5 → `{int(step_points_2)+1} Passes`. |
| **WALL_FINISH-BARREL_SWARF** | `NO DATA` |
| **ZLEVEL_5AXIS** | `mom_tool_diameter` |
| **5_AXIS_DEBURRING** | `mom_deburring_edge_depth` |

### 3b. Final Ap — per-subtype logic

| Subtype | Handler | Logic |
|---|---|---|
| MULTI_AXIS_ROUGHING | `multi_axis_rgh_ap` | `pb__ap_multi_axis_roughing`: no source → `mom_cut_level_distance`; source 4 → `/100 × Ø`; source 7 → `/100 × flute` |
| VARIABLE_AXIS_GUIDING_CURVES | `var_axis_gc` | **same proc as Ae** — `pb__ae_variable_axis_guiding_curves` |
| CONTOUR_PROFILE | `solid_profile_3d` | see `pb__ap_solid_profile_3d` |
| VARIABLE_STREAMLINE | `no_data` | `NO DATA` |
| VARIABLE_CONTOUR | `max_scallop_s` | max(scallop, horiz, vert) — **strict** double check |
| WALL_FINISH-BARREL_SWARF | `wall_finish_swarf` | `pb__ap_wall_finish_barrel_swarf`: no source → `mom_maximal_stepover_distance`; source 4 → `/100 × Ø`; source 7 → `/100 × flute` |
| ZLEVEL_5AXIS | `zlevel_5axis` | `pb__ap_zlevel_5axis`: if `mom_common_depth_per_cut_type == 1` → `NO DATA`. If `mom_scallop_common_depth_per_cut != 0` → `N/A`. Else global cut depth (MM / % Ø / % flute per source). |
| 5_AXIS_DEBURRING | `deburring` | `mom_deburring_edge_depth` |

---

## 4. `hole_making`

### 4a. Final Ae

| Subtype | Logic |
|---|---|
| SPOT_DRILLING | `mom_tool_diameter` |
| DRILLING | `mom_tool_diameter` |
| BORING_REAMING | `mom_tool_diameter` |
| TAPPING | `mom_tool_diameter` |
| DEEP_HOLE_DRILLING | `mom_tool_diameter` |
| HOLE_MILLING | `mom_tool_diameter` |
| **THREAD_MILLING** | sot 1: `mom_stepover_distance` if ≠ 0, else `mom_tool_pitch`. sot 3: `max_stepover_var_tool_dep`. sot 8: `stepover_var_1`. |

### 4b. Final Ap

| Subtype | Logic |
|---|---|
| SPOT_DRILLING | `mom_tool_flute_length` |
| BORING_REAMING | `mom_tool_flute_length` |
| DEEP_HOLE_DRILLING | `mom_tool_flute_length` |
| THREAD_MILLING | `mom_tool_flute_length` |
| **DRILLING** | Base: `mom_tool_flute_length`. Override: `mom_cycle_step1` if ≠ 0. Override: if `mom_depth_increment_distance_source == 4` → `mom_depth_increment_distance / 100 × Ø`. |
| **TAPPING** | same overrides as DRILLING |
| **HOLE_MILLING** | `mom_vertical_pitch_type == 0`: no source or source 0 → `mom_vertical_pitch_value`; source 4 → `/100 × Ø`. `pitch_type ≠ 0` → `mom_tool_flute_length`. |

---

## Handler token quick-reference

### Ae tokens by template type

| Token | Used in |
|---|---|
| `tool_dia` | mill_planar, mill_contour, hole_making |
| `no_data` | mill_planar, mill_contour, mill_multi-axis |
| `std_sot_*` | mill_planar only |
| `planar_profiling` | mill_planar only |
| `cavity_mill`, `adaptive_mill`, `plunge_mill`, `adaptive_3d`, `quick_rough` | mill_contour |
| `stepover_dist_guiding`, `stepover_flow_area`, `flow_mill_multiple` | mill_contour |
| `wall_incr`, `deburring`, `contour_surface_area` | mill_contour |
| `multi_axis_rgh`, `var_axis_gc`, `contour_profile`, `variable_contour` | mill_multi-axis |
| `thread_mill` | hole_making |

### Ap tokens by template type

| Token | Used in |
|---|---|
| `face_mill_midpass`, `planar_mill`, `groove_milling`, `depth_per_cut` | mill_planar |
| `common_depth_per_cut`, `adaptive_3d_ap`, `fixed_axis_guiding`, `stepover_flow_area`, `flow_mill_multiple`, `solid_profile_3d`, `max_scallop`, `flute_length` | mill_contour |
| `multi_axis_rgh_ap`, `var_axis_gc`, `zlevel_5axis`, `wall_finish_swarf`, `max_scallop_s` | mill_multi-axis |
| `flute_only`, `drill_with_peck`, `vertical_pitch` | hole_making |
