# `PB_CMD_shop_end_path` — Ae & Ap Logic by Template Subtype

> **Reference key for shared variables**
>
> | Variable | Meaning |
> |---|---|
> | `sot` | `mom_stepover_type` |
> | `sds` | `mom_stepover_distance_source` |
> | `sds_defined` | Whether `mom_stepover_distance_source` exists |
> | `stepover_var_1` | `mom_stepover_variable_max_min(1)` — max of variable stepover range |
> | `step_points_2` | `mom_step_points(2)` — number of passes |
> | `path_stepover_2` | 2nd recorded value from `path_stepover_distance_list` (traced during path) |
> | `stepover_percent_2` | 2nd recorded value from `path_stepover_percent_list` |
> | `sot_early` | 1st recorded `mom_stepover_type` write (traced before end of path) |
> | `max_stepover_var_tool_dep` | max of all `mom_stepover_variable_tool_dependent_values(0..99)` |
>
> **`mom_stepover_type` codes**
>
> | Code | Meaning |
> |---|---|
> | 1 | Constant |
> | 2 | Scallop |
> | 3 | Variable Average *(if region = 1/2/3)* or Multiple *(if region = 4/5/7)* |
> | 4 | Percent of Tool Flat |
> | 5 | Passes |
> | 8 | Variable (used by THREAD_MILLING) |
> | 9 | Exact |
>
> **`mom_region_cut_method` codes**
>
> | Code | Meaning |
> |---|---|
> | 1 | Zig-Zag |
> | 2 | Zig |
> | 3 | Zig-Zag with Contour |
> | 4 | Follow Periphery |
> | 5 | Profile |
> | 7 | Follow Part |

---

## 1. `mill_planar`

### 1a. Final Ae — `mill_planar`

| Operation Subtype | Condition | Final Ae |
|---|---|---|
| **FACE_MILL_MIDPASS** | *(always)* | `mom_tool_diameter` |
| **2D_WALL_MILL** | *(always)* | `mom_tool_diameter` |
| **PLANAR_PROFILING** | Tool type = "User Defined Mill Tool" | `NO DATA` |
| **PLANAR_PROFILING** | Tool type ≠ "User Defined Mill Tool" | `mom_tool_diameter` |
| **PLANAR_DEBURRING** | *(always)* | `NO DATA` |
| **MILL_CONTROL** | *(always)* | `NO DATA` |
| **FACE_MILL_SPIRAL** | sot = 1, sds = 4 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **FACE_MILL_SPIRAL** | sot = 1, sds not defined | `mom_stepover_distance` |
| **FACE_MILL_SPIRAL** | sot = 4 | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **FACE_MILL_ZIGZAG** | sot = 1, sds = 4 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **FACE_MILL_ZIGZAG** | sot = 1, sds not defined | `mom_stepover_distance` |
| **FACE_MILL_ZIGZAG** | sot = 4 | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **FACE_MILL_ZIGZAG** | sot = 5 | `NO DATA` |
| **GROOVE_MILLING** | sot = 1, sds = 4 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **GROOVE_MILLING** | sot = 1, sds not defined | `mom_stepover_distance` |
| **GROOVE_MILLING** | sot = 5 | `step_points_2` Passes |
| **FLOOR_WALL** | sot = 1, sds = 4 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **FLOOR_WALL** | sot = 1, sds not defined | `mom_stepover_distance` |
| **FLOOR_WALL** | sot = 2 | `NO DATA` |
| **FLOOR_WALL** | sot = 3, region ∈ {4,5,7} | max of all `tool_dep_values(i)` — if src=0: raw value; if src=4: `value / 100 × tool_diameter` |
| **FLOOR_WALL** | sot = 3, region ∈ {1,2,3} | `stepover_var_1` (variable stepover max) |
| **FLOOR_WALL** | sot = 4 | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **POCKETING** | *(same logic as FLOOR_WALL — sot 1/2/3/4)* | *(see FLOOR_WALL rows above)* |
| **WALL_PROFILING** | *(same logic as FLOOR_WALL — sot 1/2/3/4)* | *(see FLOOR_WALL rows above)* |
| **WALL_FLOOR_PROFILING** | *(same logic as FLOOR_WALL — sot 1/2/3/4)* | *(see FLOOR_WALL rows above)* |
| **PLANAR_MILL** | *(same logic as FLOOR_WALL — sot 1/2/3/4)* | *(see FLOOR_WALL rows above)* |
| **FACE_MILLING_MANUAL** | sot = 1, sds = 4 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **FACE_MILLING_MANUAL** | sot = 1, sds not defined | `mom_stepover_distance` |
| **FACE_MILLING_MANUAL** | sot = 2 | `NO DATA` |
| **FACE_MILLING_MANUAL** | sot = 3, region ∈ {1,2,3} | `stepover_var_1` (variable stepover max) |
| **FACE_MILLING_MANUAL** | sot = 4 | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **FLOOR_FACING** | sot = 1, sds = 4 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **FLOOR_FACING** | sot = 1, sds not defined | `mom_stepover_distance` |
| **FLOOR_FACING** | sot = 2 | `NO DATA` |
| **FLOOR_FACING** | sot = 3, region ∈ {4,5,7} | max of all `tool_dep_values(i)` (src=0: raw; src=4: `value/100 × Ø`) |
| **FLOOR_FACING** | sot = 3, region ∈ {1,2,3} | `stepover_var_1` |
| **FLOOR_FACING** | sot = 4 | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **FLOOR_FACING** | sot = 5 | `step_points_2` Passes |
| **FLOOR_FACING** | sot = 9 (Exact), sds = 4 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **FLOOR_FACING** | sot = 9 (Exact), sds not defined | `mom_stepover_distance` |

### 1b. Final Ap — `mill_planar`

| Operation Subtype | Condition | Final Ap |
|---|---|---|
| **FACE_MILL_MIDPASS** | *(always)* | `mom_cut_level_distance` |
| **FACE_MILL_SPIRAL** | *(always)* | `mom_cut_level_distance` |
| **FACE_MILL_ZIGZAG** | *(always)* | `mom_cut_level_distance` |
| **2D_WALL_MILL** | *(always)* | `mom_cut_level_distance` |
| **FLOOR_WALL** | *(always)* | `mom_depth_per_cut` |
| **POCKETING** | *(always)* | `mom_depth_per_cut` |
| **WALL_PROFILING** | *(always)* | `mom_depth_per_cut` |
| **WALL_FLOOR_PROFILING** | *(always)* | `mom_depth_per_cut` |
| **PLANAR_PROFILING** | *(always)* | `mom_z_depth_offset` |
| **PLANAR_MILL** | *(always)* | `mom_cut_level_max_depth` |
| **GROOVE_MILLING** | *(always)* | `mom_axial_stepover_distance` |
| **PLANAR_DEBURRING** | *(always)* | `NO DATA` |
| **MILL_CONTROL** | *(always)* | `NO DATA` |
| **FLOOR_FACING** | *(always)* | `mom_depth_per_cut` |
| **FACE_MILLING_MANUAL** | *(always)* | `mom_depth_per_cut` |

---

## 2. `mill_contour`

### 2a. Final Ae — `mill_contour`

| Operation Subtype | Condition | Final Ae |
|---|---|---|
| **CAVITY_MILL** | sot = 1, `mom_stepover_distance` exists, no source defined | `mom_stepover_distance` |
| **CAVITY_MILL** | sot = 1, `mom_stepover_distance` exists, source defined, tool_diameter ≠ 0 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **CAVITY_MILL** | sot = 2 | `NO DATA` |
| **CAVITY_MILL** | sot = 3, region ∈ {4,5,7} | max of `tool_dep_values(i)` (src=0: raw; src=4: `value/100 × Ø`) |
| **CAVITY_MILL** | sot = 3, region ∈ {1,2,3} | `stepover_var_1` |
| **CAVITY_MILL** | sot = 4 | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **REST_MILLING** | *(same logic as CAVITY_MILL)* | *(see CAVITY_MILL rows above)* |
| **ADAPTIVE_MILLING** | sot = 1, `mom_stepover_distance` exists, no source defined | `mom_stepover_distance` |
| **ADAPTIVE_MILLING** | sot = 1, `mom_stepover_distance` exists, sds = 4, tool_diameter ≠ 0 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **3D_ADAPTIVE_ROUGHING** | sot = 1, `path_stepover_2` exists, no source defined | `path_stepover_2` (direct) |
| **3D_ADAPTIVE_ROUGHING** | sot = 1, `path_stepover_2` exists, sds = 4, tool_diameter ≠ 0 | `path_stepover_2 / 100 × mom_tool_diameter` |
| **3D_ADAPTIVE_ROUGHING** | sot_early = 4, sds = 4 | `path_stepover_2 / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **PLUNGE_MILLING** | sot = 1, `mom_stepover_distance` exists, no source defined | `mom_stepover_distance` |
| **PLUNGE_MILLING** | sot = 1, `mom_stepover_distance` exists, source defined, tool_diameter ≠ 0 | `mom_stepover_distance / 100 × mom_tool_diameter` |
| **PLUNGE_MILLING** | sot = 4 | `mom_stepover_percent / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **QUICK_ROUGHING** | sot_early = 4, `stepover_percent_2` numeric | `stepover_percent_2 / 100 × (mom_tool_diameter − 2 × mom_tool_corner1_radius)` |
| **QUICK_ROUGHING** | sot_early = 1, sot = 1, `path_stepover_2` exists, no source | `path_stepover_2` (direct) |
| **QUICK_ROUGHING** | sot_early = 1, sot = 1, `path_stepover_2` exists, sds = 4, tool_diameter ≠ 0 | `path_stepover_2 / 100 × mom_tool_diameter` |
| **ZLEVEL_PROFILE_STEEP** | *(always)* | `mom_tool_diameter` |
| **ZLEVEL_UNDERCUT** | *(always)* | `mom_tool_diameter` |
| **FLOW_MILL_SINGLE** | *(always)* | `mom_tool_diameter` |
| **FLOWCUT_SINGLE** | *(always)* | `mom_tool_diameter` |
| **FIXED_AXIS_GUIDING_CURVES** | *(always)* | `mom_stepover_distance` |
| **AREA_MILL** | *(always)* | `mom_stepover_distance` |
| **FLOW_MILL_MULTIPLE** | *(always)* | `mom_stepover_distance` |
| **FLOWCUT_MULTIPLE** | *(always)* | `mom_stepover_distance` |
| **FLOW_MILL_REF_TOOL** | *(always)* | `mom_stepover_distance` |
| **FLOWCUT_REF_TOOL** | *(always)* | `mom_stepover_distance` |
| **CURVE_DRIVE** | *(always)* | `mom_stepover_distance` |
| **SOLID_PROFILE_3D** | *(always)* | `mom_wall_increment` |
| **PROFILE_3D** | *(always)* | `mom_wall_increment` |
| **STREAMLINE** | *(always)* | `NO DATA` |
| **CONTOUR_SURFACE_AREA** | *(always)* | max(`mom_stepover_scallop`, `mom_horizonal_limit`, `mom_vertical_limit`) |
| **3_AXIS_DEBURRING** | *(always)* | `mom_deburring_edge_depth` |

### 2b. Final Ap — `mill_contour`

| Operation Subtype | Condition | Final Ap |
|---|---|---|
| **CAVITY_MILL** | *(always)* | `mom_global_cut_depth` |
| **ADAPTIVE_MILLING** | *(always)* | `mom_global_cut_depth` |
| **REST_MILLING** | *(always)* | `mom_global_cut_depth` |
| **ZLEVEL_PROFILE_STEEP** | *(always)* | `mom_global_cut_depth` |
| **ZLEVEL_UNDERCUT** | *(always)* | `mom_global_cut_depth` |
| **3D_ADAPTIVE_ROUGHING** | *(always)* | `mom_cut_level_distance` |
| **QUICK_ROUGHING** | *(always)* | `mom_cut_level_distance` |
| **PLUNGE_MILLING** | *(always)* | `mom_tool_flute_length` |
| **FIXED_AXIS_GUIDING_CURVES** | *(always)* | `mom_stepover_distance` |
| **AREA_MILL** | *(always)* | `mom_stepover_distance` |
| **FLOW_MILL_MULTIPLE** | *(always)* | `mom_stepover_distance` |
| **FLOWCUT_MULTIPLE** | *(always)* | `mom_stepover_distance` |
| **FLOW_MILL_REF_TOOL** | *(always)* | `mom_stepover_distance` |
| **FLOWCUT_REF_TOOL** | *(always)* | `mom_stepover_distance` |
| **CURVE_DRIVE** | *(always)* | `mom_stepover_distance` |
| **SOLID_PROFILE_3D** | *(always)* | `mom_multi_depth_cut_increment` |
| **PROFILE_3D** | *(always)* | `mom_multi_depth_cut_increment` |
| **FLOW_MILL_SINGLE** | *(always)* | `NO DATA` |
| **FLOWCUT_SINGLE** | *(always)* | `NO DATA` |
| **STREAMLINE** | *(always)* | `NO DATA` |
| **CONTOUR_SURFACE_AREA** | *(always)* | max(`mom_stepover_scallop`, `mom_horizonal_limit`, `mom_vertical_limit`) |
| **3_AXIS_DEBURRING** | *(always)* | `mom_deburring_edge_depth` |

---

## 3. `mill_multi-axis`

### 3a. Final Ae — `mill_multi-axis`

| Operation Subtype | Condition | Final Ae |
|---|---|---|
| **MULTI_AXIS_ROUGHING** | *(always)* | `mom_stepover_distance` |
| **VARIABLE_AXIS_GUIDING_CURVES** | *(always)* | `mom_stepover_distance` |
| **CONTOUR_PROFILE** | *(always)* | `mom_wall_increment` |
| **VARIABLE_STREAMLINE** | *(always)* | `NO DATA` |
| **VARIABLE_CONTOUR** | *(always)* | max(`mom_stepover_scallop`, `mom_horizonal_limit`, `mom_vertical_limit`) — strict numeric check |
| **WALL_FINISH-BARREL_SWARF** | *(always)* | `mom_maximal_stepover_distance` |
| **ZLEVEL_5AXIS** | *(always)* | `mom_tool_diameter` |
| **5_AXIS_DEBURRING** | *(always)* | `mom_deburring_edge_depth` |

### 3b. Final Ap — `mill_multi-axis`

| Operation Subtype | Condition | Final Ap |
|---|---|---|
| **MULTI_AXIS_ROUGHING** | *(always)* | `mom_cut_level_distance` |
| **VARIABLE_AXIS_GUIDING_CURVES** | *(always)* | `mom_stepover_distance` |
| **CONTOUR_PROFILE** | *(always)* | `mom_multi_depth_cut_increment` |
| **VARIABLE_STREAMLINE** | *(always)* | `NO DATA` |
| **VARIABLE_CONTOUR** | *(always)* | max(`mom_stepover_scallop`, `mom_horizonal_limit`, `mom_vertical_limit`) — strict numeric check |
| **WALL_FINISH-BARREL_SWARF** | *(always)* | `mom_depth_per_cut` |
| **ZLEVEL_5AXIS** | *(always)* | `mom_global_cut_depth` |
| **5_AXIS_DEBURRING** | *(always)* | `mom_deburring_edge_depth` |

---

## 4. `hole_making`

### 4a. Final Ae — `hole_making`

| Operation Subtype | Condition | Final Ae |
|---|---|---|
| **SPOT_DRILLING** | *(always)* | `mom_tool_diameter` |
| **DRILLING** | *(always)* | `mom_tool_diameter` |
| **BORING_REAMING** | *(always)* | `mom_tool_diameter` |
| **TAPPING** | *(always)* | `mom_tool_diameter` |
| **DEEP_HOLE_DRILLING** | *(always)* | `mom_tool_diameter` |
| **HOLE_MILLING** | *(always)* | `mom_tool_diameter` |
| **THREAD_MILLING** | sot = 1, `mom_stepover_distance` ≠ 0 | `mom_stepover_distance` |
| **THREAD_MILLING** | sot = 1, `mom_stepover_distance` = 0 or missing | `mom_tool_pitch` |
| **THREAD_MILLING** | sot = 3 | `max_stepover_var_tool_dep` (max of all variable tool-dependent values) |
| **THREAD_MILLING** | sot = 8 | `stepover_var_1` (`mom_stepover_variable_max_min(1)`) |

### 4b. Final Ap — `hole_making`

| Operation Subtype | Condition | Final Ap |
|---|---|---|
| **SPOT_DRILLING** | *(always)* | `mom_tool_flute_length` |
| **BORING_REAMING** | *(always)* | `mom_tool_flute_length` |
| **DEEP_HOLE_DRILLING** | *(always)* | `mom_tool_flute_length` |
| **THREAD_MILLING** | *(always)* | `mom_tool_flute_length` |
| **DRILLING** | Base | `mom_tool_flute_length` |
| **DRILLING** | Override: peck cycle active (`mom_cycle_step1` ≠ 0) | `mom_cycle_step1` |
| **DRILLING** | Override: depth increment source = 4 (% of Ø) | `mom_depth_increment_distance / 100 × mom_tool_diameter` |
| **TAPPING** | Base | `mom_tool_flute_length` |
| **TAPPING** | Override: peck cycle active (`mom_cycle_step1` ≠ 0) | `mom_cycle_step1` |
| **TAPPING** | Override: depth increment source = 4 (% of Ø) | `mom_depth_increment_distance / 100 × mom_tool_diameter` |
| **HOLE_MILLING** | `mom_vertical_pitch_type` = 0, no source or source = 0 | `mom_vertical_pitch_value` |
| **HOLE_MILLING** | `mom_vertical_pitch_type` = 0, source = 4 (% of Ø) | `mom_vertical_pitch_value / 100 × mom_tool_diameter` |
| **HOLE_MILLING** | `mom_vertical_pitch_type` ≠ 0 | `mom_tool_flute_length` |
