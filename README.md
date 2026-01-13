# Fusion360-Batch-Post (Enhanced Fork)

Fork of [TimPaterson/Fusion360-Batch-Post](https://github.com/TimPaterson/Fusion360-Batch-Post) with additional features.

## Enhancements

### Combine Setups (Minimize Tool Changes)

When using Fusion for Personal Use with split operations enabled, this feature combines multiple setups into a single output file, **reordering operations by tool number** to minimize tool changes.

**Enable:** Check "Combine setups (minimize tool changes)" in the Personal Use section.

**How it works:**
- Collects all operations from all setups
- Groups operations by tool number
- Outputs all operations for Tool 1, then Tool 2, etc.
- Suppresses redundant commands between same-tool operations
- Output filename: `<FirstSetupName>-COMBINED.nc`

**Intelligent command suppression:**

| Scenario | M9 (Coolant Off) | G28/G53 (Return Home) | Spindle Start | Coolant On | Dwell |
|----------|------------------|----------------------|---------------|------------|-------|
| Real tool change | ✅ Output | ✅ Output | ✅ Output | ✅ Output | ✅ Output |
| Same tool, WCS changing | ❌ Suppressed | ✅ Output (safety) | ❌ Suppressed | ❌ Suppressed | ❌ Suppressed |
| Same tool, same WCS | ❌ Suppressed | ❌ Suppressed | ❌ Suppressed | ❌ Suppressed | ❌ Suppressed |

**WCS Handling:** Each setup in Fusion has its own WCS (G54, G55, etc.). The post processor automatically outputs the correct WCS code for each operation. When transitioning between setups with different WCS (but the same tool), the return-to-home move (G28/G53) is kept for safety while other redundant commands are suppressed.

**Note:** Your `toolChange` setting (e.g., `M9:G28 G91 Z0:G90` or `M9:G0 G53 Z0`) is automatically parsed - M9 is filtered out when not needed, but G28/G53 is kept for safety when changing WCS.

### Append Origin Location to Filename

Automatically appends the WCS origin location to output filenames based on the stock point setting.

**Enable:** Check "Append origin location to filename" in the Personal Use section.

**Examples:**
- Origin at back-right, top of stock → `-BR-TOP`
- Origin at front-left, bottom of stock → `-FL-BOT`
- Origin at center XY, top of stock → `-TOP`
- Origin at center XY, bottom of stock → `-BOT`

**Corner positions:** FL (Front-Left), FR (Front-Right), BL (Back-Left), BR (Back-Right)

If the origin uses a custom location (not a stock corner), nothing is appended.

### Personal Use Warning Suppression

Automatically removes the repeated 4-line "When using Fusion for Personal Use..." warning comments from output files.

## Original Documentation

See the [upstream repository](https://github.com/TimPaterson/Fusion360-Batch-Post) for full documentation on base features.
